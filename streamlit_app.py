"""
右脳設計 営業リスト自動生成ツール
商材を入力 → AIがマッチする企業を自動発掘 → 電話・担当者・LinkedInをExcel出力
"""

import io
import re
import time
import zipfile
from datetime import datetime, timedelta
from typing import List, Optional
from urllib.parse import quote, urljoin

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from company_list import COMPANIES_BY_INDUSTRY

HEADERS = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"}
EDINET_BASE = "https://disclosure.edinet-fsa.go.jp/api/v2"

DX_KEYWORDS = [
    "DX推進", "デジタルトランスフォーメーション", "デジタル変革", "デジタル化",
    "AI活用", "人工知能", "機械学習", "生成AI", "ChatGPT",
    "デジタル人材", "DX人材", "AI人材", "デジタルスキル",
    "リスキリング", "アップスキリング", "人材育成", "研修",
    "eラーニング", "学習", "教育", "スキルアップ",
]
PROBLEM_KEYWORDS = [
    "課題", "取り組み中", "推進中", "整備中", "検討中",
    "人材不足", "スキル不足", "育成が必要", "強化が必要",
    "遅れ", "困難", "十分ではない", "今後", "目指",
]
EXEC_TITLES = [
    "CDO", "CIO", "CHRO", "CTO", "最高デジタル責任者",
    "DX推進", "デジタル変革推進", "デジタル推進",
    "人事部長", "人事本部長", "人材開発部長", "人材育成",
    "組織開発部長", "組織開発", "研修部長", "研修担当",
    "DX推進部長", "DX推進室長", "HRBP",
]
PHONE_PATTERN = re.compile(r'0\d{1,4}[-－ー]\d{1,4}[-－ー]\d{4}')


# ── ページ取得 ────────────────────────────────

def fetch(url: str) -> Optional[BeautifulSoup]:
    try:
        r = requests.get(url, headers=HEADERS, timeout=12)
        r.raise_for_status()
        return BeautifulSoup(r.content, "html.parser")
    except Exception:
        return None


# ── 商材キーワード ────────────────────────────

def extract_product_keywords(text: str) -> List[str]:
    pool = DX_KEYWORDS + [
        "人事", "組織", "育成", "研修", "スキル", "学習", "eラーニング",
        "DX", "AI", "デジタル", "変革", "人材",
    ]
    found = [kw for kw in pool if kw in text]
    return found[:12] if found else ["DX推進", "人材育成", "AI活用", "リスキリング"]


# ── 電話番号 ──────────────────────────────────

def get_phone(company_name: str) -> str:
    try:
        q = quote(f"{company_name} 代表電話番号")
        r = requests.get(
            f"https://www.google.com/search?q={q}",
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=10,
        )
        soup = BeautifulSoup(r.content, "html.parser")
        phones = PHONE_PATTERN.findall(soup.get_text())
        if phones:
            return phones[0]
    except Exception:
        pass
    return "要確認"


# ── EDINET ───────────────────────────────────

def find_ir_doc(company_name: str) -> Optional[dict]:
    today = datetime.now()

    def weekdays(start: int, end: int):
        for d in range(start, end):
            dt = today - timedelta(days=d)
            if dt.weekday() < 5:
                yield dt.strftime("%Y-%m-%d")

    for date_str in (
        list(weekdays(0, 30))
        + list(weekdays(300, 430))
        + list(weekdays(30, 300))
        + list(weekdays(430, 500))
    ):
        try:
            resp = requests.get(
                f"{EDINET_BASE}/documents.json",
                params={"date": date_str, "type": 2},
                timeout=10,
            )
            if resp.status_code != 200:
                time.sleep(0.2)
                continue
            for doc in resp.json().get("results", []):
                if (company_name in doc.get("filerName", "")
                        and doc.get("docTypeCode") == "120"):
                    return doc
        except Exception:
            pass
        time.sleep(0.15)
    return None


def download_ir_text(doc_id: str) -> str:
    try:
        resp = requests.get(
            f"{EDINET_BASE}/documents/{doc_id}",
            params={"type": 5}, timeout=60,
        )
        resp.raise_for_status()
        texts = []
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            for fname in sorted(
                [f for f in zf.namelist() if f.lower().endswith((".htm", ".xhtml"))],
                key=lambda f: zf.getinfo(f).file_size, reverse=True,
            )[:5]:
                try:
                    soup = BeautifulSoup(zf.read(fname), "html.parser")
                    for tag in soup(["script", "style"]):
                        tag.decompose()
                    t = "\n".join(
                        l for l in soup.get_text(separator="\n").splitlines() if l.strip()
                    )
                    if len(t) > 300:
                        texts.append(t)
                except Exception:
                    continue
        return "\n".join(texts)[:60000]
    except Exception:
        return ""


def score_text(text: str, product_keywords: List[str]) -> int:
    kw = sum(text.count(k) for k in product_keywords)
    pb = sum(text.count(k) for k in PROBLEM_KEYWORDS)
    raw = kw * 0.3 + pb * 0.5
    return min(5, max(1, int(raw / 5))) if kw else 1


def extract_person_ir(text: str) -> dict:
    tail_name = re.compile(r'([一-鿿]{2,3}[一-鿿]{1,4})\s*$')
    for title in EXEC_TITLES:
        idx = text.find(title)
        if idx < 0:
            continue
        ctx = text[max(0, idx - 80): min(len(text), idx + 120)]
        if 'さん' in ctx:
            for seg in ctx.split('さん')[:-1]:
                m = tail_name.search(seg.strip())
                if m and len(m.group(1)) >= 3:
                    return {"担当者名": m.group(1), "部署・役職": title}
        m2 = re.search(r'[一-鿿]{2,4}', ctx[len(title):len(title)+20])
        if m2 and len(m2.group()) >= 3:
            return {"担当者名": m2.group(), "部署・役職": title}
    return {"担当者名": "（IR記載なし）", "部署・役職": "DX推進部・人事部（推定）"}


# ── 導入事例スクレイピング ─────────────────────

def get_article_urls(index_url: str) -> List[str]:
    soup = fetch(index_url)
    if not soup:
        return []
    seen, urls = set(), []
    for a in soup.find_all("a", href=True):
        full = urljoin(index_url, a["href"])
        if (full not in seen and full != index_url
                and full.endswith(".html")
                and index_url.split("/")[2] in full):
            seen.add(full)
            urls.append(full)
    return urls


def scrape_case_article(url: str) -> List[dict]:
    soup = fetch(url)
    if not soup:
        return []
    for tag in soup(["script", "style", "nav", "header", "footer"]):
        tag.decompose()
    text = "\n".join(l for l in soup.get_text(separator="\n").splitlines() if l.strip())

    company = ""
    for pat in [r'株式会社[^\s　、。「」（）]{2,15}',
                r'[^\s　、。「」（）]{2,15}株式会社',
                r'[^\s　、。「」（）]{2,15}ホールディングス']:
        m = re.search(pat, text[:500])
        if m and 4 <= len(m.group()) <= 20:
            company = m.group()
            break
    if not company:
        return []

    persons = []
    seen_names: set = set()
    photo_pat = re.compile(r'[（(][^）)]*?(左|右|中央|番目)[^）)]*?[）)]')
    head_bracket = re.compile(r'^[（(][^）)]+[）)]\s*')
    tail_name = re.compile(r'([一-鿿]{2,3}[一-鿿]{1,4})\s*$')

    for line in text.splitlines():
        if 'さん' not in line:
            continue
        for seg in line.split('さん')[:-1]:
            clean = photo_pat.sub('', seg).strip()
            clean = head_bracket.sub('', clean).strip()
            m = tail_name.search(clean)
            if not m:
                continue
            name = m.group(1).strip()
            if name in seen_names or len(name) < 3:
                continue
            before = clean[:m.start()].strip().rstrip('　').rstrip()
            parts = [p.strip() for p in re.split(r'[　\s]+', before) if p.strip()]
            seen_names.add(name)
            persons.append({"name": name, "dept_role": '　'.join(parts)})

    if not persons:
        return []

    phone = get_phone(company)
    rows = []
    for p in persons[:6]:
        q = quote(f"{p['name']} {company}")
        rows.append({
            "会社名":        company,
            "電話番号":      phone,
            "部署・役職":    p["dept_role"],
            "担当者名":      p["name"],
            "LinkedIn URL":  f"https://www.linkedin.com/search/results/people/?keywords={q}",
            "マッチスコア":  3,
            "ソース":        "導入事例",
        })
    return rows


# ── Excel ────────────────────────────────────

COLUMNS = [
    ("会社名",           22),
    ("電話番号",         18),
    ("部署・役職",       36),
    ("担当者名",         14),
    ("LinkedIn URL",     40),
    ("マッチスコア",     12),
    ("ソース",           12),
]
SCORE_COLORS = {5: "C00000", 4: "FF6600", 3: "FFD700", 2: "92D050", 1: "BFBFBF"}
SCORE_FONT   = {5: "FFFFFF", 4: "FFFFFF", 3: "000000", 2: "000000", 1: "000000"}


def build_excel(rows: List[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "営業リスト"
    hf = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    for col, (name, width) in enumerate(COLUMNS, 1):
        c = ws.cell(row=1, column=col, value=name)
        c.fill = hf
        c.font = Font(color="FFFFFF", bold=True, size=11)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 28

    score_col = next(i+1 for i, (n, _) in enumerate(COLUMNS) if n == "マッチスコア")
    li_col    = next(i+1 for i, (n, _) in enumerate(COLUMNS) if n == "LinkedIn URL")

    for row, r in enumerate(
        sorted(rows, key=lambda x: x.get("マッチスコア", 0), reverse=True), 2
    ):
        score = r.get("マッチスコア", 1)
        for col, (key, _) in enumerate(COLUMNS, 1):
            c = ws.cell(row=row, column=col, value=r.get(key, ""))
            c.alignment = Alignment(vertical="center", wrap_text=True)
            if col == score_col:
                color = SCORE_COLORS.get(score, "BFBFBF")
                c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                c.font = Font(bold=True, color=SCORE_FONT.get(score, "000000"), size=13)
                c.alignment = Alignment(horizontal="center", vertical="center")
            if col == li_col:
                c.font = Font(color="0563C1", underline="single")
        ws.row_dimensions[row].height = 45

    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Streamlit UI ──────────────────────────────

st.set_page_config(
    page_title="営業リスト自動生成 | 右脳設計",
    page_icon="📋",
    layout="wide",
)

st.title("📋 営業リスト自動生成ツール")
st.caption("商材を入力するだけで、刺さりそうな企業と担当者を自動で発掘します")

# ── サイドバー ────────────────────────────────
with st.sidebar:
    st.header("① 商材を入力する")
    product_text = st.text_area(
        "商材・サービスの説明",
        value=(
            "AI・DX人材の育成を支援する研修ツール。\n"
            "大手企業の人事部・DX推進部向け。\n"
            "リスキリングや生成AI活用研修が中心。"
        ),
        height=160,
    )
    product_keywords = extract_product_keywords(product_text)
    st.markdown("**探すキーワード**")
    st.write("　".join(product_keywords))

    st.divider()
    st.markdown("""
**スコアの見方**
🔴 5 最優先 / 🟠 4 優先
🟡 3 検討 / 🟢 2 様子見
""")

# ── タブ ─────────────────────────────────────
tab1, tab2 = st.tabs(["🤖 企業を自動発掘（IR分析）", "📰 導入事例から担当者を探す"])

# ─────────────────────────────────────────────
# Tab1: 自動発掘
# ─────────────────────────────────────────────
with tab1:
    st.markdown("### 業種を選ぶだけで、商材にマッチする企業を自動でピックアップします")

    industries = list(COMPANIES_BY_INDUSTRY.keys())
    selected = st.multiselect(
        "ターゲット業種（複数選択可）",
        industries,
        default=["製造業", "情報・通信業", "サービス業"],
    )

    target_companies = []
    for ind in selected:
        target_companies.extend(COMPANIES_BY_INDUSTRY.get(ind, []))

    col1, col2, col3 = st.columns(3)
    col1.metric("スキャン対象", f"{len(target_companies)}社")
    max_companies = col2.number_input("上限件数", min_value=3, max_value=50, value=10, step=1)
    col3.metric("推定時間", f"{max_companies * 4}〜{max_companies * 6}分")

    if selected:
        with st.expander(f"対象企業一覧（{len(target_companies)}社）"):
            cols = st.columns(3)
            for i, c in enumerate(target_companies):
                cols[i % 3].write(f"・{c}")

    run_btn = st.button(
        "🚀 発掘スタート",
        type="primary",
        disabled=not selected or not product_keywords,
    )

    if run_btn:
        companies_to_scan = target_companies[:max_companies]
        all_rows: List[dict] = []
        progress = st.progress(0)
        status = st.empty()
        log_area = st.empty()

        for i, company in enumerate(companies_to_scan):
            status.info(f"**[{i+1}/{len(companies_to_scan)}]** {company} を分析中...")

            log_area.caption(f"{company}：IR文書を検索中...")
            doc = find_ir_doc(company)

            if not doc:
                log_area.caption(f"{company}：IR文書が見つかりませんでした")
                progress.progress((i+1)/len(companies_to_scan))
                time.sleep(0.3)
                continue

            filer = doc.get("filerName", company)
            log_area.caption(f"{filer}：テキスト解析中...")
            text = download_ir_text(doc["docID"])
            score = score_text(text, product_keywords)
            person = extract_person_ir(text)

            log_area.caption(f"{filer}：電話番号を検索中...")
            phone = get_phone(filer)

            q = quote(f"{person['担当者名']} {filer}")
            all_rows.append({
                "会社名":        filer,
                "電話番号":      phone,
                "部署・役職":    person["部署・役職"],
                "担当者名":      person["担当者名"],
                "LinkedIn URL":  f"https://www.linkedin.com/search/results/people/?keywords={q}",
                "マッチスコア":  score,
                "ソース":        "IR文書",
            })

            progress.progress((i+1)/len(companies_to_scan))
            time.sleep(0.5)

        log_area.empty()
        status.success(f"✅ 完了！{len(all_rows)}社を発掘しました")

        if all_rows:
            df = pd.DataFrame(all_rows)[
                ["会社名", "電話番号", "部署・役職", "担当者名", "マッチスコア"]
            ].sort_values("マッチスコア", ascending=False)
            st.dataframe(df, use_container_width=True, hide_index=True)

            excel = build_excel(all_rows)
            st.download_button(
                "📥 Excelダウンロード（LinkedIn URL付き）",
                data=excel,
                file_name=f"prospect_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        else:
            st.warning("IR文書が取得できた企業がありませんでした。業種や対象件数を変えてお試しください。")

# ─────────────────────────────────────────────
# Tab2: 導入事例
# ─────────────────────────────────────────────
with tab2:
    st.markdown("### 導入事例サイトから担当者名・部署を自動抽出します")
    st.caption("UdemyやAllyなどのDXツール導入事例ページ → 担当者名・部署・電話・LinkedInをまとめてExcel出力")

    case_urls_input = st.text_area(
        "導入事例の一覧ページURL（1行に1URL）",
        height=120,
        placeholder=(
            "https://www.benesse.co.jp/udemy/business/case/\n"
            "https://..."
        ),
    )
    max_articles = st.slider("1サイトあたりの上限記事数", 5, 50, 20)

    if st.button("🚀 担当者を抽出する", type="primary"):
        case_urls = [u.strip() for u in case_urls_input.splitlines() if u.strip()]
        if not case_urls:
            st.warning("URLを入力してください")
        else:
            all_rows = []
            progress = st.progress(0)
            status = st.empty()

            for index_url in case_urls:
                status.info(f"一覧ページを取得中: {index_url[:60]}...")
                article_urls = get_article_urls(index_url)[:max_articles]

                for i, url in enumerate(article_urls):
                    status.info(f"[{i+1}/{len(article_urls)}] 記事を解析中...")
                    rows = scrape_case_article(url)
                    all_rows.extend(rows)
                    progress.progress((i+1)/len(article_urls))
                    time.sleep(1)

            status.success(f"✅ {len(all_rows)}名の担当者を発掘しました")

            if all_rows:
                df = pd.DataFrame(all_rows)[
                    ["会社名", "電話番号", "部署・役職", "担当者名"]
                ]
                st.dataframe(df, use_container_width=True, hide_index=True)

                excel = build_excel(all_rows)
                st.download_button(
                    "📥 Excelダウンロード（LinkedIn URL付き）",
                    data=excel,
                    file_name=f"case_list_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
