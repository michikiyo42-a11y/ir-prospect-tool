"""
右脳設計 IR自動リスト生成ツール
DX投資済み・未活用企業を有価証券報告書から自動特定し、Excelで出力する
"""

import io
import re
import time
import zipfile
from datetime import datetime, timedelta
from typing import List, Optional
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from company_list import COMPANIES_BY_INDUSTRY

# ── 定数 ──────────────────────────────────────
EDINET_BASE = "https://disclosure.edinet-fsa.go.jp/api/v2"

DX_KEYWORDS = [
    "DX推進", "デジタルトランスフォーメーション", "デジタル変革", "デジタル化",
    "AI活用", "人工知能", "機械学習", "生成AI", "ChatGPT",
    "デジタル人材", "DX人材", "AI人材", "デジタルスキル",
    "リスキリング", "アップスキリング", "クラウド", "データ活用",
]

PROBLEM_KEYWORDS = [
    "課題", "取り組み中", "推進中", "整備中", "検討中",
    "人材不足", "スキル不足", "育成が必要", "強化が必要",
    "遅れ", "困難", "リスク", "十分ではない", "今後", "目指",
]

EXEC_TITLES = [
    "最高デジタル責任者", "CDO", "最高情報責任者", "CIO",
    "最高人事責任者", "CHRO", "最高技術責任者", "CTO",
    "DX推進担当", "デジタル変革推進", "デジタル推進担当",
    "人事部長", "人事本部長", "人材開発部長", "人材育成担当",
    "組織開発部長", "組織開発担当", "タレントマネジメント",
    "研修部長", "研修担当", "HRBP", "HRビジネスパートナー",
    "DX推進部長", "DX推進室長",
]

COLUMNS = [
    ("企業名",                    18),
    ("業種",                      15),
    ("提出日",                    12),
    ("DXキーワード数",            14),
    ("課題キーワード数",          14),
    ("DX投資の記述（抜粋）",      55),
    ("課題・未活用の記述（抜粋）", 55),
    ("アプローチ対象者",          32),
    ("LinkedIn検索URL",           35),
    ("スコア",                    10),
]

SCORE_COLORS = {5: "C00000", 4: "FF6600", 3: "FFD700", 2: "92D050", 1: "BFBFBF"}


# ── EDINET ────────────────────────────────────

def find_annual_report(company_name: str, log_fn) -> Optional[dict]:
    today = datetime.now()

    def weekdays(start: int, end: int):
        for d in range(start, end):
            dt = today - timedelta(days=d)
            if dt.weekday() < 5:
                yield dt.strftime("%Y-%m-%d")

    ranges = (
        list(weekdays(0, 30)) +
        list(weekdays(300, 430)) +
        list(weekdays(30, 300)) +
        list(weekdays(430, 500))
    )

    for date_str in ranges:
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
                    log_fn(f"発見: {doc['filerName']} ({date_str})")
                    return doc
        except Exception:
            pass
        time.sleep(0.15)
    return None


def download_text(doc_id: str) -> str:
    try:
        resp = requests.get(
            f"{EDINET_BASE}/documents/{doc_id}",
            params={"type": 5},
            timeout=60,
        )
        resp.raise_for_status()
        texts = []
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            htm_files = sorted(
                [f for f in zf.namelist() if f.lower().endswith((".htm", ".xhtml"))],
                key=lambda f: zf.getinfo(f).file_size,
                reverse=True,
            )
            for fname in htm_files[:5]:
                try:
                    soup = BeautifulSoup(zf.read(fname), "html.parser")
                    for tag in soup(["script", "style"]):
                        tag.decompose()
                    text = "\n".join(
                        l for l in soup.get_text(separator="\n").splitlines() if l.strip()
                    )
                    if len(text) > 300:
                        texts.append(text)
                except Exception:
                    continue
        return "\n".join(texts)[:60000]
    except Exception:
        return ""


# ── テキスト分析 ──────────────────────────────

def excerpts(text: str, keywords: List[str], window: int = 120) -> List[str]:
    results = []
    for kw in keywords:
        for m in re.finditer(re.escape(kw), text):
            s = max(0, m.start() - window)
            e = min(len(text), m.end() + window)
            snippet = text[s:e].replace("\n", " ").strip()
            if snippet not in results:
                results.append(snippet)
    return results[:3]


def extract_executives(text: str, company: str) -> List[dict]:
    JP_NAME = r'[一-鿿]{1,4}[　\s]?[一-鿿]{1,4}'
    found, seen = [], set()
    for title in EXEC_TITLES:
        for m in re.finditer(re.escape(title), text):
            ctx = text[max(0, m.start()-80): min(len(text), m.end()+80)]
            for nm in re.finditer(JP_NAME, ctx):
                name = nm.group().replace("　", "").replace(" ", "").strip()
                if len(name) >= 3 and name not in seen:
                    seen.add(name)
                    q = quote(f"{name} {company}")
                    found.append({
                        "name": name,
                        "title": title,
                        "url": f"https://www.linkedin.com/search/results/people/?keywords={q}",
                    })
    return found[:5]


def analyze(company_name: str, industry: str, text: str, doc_info: Optional[dict]) -> dict:
    dx_count   = sum(text.count(k) for k in DX_KEYWORDS)
    prob_count = sum(text.count(k) for k in PROBLEM_KEYWORDS)
    filer      = doc_info.get("filerName", company_name) if doc_info else company_name
    execs      = extract_executives(text, filer) if text else []

    exec_names = "\n".join(f"{e['name']}（{e['title']}）" for e in execs) or "記載なし"
    li_urls    = "\n".join(e["url"] for e in execs) or \
                 f"https://www.linkedin.com/search/results/people/?keywords={quote('人事部長 ' + filer)}"

    score = min(5, max(1, int((dx_count * 0.3 + prob_count * 0.5) / 5))) if dx_count else 1

    return {
        "企業名":                    filer,
        "業種":                      industry,
        "提出日":                    (doc_info.get("submitDateTime", "")[:10] if doc_info else "-"),
        "DXキーワード数":            dx_count,
        "課題キーワード数":          prob_count,
        "DX投資の記述（抜粋）":      "\n\n".join(excerpts(text, DX_KEYWORDS)) or "記述なし",
        "課題・未活用の記述（抜粋）": "\n\n".join(excerpts(text, PROBLEM_KEYWORDS)) or "記述なし",
        "アプローチ対象者":          exec_names,
        "LinkedIn検索URL":           li_urls,
        "スコア":                    score,
    }


# ── Excel 生成 ────────────────────────────────

def build_excel(results: List[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "IRアプローチリスト"

    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_font = Font(color="FFFFFF", bold=True, size=11)
    for col, (name, width) in enumerate(COLUMNS, 1):
        c = ws.cell(row=1, column=col, value=name)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 28

    score_col = next(i+1 for i, (n, _) in enumerate(COLUMNS) if n == "スコア")
    li_col    = next(i+1 for i, (n, _) in enumerate(COLUMNS) if n == "LinkedIn検索URL")

    for row, r in enumerate(sorted(results, key=lambda x: x.get("スコア", 0), reverse=True), 2):
        score = r.get("スコア", 1)
        for col, (key, _) in enumerate(COLUMNS, 1):
            c = ws.cell(row=row, column=col, value=r.get(key, ""))
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if col == score_col:
                color = SCORE_COLORS.get(score, "BFBFBF")
                c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                c.font = Font(bold=True, color="FFFFFF" if score >= 4 else "000000", size=13)
                c.alignment = Alignment(horizontal="center", vertical="center")
            if col == li_col:
                c.font = Font(color="0563C1", underline="single")
        ws.row_dimensions[row].height = 120

    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── 共通：企業リストを処理 ─────────────────────

def run_analysis(companies: List[str], industry: str):
    results = []
    progress = st.progress(0)
    status_box = st.empty()
    log_box    = st.empty()

    for i, company in enumerate(companies):
        status_box.info(f"**[{i+1}/{len(companies)}] {company}** を処理中...")

        def log(msg): log_box.caption(msg)

        doc_info = find_annual_report(company, log)
        text     = download_text(doc_info["docID"]) if doc_info else ""
        if not doc_info:
            log("EDINETで見つかりませんでした")

        result = analyze(company, industry, text, doc_info)
        results.append(result)
        progress.progress((i + 1) / len(companies))
        time.sleep(0.5)

    status_box.success(f"✅ {len(companies)}社の分析が完了しました！")
    log_box.empty()
    return results


def show_results(results: List[dict]):
    st.subheader("分析結果")
    display_cols = ["企業名", "業種", "DXキーワード数", "課題キーワード数", "アプローチ対象者", "スコア"]
    df = pd.DataFrame(results)[display_cols].sort_values("スコア", ascending=False)
    st.dataframe(df, use_container_width=True)

    high = [r for r in results if r.get("スコア", 0) >= 4]
    if high:
        st.success(f"🎯 最優先アプローチ対象：{len(high)}社")

    excel_bytes = build_excel(results)
    filename = f"ir_prospect_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        label="📥 Excelをダウンロード",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


# ── Streamlit UI ──────────────────────────────

st.set_page_config(
    page_title="IRリスト自動生成 | 右脳設計",
    page_icon="📊",
    layout="wide",
)

st.title("📊 IRリスト自動生成ツール")
st.caption("有価証券報告書からDX投資済み・未活用企業を自動特定し、LinkedIn付きExcelを出力します")

with st.sidebar:
    st.header("使い方")
    st.markdown("""
**スコアの見方**
- 🔴 5点：最優先アプローチ
- 🟠 4点：優先
- 🟡 3点：検討
- 🟢 2点：様子見
- ⚪ 1点：優先度低

**目安時間**
- 1社あたり 1〜3分
""")

tab1, tab2 = st.tabs(["🔍 自動発掘モード", "📝 企業名を直接入力"])

# ────────────────────────────────
# Tab1: 自動発掘モード
# ────────────────────────────────
with tab1:
    st.markdown("### 業種を選ぶだけで自動的に候補企業を分析します")

    col1, col2 = st.columns([1, 1])

    with col1:
        industry = st.selectbox(
            "対象業種",
            list(COMPANIES_BY_INDUSTRY.keys()),
        )

        companies_in_industry = COMPANIES_BY_INDUSTRY[industry]
        max_companies = st.slider(
            "分析する企業数（上限）",
            min_value=3,
            max_value=len(companies_in_industry),
            value=min(10, len(companies_in_industry)),
        )
        selected_companies = companies_in_industry[:max_companies]

        st.info(f"推定時間: 約 {max_companies * 2}〜{max_companies * 3} 分")

    with col2:
        st.markdown(f"#### {industry}の分析対象企業")
        for c in selected_companies:
            st.write(f"・{c}")

    st.divider()

    if st.button("🚀 自動発掘スタート", type="primary", key="auto"):
        results = run_analysis(selected_companies, industry)
        show_results(results)

# ────────────────────────────────
# Tab2: 手動入力モード
# ────────────────────────────────
with tab2:
    st.markdown("### 企業名を自由に入力して分析します")

    col1, col2 = st.columns([1, 1])

    with col1:
        company_input = st.text_area(
            "企業名を入力（1行に1社・EDINETの正式名称推奨）",
            height=300,
            placeholder="富士通株式会社\n日本電気株式会社\n株式会社日立製作所",
        )
        industry_manual = st.text_input("業種（任意）", placeholder="製造業")

    with col2:
        st.markdown("#### 入力リスト")
        manual_companies = [c.strip() for c in company_input.splitlines() if c.strip()]
        if manual_companies:
            for c in manual_companies:
                st.write(f"・{c}")
            st.info(f"合計 {len(manual_companies)} 社　推定 {len(manual_companies)*2}〜{len(manual_companies)*3} 分")
        else:
            st.write("企業名を入力するとここに表示されます")

    st.divider()

    if st.button("🚀 分析スタート", type="primary", key="manual", disabled=not manual_companies):
        results = run_analysis(manual_companies, industry_manual or "未分類")
        show_results(results)
