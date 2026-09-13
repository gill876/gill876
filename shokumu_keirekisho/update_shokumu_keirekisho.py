#!/usr/bin/env python3
"""Refreshes the 職務経歴書 .docx: live "as of" date field, GA technologies
content (replaced with OWNR-era work), and interpolated experience-years
for a specific subset of skills.

Word has no native equivalent of Excel's DATEDIF, so date-diff math can't
live-recalculate inside the document without VBA macros (which would force
a macro-enabled .docm — not appropriate for a resume sent to recruiters,
since many mail/security systems flag or strip .docm attachments). Instead:
  - The top "as of" date becomes a real Word DATE field (updates on
    F9 / print / "update fields on open", same as any Word date field).
  - The GA tenure and the handful of dynamic skill-years are recomputed by
    this script from a fixed baseline (2026-03-23, the document's last
    known-accurate snapshot) each time you run it. Re-run whenever you
    want fresh numbers.

Only these skills are treated as dynamic (per explicit request); every
other skill's 経験年数 is left completely untouched (same XML, not just
same displayed value):
  実装, テスト, 運用保守, macOS, Ruby, Ruby on Rails
"""
import copy
import re
from datetime import date

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.shared import Pt
from dateutil.relativedelta import relativedelta

SRC = "職務経歴書シーブライトカーギルドゥジョン_original_backup.docx"
DST = "職務経歴書シーブライトカーギルドゥジョン.docx"

BASELINE = date(2026, 3, 23)  # the document's last manually-verified snapshot date
TODAY = date.today()

FONT_NAME = "ＭＳ Ｐ明朝"
FONT_SIZE = Pt(9)


# ---------------------------------------------------------------- duration ----
def parse_duration(s):
    y = re.search(r"(\d+)年", s)
    m = re.search(r"(\d+)ヶ月", s)
    years = int(y.group(1)) if y else 0
    months = int(m.group(1)) if m else 0
    return years * 12 + months


def format_duration(total_months):
    y, m = divmod(total_months, 12)
    if y and m:
        return f"{y}年{m}ヶ月"
    if y:
        return f"{y}年"
    return f"{m}ヶ月"


def refresh_duration(duration_str):
    months = parse_duration(duration_str)
    anchor = BASELINE - relativedelta(months=months)
    rd = relativedelta(TODAY, anchor)
    return format_duration(rd.years * 12 + rd.months)


def tenure_since(start_year, start_month):
    rd = relativedelta(TODAY, date(start_year, start_month, 1))
    return format_duration(rd.years * 12 + rd.months)


# ------------------------------------------------------------------- font ----
def set_run_font(run, name=FONT_NAME, size=FONT_SIZE):
    """Sets BOTH the Latin and East Asian font (python-docx's run.font.name
    only sets the Latin/ascii font; Word renders Japanese text using the
    separate w:eastAsia attribute, which falls back to the document default
    -- a heavier-looking font -- if left unset)."""
    run.font.size = size
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), name)


# ------------------------------------------------------------ word fields ----
def insert_date_field(paragraph, keep_first_run=True):
    runs = paragraph.runs
    for r in runs[1 if keep_first_run else 0:]:
        r._element.getparent().remove(r._element)

    def make_run():
        r = paragraph.add_run()
        set_run_font(r)
        return r

    r1 = make_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    r1._r.append(fld_begin)

    r2 = make_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' DATE \\@ "yyyy年M月d日" \\* MERGEFORMAT '
    r2._r.append(instr)

    r3 = make_run()
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r3._r.append(fld_sep)

    r4 = make_run()
    r4.text = TODAY.strftime("%Y年%-m月%-d日")

    r5 = make_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r5._r.append(fld_end)


# ------------------------------------------------------------- cell setup ----
def set_cell_blocks(cell, blocks):
    """blocks: list of (text, bold) tuples. text may contain \\n for line breaks.
    Clears the cell and writes one paragraph per block, with correct CJK font."""
    for p in list(cell.paragraphs):
        p._element.getparent().remove(p._element)

    for text, bold in blocks:
        p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        set_run_font(run)
        run.bold = bold


def remove_row(table, index):
    row = table.rows[index]
    row._element.getparent().remove(row._element)


def set_paragraph_value(paragraph, new_text):
    """Overwrite a paragraph's displayed value in-place: put the full new
    text in the first run (preserving its original font/size/bold exactly
    as authored) and blank out any remaining runs. Used only for the
    specific paragraphs we're intentionally making dynamic -- every other
    paragraph in the skills table is left completely untouched."""
    runs = paragraph.runs
    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""


def split_merged_paragraph(cell, index, first_text, second_text):
    """Repairs a pre-existing bug in the source document: this paragraph
    holds two values merged together with an embedded line break (e.g.
    '1年5ヶ月\\n3ヶ月'), one paragraph short of its sibling columns. Splits
    it into two proper paragraphs, restoring 1:1 alignment with the rest
    of the row."""
    p0 = cell.paragraphs[index]
    p0_elem = p0._p
    runs = p0.runs

    break_idx = next(i for i, r in enumerate(runs) if r.text == "\n")
    tail_runs = runs[break_idx:]  # the break run + everything after it
    tail_font_name = tail_runs[-1].font.name
    tail_font_size = tail_runs[-1].font.size

    for r in tail_runs:
        r._element.getparent().remove(r._element)

    new_p_elem = copy.deepcopy(p0_elem)
    for r in new_p_elem.findall(qn("w:r")):
        new_p_elem.remove(r)
    p0_elem.addnext(new_p_elem)

    new_p = Paragraph(new_p_elem, p0._parent)
    new_run = new_p.add_run(second_text)
    new_run.font.name = tail_font_name
    new_run.font.size = tail_font_size

    set_paragraph_value(p0, first_text)


# =============================================================== MAIN ========
d = docx.Document(SRC)

# ---- 1. live "as of" date field ----
insert_date_field(d.paragraphs[1])

# ---- 2. current company headcount (1705 -> 1916) ----
d.paragraphs[58].runs[3].text = "1916"

# ---- 3. GA technologies summary paragraphs ----
p7_runs = d.paragraphs[7].runs
p7_runs[0].text = "2024年、株式会社GA technologiesに入社。"
for r in p7_runs[1:]:
    r.text = ""
p8_runs = d.paragraphs[8].runs
p8_runs[0].text = (
    "Ruby on Railsを用いた社内Webアプリケーション開発を経て、現在は同社が提供する"
    "不動産投資プラットフォーム「OWNR」の本番・ステージング環境におけるエラートリアージ、"
    "OpenAPI仕様書の整備、および既存テストコードの保守性向上に従事しています。"
)
for r in p8_runs[1:]:
    r.text = ""

# ---- 4. dynamic skill-experience years (everything else stays untouched) ----
skills_table = d.tables[0]

# 担当業務: 実装, テスト, 運用保守 -- all three dynamic
row1_years = skills_table.rows[1].cells[2]
set_paragraph_value(row1_years.paragraphs[0], refresh_duration("3年"))       # 実装
set_paragraph_value(row1_years.paragraphs[1], refresh_duration("3年"))       # テスト
set_paragraph_value(row1_years.paragraphs[2], refresh_duration("1年5ヶ月"))  # 運用保守

# OS・開発環境: only macOS (index 0) dynamic; Linux/Windows untouched
row2_years = skills_table.rows[2].cells[2]
set_paragraph_value(row2_years.paragraphs[0], refresh_duration("4年"))       # macOS

# 言語: only Ruby (index 5) dynamic; everything else untouched
row3_years = skills_table.rows[3].cells[2]
set_paragraph_value(row3_years.paragraphs[5], refresh_duration("2年1ヶ月"))  # Ruby

# フレームワーク: repair the pre-existing merged-paragraph bug first (index 0
# holds AngularJS + React merged together), THEN update Ruby on Rails (which
# is at index 4 once the row is correctly split).
row4_years = skills_table.rows[4].cells[2]
split_merged_paragraph(row4_years, 0, "1年5ヶ月", "3ヶ月")  # AngularJS | React
set_paragraph_value(row4_years.paragraphs[4], refresh_duration("1年5ヶ月"))  # Ruby on Rails

# DB row: nothing requested as dynamic -- left entirely untouched

# ---- 5. GA technologies project entry: replace with OWNR-era content ----
ga_table = d.tables[1]

period_text = f"2024年\n11月\n|\n現在\n（{tenure_since(2024, 11)}）"
ga_period_cell = ga_table.rows[1].cells[0]
for p in list(ga_period_cell.paragraphs):
    p._element.getparent().remove(p._element)
p = ga_period_cell.add_paragraph()
run = p.add_run(period_text)
set_run_font(run)

# Content is split across the same two rows the ORIGINAL document used for
# this table (row 1 + row 2) rather than crammed into one giant row. A
# single ~900-character cell doesn't fit in whatever page space is left
# after the row above it, so Word/Pages pushes the *entire* row to the next
# page instead of splitting it -- stranding the table header alone. This is
# almost certainly why the original author split the content across two
# rows in the first place; collapsing it into one row reintroduced the
# exact problem that split was avoiding.
row1_blocks = [
    ("■（OWNRプラットフォームにおける本番運用・API整備）", True),
    (
        "【プロジェクト概要】\n"
        "GA technologiesが提供する不動産投資プラットフォーム「OWNR」（社内外14サービス・"
        "4開発チームと連携）における、本番・ステージング環境のエラートリアージおよび"
        "OpenAPI仕様書の整備",
        False,
    ),
    (
        "【担当フェーズ】\n"
        "・エラートリアージ（本番・ステージング環境）\n"
        "・API仕様書の作成・整合性検証\n"
        "・データ不整合の調査・是正\n"
        "・テストコードの保守性改善（RSpec/FactoryBot）",
        False,
    ),
    (
        "【業務内容】\n"
        "・Rollbar通知を起点としたエラートリアージ：\n"
        "発生したエラーがドメインサービス起因か、インフラ起因か、アプリケーションバグかを"
        "切り分け、CloudWatch、CloudTrail、ECSコンテナログを用いて原因調査を実施。"
        "データ起因の可能性がある場合はSnowflakeのミラー環境と突合し、ドメインサービス"
        "担当チームと連携して仮説を検証。\n\n"
        "・OpenAPI仕様書の作成・整合化：\n"
        "プロバイダー／コンシューマー間でAPI仕様書が二重管理されている課題を解消する"
        "全社的な取り組みの一環として、実装内容と突合しながらOWNRのAPI（全100以上の"
        "エンドポイントのうち30以上を担当）についてOpenAPI仕様書を作成・修正。"
        "iOS/Androidクライアントチームを含む関係チームとの調整を実施。",
        False,
    ),
]
set_cell_blocks(ga_table.rows[1].cells[1], row1_blocks)

row2_blocks = [
    (
        "・本番データの不整合対応：\n"
        "本番環境におけるデータ不整合を2件特定し、ドメインサービス担当チームと連携して"
        "原因究明・是正を実施。\n\n"
        "・テストコードの保守性向上：\n"
        "レガシーなRSpecのfixturesをFactoryBotへ移行し、可読性・再現性を向上。"
        "あわせて永続化を伴うcreate呼び出しをメモリ上で完結するbuild呼び出しに"
        "置き換え、CI/CDのテスト実行時間を短縮。",
        False,
    ),
    (
        "【実績・取り組み等】\n"
        "・エラートリアージのプロセスを通じて、ドメインサービス・インフラ・アプリケーション"
        "起因の切り分け精度を向上。\n"
        "・APIドキュメントの二重管理解消に貢献し、iOS/Androidを含む関係チームとの"
        "認識齟齬を削減。",
        False,
    ),
]
set_cell_blocks(ga_table.rows[2].cells[1], row2_blocks)

devenv_blocks = [
    (
        "【OS】\nmacOS\n\n"
        "【言語】\nRuby\nShell Script\nHTML\nCSS\nJavaScript\n\n"
        "【フレームワーク】\nRuby on Rails\n\n"
        "【DB】\nPostgreSQL\nSnowflake（参照）\n\n"
        "【ツール】\nRSpec\nFactoryBot\nDocker\nRollbar\nOpenAPI\nGit",
        False,
    ),
]
set_cell_blocks(ga_table.rows[1].cells[2], devenv_blocks)

role_blocks = [
    ("【役割】\nメンバー\n\n【プロジェクト規模】\nチーム：15名\n全体：1916名", False),
]
set_cell_blocks(ga_table.rows[1].cells[3], role_blocks)

# row 2's other columns stay empty, matching the original document's own
# convention for this continuation row.
for cell in (ga_table.rows[2].cells[0], ga_table.rows[2].cells[2], ga_table.rows[2].cells[3]):
    for p in list(cell.paragraphs):
        p._element.getparent().remove(p._element)
    p = cell.add_paragraph()
    run = p.add_run("")
    set_run_font(run)

# ---- 6. strip stale fixed row heights across every table ----
# The source document has several rows with an explicit minimum height
# (w:trHeight) baked in from whatever content used to be there. When a
# row's real content needs more room than fits in the space left on the
# current page, some renderers (observed in Pages) push the *entire* row
# to the next page rather than splitting it, stranding whatever precedes
# it (a table header, in both cases found so far). Clearing every explicit
# trHeight lets every row size to its actual content instead -- this is a
# presentation-only change, not a content edit, so it doesn't touch any of
# the "leave untouched" skill values.
for table in d.tables:
    for row in table.rows:
        trPr = row._tr.find(qn("w:trPr"))
        if trPr is not None:
            trHeight = trPr.find(qn("w:trHeight"))
            if trHeight is not None:
                trPr.remove(trHeight)

d.save(DST)
print("saved", DST)
