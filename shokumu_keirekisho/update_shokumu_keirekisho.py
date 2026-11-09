#!/usr/bin/env python3
"""Refreshes the 職務経歴書 .docx: live "as of" date field, GA technologies
content (split into the internal-platform period and the RENOSY backend
period),
研究実績 / 自己研鑽・個人開発 sections with hyperlinks, a rewritten 自己PR,
and interpolated experience-years for a specific subset of skills.

Word has no native equivalent of Excel's DATEDIF, so date-diff math can't
live-recalculate inside the document without VBA macros (which would force
a macro-enabled .docm — not appropriate for a resume sent to recruiters,
since many mail/security systems flag or strip .docm attachments). Instead:
  - The top "as of" date becomes a real Word DATE field (updates on
    F9 / print / "update fields on open", same as any Word date field).
  - The GA tenure and the handful of dynamic skill-years are recomputed by
    this script each time you run it: pre-existing skills from a fixed
    baseline (2026-03-23, the document's last known-accurate snapshot),
    skills added later from the month they were first used. Re-run
    whenever you want fresh numbers.

The skills table is rebuilt with one table row per skill (see step 4); only
these skills' 経験年数 are dynamic, and every other value is carried over
from the source unchanged (asserted at run time):
  実装, テスト, 運用保守, macOS, Ruby, Ruby on Rails             (from BASELINE)
  TypeScript, Next.js, FastAPI, Firestore, Cloud Run,
  GitHub Actions, OpenAPI, Sentry, Grafana                    (from first use)
"""
import copy
import re
from datetime import date

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.shared import Pt
from dateutil.relativedelta import relativedelta

SRC = "職務経歴書シーブライトカーギルドゥジョン_original_backup.docx"
DST = "職務経歴書シーブライトカーギルドゥジョン.docx"

BASELINE = date(2026, 3, 23)  # the document's last manually-verified snapshot date
TODAY = date.today()

# Moved from the internal platform to the team that runs RENOSY's backend. The
# team is internally still named after OWNR, the user-facing app's old name;
# the resume describes it as RENOSY's backend instead.
BACKEND_TEAM_START = (2026, 5)
HEADCOUNT = "1853"  # ga-tech.co.jp company outline, as of end of July 2026 (group total)

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


def months_since(start_year, start_month):
    rd = relativedelta(TODAY, date(start_year, start_month, 1))
    return max(1, rd.years * 12 + rd.months)


def tenure_since(start_year, start_month):
    return format_duration(months_since(start_year, start_month))


def tenure_between(start, end):
    """Months from `start` to `end` (both (year, month)), counted the way the
    Artuvic entry already is: 2021年3月〜2022年8月 is 1年5ヶ月."""
    rd = relativedelta(date(*end, 1), date(*start, 1))
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


def append_hyperlink_field(p_elem, url, text, rPr):
    """Appends a HYPERLINK field (begin / instr / separate / text / end),
    the same construction the document's existing GitHub link uses, so
    Word, Pages and LibreOffice's PDF export all treat it as a link."""
    for kind in ("begin", "instr", "separate", "text", "end"):
        r = OxmlElement("w:r")
        r.append(copy.deepcopy(rPr))
        if kind == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = f' HYPERLINK "{url}" '
        elif kind == "text":
            el = OxmlElement("w:t")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        r.append(el)
        p_elem.append(r)


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
    as authored) and blank out any remaining runs."""
    runs = paragraph.runs
    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""


# =============================================================== MAIN ========
d = docx.Document(SRC)
P = list(d.paragraphs)  # fixed references: later insertions don't shift these

# Templates for the new sections, copied before anything is modified.
HEADING_P = copy.deepcopy(P[69]._p)  # "■　取得資格等"
BLANK_P = copy.deepcopy(P[68]._p)
BODY_P = copy.deepcopy(P[79]._p)  # first 自己PR paragraph
BODY_RPR = copy.deepcopy(P[79].runs[0]._r.find(qn("w:rPr")))
for r in BODY_P.findall(qn("w:r")):
    BODY_P.remove(r)
LINK_RPR = copy.deepcopy(
    next(r for r in P[73].runs if r.text == "GitHub")._r.find(qn("w:rPr"))
)
assert P[69].text == "■　取得資格等" and P[73].text == "GitHub"

# ---- 1. live "as of" date field ----
insert_date_field(P[1])

# ---- 2. current company headcount ----
assert P[58].runs[3].text == "1705"
P[58].runs[3].text = HEADCOUNT

# ---- 3. GA technologies summary paragraphs ----
p7_runs = P[7].runs
p7_runs[0].text = "2024年、株式会社GA technologiesに入社。"
for r in p7_runs[1:]:
    r.text = ""
p8_runs = P[8].runs
p8_runs[0].text = (
    "入社後、Ruby on Railsを用いた社内向けWebアプリケーションの設計・開発を担当しました。"
    "2026年5月からは不動産投資サービス「RENOSY」のバックエンドを担当するチームに所属し、本番環境の"
    "エラートリアージ、Claude CodeとSnowflakeを活用した調査の効率化、OpenAPI仕様書の整備、"
    "テストコードの改善に従事しています。"
)
for r in p8_runs[1:]:
    r.text = ""

# ---- 4. skills table: rebuilt with one table row per skill ----
# The source packs each category into ONE row and lists its skills as stacked
# paragraphs in parallel columns, so a skill only lines up with its 経験年数
# and note while every column wraps identically. They don't: the notes use
# 1.3x line spacing and the other columns don't, and any wrapped line
# ("・Java Play Framework", a long note) shifts everything below it by a line.
# Pages happened to hide this; the LibreOffice-built PDF doesn't. One row per
# skill keeps each skill's cells together in any renderer. Inner borders and
# padding are removed so each category still reads as one block.
STUDENT = "学生時代に使用経験有"
LOOKUP = "調べながらであれば使用可能"
FLUENT = "通常使用に問題なし"
PERSONAL = "個人開発で使用"
SPAN = object()  # note continues the one above (vertically merged cell)

# The source's 担当業務 rows (実装・テスト・運用保守) are left out: the project
# tables already show those phases.
SKILLS = [
    ("OS・開発環境", [
        ("macOS", refresh_duration("4年"), ""),
        ("Linux", "8年", ""),
        ("Windows", "8年", ""),
    ]),
    ("言語", [
        ("PHP", "1年", STUDENT),
        ("Python", "5年", "研究（LLMパイプライン）・個人開発（FastAPI）で使用"),
        ("C言語", "2年", STUDENT),
        ("Java", "3年", LOOKUP),
        ("Ruby", refresh_duration("2年1ヶ月"), FLUENT),
        ("R言語", "3ヶ月", STUDENT),
        ("HTML", "7年", FLUENT),
        ("CSS", "7年", LOOKUP),
        ("JavaScript", "3年", LOOKUP),
        ("MD Markup", "2年", FLUENT),
        ("SQL", "2年", LOOKUP),
        ("TypeScript", tenure_since(2026, 10), PERSONAL),
    ]),
    ("フレームワーク", [
        ("AngularJS", "1年5ヶ月", LOOKUP),
        ("React", "3ヶ月", "学生時代に使用経験有、自己研鑽中"),
        ("HuggingFace", "3ヶ月", LOOKUP),
        ("Java Play Framework", "1年5ヶ月", LOOKUP),
        ("Ruby on Rails", refresh_duration("1年5ヶ月"), FLUENT),
        # the 1ヶ月 of self-study already listed, plus Kechi+ since 2026/10
        ("Next.js", format_duration(1 + months_since(2026, 10)), PERSONAL),
        ("Spring Boot", "2ヶ月", "自己研鑽中"),
        ("Flask", "2年", LOOKUP),
        ("Bootstrap", "2年", LOOKUP),
        ("Tailwind CSS", "3ヶ月", LOOKUP),
        ("FastAPI", tenure_since(2026, 9), PERSONAL),
    ]),
    ("DB", [
        ("MySQL", "2年", "DB移行経験ありテーブル設計可能/調べながらであれば作業可能/大学で使用"),
        ("PostgreSQL", "2年", SPAN),
        ("Firestore", tenure_since(2026, 9), PERSONAL),
    ]),
    ("クラウド・ツール", [
        ("Cloud Run", tenure_since(2026, 9), PERSONAL),
        ("GitHub Actions", tenure_since(2026, 9), PERSONAL),
        ("OpenAPI", tenure_since(*BACKEND_TEAM_START), "業務・個人開発で使用"),
        ("Sentry", tenure_since(2026, 10), PERSONAL),
        ("Grafana", tenure_since(2026, 10), PERSONAL),
        # a DB admin tool, so it sits with the tools rather than under DB
        ("phpMyAdmin", "1年5ヶ月", "前職でMySQLの管理に使用"),
    ]),
]

skills_table = d.tables[0]
old_rows = [row._tr for row in skills_table.rows[1:]]

# Guard against transcription slips: every skill in the source must appear
# above with the same years and note, unless it's one we deliberately change
# or drop.
CHANGED = {"実装", "テスト", "運用保守", "macOS", "Ruby", "Ruby on Rails", "Python", "Next.js",
           "phpMyAdmin"}
listed = {name: (years, note) for _, items in SKILLS for name, years, note in items}
for row in skills_table.rows[1:6]:
    names = [p.text.lstrip("・") for p in row.cells[1].paragraphs if p.text]
    years = [y for p in row.cells[2].paragraphs for y in p.text.split("\n") if y]
    notes = [p.text for p in row.cells[3].paragraphs if p.text]
    assert len(names) == len(years), names
    for i, name in enumerate(names):
        if name in CHANGED:
            continue
        assert listed[name][0] == years[i], (name, years[i])
        if len(notes) == len(names):
            assert listed[name][1] == notes[i], (name, notes[i])

LINE_P = copy.deepcopy(skills_table.rows[3].cells[3].paragraphs[0]._p)  # a note line
for r in LINE_P.findall(qn("w:r"))[1:]:
    LINE_P.remove(r)
# Tag Latin text as English (Japanese text stays ja-JP via w:eastAsia).
# Left as ja-JP, LibreOffice applies Japanese line breaking to Latin words
# too and splits them mid-word ("Framewor" / "k").
LINE_P.find(qn("w:r")).find(qn("w:rPr")).find(qn("w:lang")).set(qn("w:val"), "en-US")


def set_tc_edge(tcPr, tag, edge, value):
    """Sets one edge of a cell's tcBorders/tcMar: border edges to 'nil',
    margin edges to a width in twips."""
    el = tcPr.find(qn(tag)).find(qn(edge))
    if tag == "w:tcBorders":
        el.set(qn("w:val"), value)
    else:
        el.set(qn("w:w"), value)


def set_vmerge(tcPr, restart):
    vm = OxmlElement("w:vMerge")
    if restart:
        vm.set(qn("w:val"), "restart")
    tcPr.find(qn("w:tcW")).addnext(vm)  # schema order: tcW, gridSpan, vMerge, ...


anchor = skills_table.rows[0]._tr
for category, items in SKILLS:
    for i, (name, years, note) in enumerate(items):
        first, last = i == 0, i == len(items) - 1
        note_continues = not last and items[i + 1][2] is SPAN
        tr = copy.deepcopy(old_rows[0])
        for ci, tc in enumerate(tr.findall(qn("w:tc"))):
            for p in tc.findall(qn("w:p")):
                tc.remove(p)
            text = (category if first else "", f"・{name}", years, "" if note is SPAN else note)[ci]
            p = copy.deepcopy(LINE_P)
            p.find(qn("w:r")).find(qn("w:t")).text = text
            tc.append(p)

            tcPr = tc.tcPr
            if ci == 0:
                set_vmerge(tcPr, restart=first)
            elif ci == 3 and (note is SPAN or note_continues):
                set_vmerge(tcPr, restart=note is not SPAN)
            if not first:
                set_tc_edge(tcPr, "w:tcBorders", "w:top", "nil")
                set_tc_edge(tcPr, "w:tcMar", "w:top", "0")
            if not last:
                set_tc_edge(tcPr, "w:tcBorders", "w:bottom", "nil")
                set_tc_edge(tcPr, "w:tcMar", "w:bottom", "0")
        tr.get_or_add_trPr().insert(0, OxmlElement("w:cantSplit"))
        anchor.addnext(tr)
        anchor = tr

for tr in old_rows:
    tr.getparent().remove(tr)

# ---- 5. GA technologies project entries ----
# Two projects, each split across two table rows (content + continuation)
# the same way the ORIGINAL document split its single GA entry. A single
# ~900-character cell doesn't fit in whatever page space is left after the
# row above it, so Word/Pages pushes the *entire* row to the next page
# instead of splitting it -- stranding the table header alone. That's why
# the original author split the content across two rows; collapsing a
# project into one row reintroduces exactly that problem.
ga_table = d.tables[1]
backend_tr1 = copy.deepcopy(ga_table.rows[1]._tr)
backend_tr2 = copy.deepcopy(ga_table.rows[2]._tr)
ga_table.rows[2]._tr.addnext(backend_tr1)
backend_tr1.addnext(backend_tr2)


def set_period(cell, text):
    for p in list(cell.paragraphs):
        p._element.getparent().remove(p._element)
    p = cell.add_paragraph()
    run = p.add_run(text)
    set_run_font(run)


def clear_cell(cell):
    for p in list(cell.paragraphs):
        p._element.getparent().remove(p._element)
    p = cell.add_paragraph()
    run = p.add_run("")
    set_run_font(run)


internal_end = (BACKEND_TEAM_START[0], BACKEND_TEAM_START[1] - 1) if BACKEND_TEAM_START[1] > 1 else (BACKEND_TEAM_START[0] - 1, 12)

# -- (a) internal learning / work-management platform: rows 1-2 --
set_period(
    ga_table.rows[1].cells[0],
    f"2024年\n11月\n|\n{internal_end[0]}年\n{internal_end[1]}月\n"
    f"（{tenure_between((2024, 11), internal_end)}）",
)
set_cell_blocks(ga_table.rows[1].cells[1], [
    ("■（社内向け学習・業務管理プラットフォームの設計・開発（個人開発））", True),
    (
        "【プロジェクト概要】\n"
        "社内の学習進捗および業務工数を一括管理するWebアプリケーションのフルスタック開発",
        False,
    ),
    (
        "【担当フェーズ】\n"
        "・基本設計・詳細設計\n"
        "・DBモデリング\n"
        "・機能実装（API/Frontend）\n"
        "・テスト実装（RSpec/Mock）",
        False,
    ),
    (
        "【業務内容】\n"
        "・データモデリング：\n"
        "多対多のリレーションを設計し、ビジネスロジックに基づく厳格なバリデーションを実装。\n\n"
        "・認証・認可：\n"
        "セッションベースの認証機能と、ネストされたリソースに対する認可ロジックを実装。",
        False,
    ),
])
set_cell_blocks(ga_table.rows[2].cells[1], [
    (
        "・保守性の向上：\n"
        "Helperへのロジック切り出し（DRY化）とRuboCopによる静的解析で保守性を維持。\n\n"
        "・国際化（I18n）対応：\n"
        "日本語ロケールファイル（ja.yml）の整備とUIのローカライズ。",
        False,
    ),
    (
        "【実績・取り組み等】\n"
        "・RSpec/FactoryBotで300行以上のテストを実装し、境界値・エッジケースを網羅。\n"
        "・シェルスクリプトでDBのバックアップ・リストアを自動化し、開発環境の構築工数を削減。",
        False,
    ),
])
set_cell_blocks(ga_table.rows[1].cells[2], [(
    "【OS】\nmacOS\n\n"
    "【言語】\nRuby\nShell Script\nHTML\nCSS\nJavaScript\n\n"
    "【フレームワーク】\nRuby on Rails\n\n"
    "【DB】\nPostgreSQL\n\n"
    "【ツール】\nRSpec\nDocker\nRuboCop\nHerb\nBrakeman\nBundler-audit",
    False,
)])
set_cell_blocks(ga_table.rows[1].cells[3], [
    (f"【役割】\nメンバー\n\n【プロジェクト規模】\nチーム：1名\n全体：{HEADCOUNT}名", False),
])

# -- (b) RENOSY backend: rows 3-4 --
set_period(
    ga_table.rows[3].cells[0],
    f"{BACKEND_TEAM_START[0]}年\n{BACKEND_TEAM_START[1]}月\n|\n現在\n（{tenure_since(*BACKEND_TEAM_START)}）",
)
set_cell_blocks(ga_table.rows[3].cells[1], [
    ("■（RENOSYバックエンドの本番運用・API整備）", True),
    (
        "【プロジェクト概要】\n"
        "株式会社GA technologiesが提供する不動産投資サービス「RENOSY」の、ユーザー向けアプリ"
        "（iOS/Android・Web）を支えるバックエンドサーバー（社内外14サービス・5開発チームと連携）"
        "における、本番・ステージング環境のエラートリアージ、OpenAPI仕様書の整備、テストコードの改善",
        False,
    ),
    (
        "【担当フェーズ】\n"
        "・エラートリアージ（本番・ステージング環境）\n"
        "・データ不整合の調査・是正\n"
        "・API仕様書の作成・整合性検証\n"
        "・リリース時の動作確認\n"
        "・テストコードの保守性改善（RSpec/FactoryBot）",
        False,
    ),
    (
        "【業務内容】\n"
        "・Rollbar起点のエラートリアージ：\n"
        "エラーの原因がドメインサービス・インフラ・アプリケーションのどこにあるかを切り分け、"
        "CloudWatch（ECSのコンテナログ・メトリクス）、CloudTrail、Snowflakeのミラーを用いて調査。\n\n"
        "・エラー調査の効率化（自ら導入）：\n"
        "本番エラーの調査に1〜2日かかっていたため、チームで初めてSnowflake CLIをClaude Codeに"
        "連携し、調査時間を1時間以内に短縮。さらに調査手順をClaude Codeのスキルにまとめて社内の"
        "マーケットプレイスで公開し、チームの誰でも同じ手順で調査できるようにした。\n\n"
        "・本番データの不整合の特定・是正：\n"
        "調査の中で本番データの不整合を3件特定し、ドメインチームと連携して是正。うち1件は、"
        "誰も把握していなかった不整合だった。",
        False,
    ),
])
set_cell_blocks(ga_table.rows[4].cells[1], [
    (
        "・メモリ断片化の兆候への対応（自ら提案）：\n"
        "発生していた問題がメモリ断片化（フラグメンテーション）によるものかを、アプリのサイドカーで収集している"
        "Datadogのメモリメトリクスで確認。その兆候を踏まえ、原因を切り分けるためのツールの導入を自ら提案。\n\n"
        "・エスカレーション時の報告の改善：\n"
        "報告の構成を「対応要否 → 影響範囲 → 原因」の順に変え、エスカレーション時の確認の"
        "やり取りを不要にした。\n\n"
        "・OpenAPI仕様書の作成・整合化：\n"
        "プロバイダー／コンシューマー間でAPI仕様書が二重管理されている課題を解消する全社的な"
        "取り組みの一環として、担当バックエンドのAPI（全100以上のエンドポイント）のうち30以上について、"
        "実装と突合しながら仕様書を作成・修正。iOS/Androidを含む関係チームと調整。\n\n"
        "・リリース時の確認フローの整備：\n"
        "機能リリース時にモバイルとWebの両方で動作を確認するフローを整備。\n\n"
        "・テストコードの保守性向上：\n"
        "RSpecのfixturesをFactoryBotへ移行し、永続化を伴うcreateをbuildに置き換えることで、"
        "CIのテスト実行時間を短縮。",
        False,
    ),
    (
        "【実績・取り組み等】\n"
        "・エラー調査の時間を1〜2日から1時間以内に短縮し、その手順をチーム全体で使えるようにした。\n"
        "・本番データの不整合を3件特定・是正（うち1件は誰も把握していなかったもの）。\n"
        "・OpenAPI仕様書を30以上のエンドポイントで作成・修正。",
        False,
    ),
])
set_cell_blocks(ga_table.rows[3].cells[2], [(
    "【OS】\nmacOS\n\n"
    "【言語】\nRuby\nShell Script\nHTML\nCSS\nJavaScript\n\n"
    "【フレームワーク】\nRuby on Rails\n\n"
    "【DB】\nPostgreSQL\nSnowflake（参照）\n\n"
    "【インフラ・監視】\nAWS（ECS, CloudWatch, CloudTrail）\nRollbar\nDatadog\n\n"
    "【ツール】\nRSpec\nFactoryBot\nDocker\nOpenAPI\nGit\nClaude Code\nSnowflake CLI",
    False,
)])
set_cell_blocks(ga_table.rows[3].cells[3], [
    (f"【役割】\nメンバー\n\n【プロジェクト規模】\nチーム：15名\n全体：{HEADCOUNT}名", False),
])

# The RENOSY backend content row otherwise starts with a single orphaned line (the
# period and title) at the foot of the internal-platform project's page.
# It fits on one page, and the rows above it stay put, so keeping it whole
# can't strand the table header the way the trHeight bug did.
backend_trPr = ga_table.rows[3]._tr.get_or_add_trPr()
backend_trPr.insert(0, OxmlElement("w:cantSplit"))

# continuation rows' other columns stay empty, matching the original
# document's own convention.
for row_index in (2, 4):
    for ci in (0, 2, 3):
        clear_cell(ga_table.rows[row_index].cells[ci])

# ---- 6. new sections: 研究実績 and 自己研鑽・個人開発 ----
# Inserted between "※主要な案件を抜粋して記載しております。" and 取得資格等.
# Each paragraph is a list of segments: plain text, or (text, url) for a link.
THESIS_URL = "https://library.naist.jp/opac/en/book/111663"
ARXIV_URL = "https://arxiv.org/abs/2403.15430"
KECHI_URL = "https://kechi-plus-g6bg.web.app/"
FLAGSHIP_URL = "https://github.com/gill876/trunk-development-feature-flag"
HACKATHON_URL = "https://zenn.dev/hackathons/google-cloud-japan-ai-hackathon-vol5/projects/bf01c70153667d"

HEADING, BLANK = object(), object()
new_sections = [
    (HEADING, "　研究実績"),
    BLANK,
    ["・修士論文：",
     ("Can Large Language Models Outperform Rule-based Dataset Converters for Taxonomic Entities?",
      THESIS_URL),
     "（奈良先端科学技術大学院大学、2024年）"],
    ["　LLMを用いて、分類群（taxon）の固有表現認識（NER）データセットをStandoff形式からCoNLL形式へ"
     "変換するパイプラインを設計・評価。ルールベースの変換が失敗しやすいエンティティ境界のずれ、"
     "余分な空白、文字コードの混在に対応し、GPT-4oの出力に後処理を組み合わせて3つのデータセット"
     "すべてでF1 0.99以上を達成。"],
    BLANK,
    ["・共著論文（第2著者）：",
     ("Distilling Named Entity Recognition Models for Endangered Species from Large Language Models",
      ARXIV_URL),
     "（arXiv、2024年）"],
    ["　GPT-4で絶滅危惧種に関する合成データ（NER・関係抽出、計3,600文）を生成し、BERT系の小型"
     "モデルに知識を蒸留する研究。プロンプト設計によるデータ生成、IUCNレッドリストとの照合による"
     "事実確認、論文の一部執筆を担当。"],
    BLANK,
    (HEADING, "　自己研鑽・個人開発"),
    BLANK,
    # The repo is private, so this links the live app, which opens on a login screen.
    ["・個人開発：", ("Kechi+（ケチ＋）", KECHI_URL), "（ユーザー登録で利用可能）"
     "（2026年10月〜、毎週自分で使用）"],
    ["　日本のレシートを撮影すると、品目・価格・税区分を読み取り、店舗や月ごとに単価を比較できる"
     "家計簿アプリ。企画・設計判断を担当し、Claude Codeを活用して開発・運用。"],
    ["　・AI：GMI Cloud上の11種類のビジョンモデルを、自分のレシートで精度・コスト・速度の観点から"
     "比較評価。高速なモデルで読み取り、合計が合わない時だけ上位モデルで読み直す構成に決定"],
    ["　・開発・運用：トランクベース開発。GitHub ActionsでLint・テストの通過後にCloud Runへ自動"
     "デプロイし、スモークテストで確認。Workload Identity Federationによるキーレス認証"],
    ["　・監視：Sentry（エラー・トレース・リリース単位の追跡）、Grafanaダッシュボード、SLO・"
     "ランブックを整備"],
    ["　・技術：Python（FastAPI）、Next.js／TypeScript、PostgreSQL（Cloud SQL）、Cloud Run、"
     "Firebase Authentication"],
    BLANK,
    ["・個人開発：", ("Flagship", FLAGSHIP_URL),
     "（2026年）：トランクベース開発・フィーチャーフラグ・カナリアリリースを実践する検証環境を構築。"
     "企画・設計判断を担当し、Claude Codeを活用して開発・運用。"],
    ["　・mainへのマージごとに新しいリビジョンをトラフィック0%でCloud Runへデプロイし、"
     "0% → 10% → 100%の段階的なカナリアリリースで公開"],
    ["　・gemを使わずRailsの設定でフィーチャーフラグを実装（dev・alpha・betaへのセグメント配信、"
     "ユーザーごとに固定される割合配信、キルスイッチ）"],
    ["　・カナリア中は新旧リビジョンが並行稼働するため、CIでoasdiffによりOpenAPIの破壊的変更を検出"],
    ["　・12体のボットが利用者として巡回し、どのリビジョン・フラグ状態で応答したかをダッシュボードで"
     "可視化。Sentry、Grafana、SLOで監視"],
    ["　・技術：Ruby on Rails 8.1、Cloud Run、Firestore、GitHub Actions（Brakeman・bundler-audit・"
     "RuboCop）"],
    BLANK,
    ["・", ("第5回 Agentic AI Hackathon with Google Cloud", HACKATHON_URL),
     "（2026年）：LINEグループでの遊びの計画を代行するAIエージェント「AsobiAgent」を個人で企画・開発"],
    ["　企画から設計判断まで担当し、Claude CodeとAntigravityに指示を出して実装。AIによる自動の"
     "電話予約は、法人契約が必要な点や飲食店側の負担を考えて採用せず、人がワンタップで電話をかける"
     "形にした。"],
    ["　・使用技術：Vertex AI（Gemini）、Cloud Run、Firestore、Secret Manager、LINE（LIFF）、"
     "駅すぱあと API、ホットペッパー API"],
    BLANK,
    ["・学部時代にACM ICPCに参加"],
    BLANK,
]

cursor = P[68]._p  # the blank line right before 取得資格等
for item in new_sections:
    if item is BLANK:
        elem = copy.deepcopy(BLANK_P)
    elif isinstance(item, tuple) and item[0] is HEADING:
        elem = copy.deepcopy(HEADING_P)
        Paragraph(elem, P[69]._parent).runs[1].text = item[1]
    else:
        elem = copy.deepcopy(BODY_P)
        for seg in item:
            if isinstance(seg, tuple):
                append_hyperlink_field(elem, seg[1], seg[0], LINK_RPR)
            else:
                r = OxmlElement("w:r")
                r.append(copy.deepcopy(BODY_RPR))
                t = OxmlElement("w:t")
                t.set(qn("xml:space"), "preserve")
                t.text = seg
                r.append(t)
                elem.append(r)
    cursor.addnext(elem)
    cursor = elem

# ---- 7. 自己PR ----
assert P[77].text == "■　自己PR" and P[86].text == "以上"
set_paragraph_value(P[79], (
    "大学院（NAIST）では自然言語処理を専攻し、LLMを活用したデータ処理に取り組みました。修士論文では、"
    "形式のばらつきによりルールベースの変換が失敗しやすいNERデータセットを、LLMと後処理を組み合わせて"
    "変換するパイプラインを設計・評価しました。また、GPT-4で絶滅危惧種に関する合成データを作り、小型"
    "モデルに知識を蒸留する共著論文では、データ生成と事実確認、論文の一部執筆を担当しました。"
))
set_paragraph_value(P[81], (
    "現職では、本番環境のエラー調査にClaude CodeとSnowflakeを組み込み、調査時間を1〜2日から1時間以内"
    "に短縮し、その手順をチーム全体で使えるようにしました。個人開発でも、レシートを読み取る家計簿アプリ"
    "でビジョンモデルを比較評価して構成を決め、毎週自分で使いながら運用しています。今後は、AI・LLMを"
    "活用したプロダクトを本番で安定して動かせるエンジニアを目指しています。"
))
for p in P[82:86]:
    p._p.getparent().remove(p._p)

# The source closes with 以上 above the 是非 line, which is centered. Put the
# 是非 line straight after the 自己PR, left-aligned, so 以上 (right-aligned) ends it.
assert P[87].text.startswith("是非、面接の機会")
P[81]._p.addnext(P[87]._p)
P[87].alignment = WD_ALIGN_PARAGRAPH.LEFT

# ---- 7b. small text fixes from the source ----
assert P[75].runs[0].text.endswith("、自重トレ")
P[75].runs[0].text = P[75].runs[0].text.replace("自重トレ", "自重トレーニング")

# Artuvic project title: the source's "■（Doorway Express名）" was left half-filled.
artuvic_title = d.tables[2].rows[1].cells[1].paragraphs[0].runs
assert [r.text for r in artuvic_title] == ["■", "（", "Doorway Express", "名）"]
artuvic_title[1].text = "（配達向けWebアプリ「"
artuvic_title[3].text = "」の開発）"

# Artuvic 開発環境: the source lists phpMyAdmin under 【DB】. The database was
# MySQL (confirmed by Cargill); phpMyAdmin is the tool used to manage it.
artuvic_env = d.tables[2].rows[1].cells[2].paragraphs
assert artuvic_env[15].text == "【DB】" and artuvic_env[16].text == "phpMyAdmin"
assert artuvic_env[18].text == "【ツール】"
artuvic_env[18]._p.addnext(copy.deepcopy(artuvic_env[16]._p))
artuvic_env[16].runs[0].text = "MySQL"

# ---- 8. start 開発経歴 on a new page with a real page break ----
# The source pushes 開発経歴 to the next page with ~40 empty paragraphs after
# the skills table. Once the table grew, those blanks spilled onto the next
# page and left 開発経歴 starting halfway down it. Keep one blank line after
# the table and let a page break do the job regardless of the table's length.
assert P[55].text == "■　開発経歴" and not any(p.text for p in P[12:55])
for p in P[13:55]:
    p._p.getparent().remove(p._p)
P[55].paragraph_format.page_break_before = True

# ---- 9. strip stale fixed row heights across every table ----
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
