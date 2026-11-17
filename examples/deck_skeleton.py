#!/usr/bin/env python3
"""デッキ生成スクリプトのひな形（文言はすべてダミー）。

新しいデッキを作るときは、このファイルを output/<デッキ名>/build.py にコピーし、
仕様書（output/<デッキ名>/spec.md）どおりにスライドを足し引きして文言を差し替える。
主要なスライドの型（講義表紙・中扉・本文・図+説明・表・コード・画像・練習問題・課題）を一通り含む。

実行（スキルのディレクトリで）:
    .venv/bin/python examples/deck_skeleton.py                 # 日英併記で output/deck_skeleton/ に出力
    .venv/bin/python examples/deck_skeleton.py ja 2026         # 日本語のみ・Copyright 2026
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))   # output/<デッキ名>/build.py に置くときは "..", "..", "scripts"
from iniad_deck import *  # noqa: E402,F403

NAME = "deck_skeleton"                                   # デッキ名（output/<NAME>/ に出力）
BILINGUAL = (sys.argv[1] if len(sys.argv) > 1 else "bi") != "ja"   # True: 日英併記 / False: 日本語のみ
YEAR = sys.argv[2] if len(sys.argv) > 2 else "YYYY"     # 実際はユーザーに確認した年を入れる
OUT = os.path.join(deck_dir(NAME), f"{NAME}.pptx")

d = Deck()

# 言語モードで変わるもの（レイアウト・中身の開始位置）
L_BODY = "本文日英" if BILINGUAL else "本文"
L_FIG = "図説明日英" if BILINGUAL else "図説明"
L_TITLE_ONLY = "タイトルのみ日英" if BILINGUAL else "タイトルのみ"
Y_TOP = Y_BODY_BI if BILINGUAL else Y_BODY              # 本文・表・図の上端
Y_CODE = Y_CODE_BI if BILINGUAL else 1600000             # コード枠の上端（ラベルはその 380000 上）


def title(s, jp, en):
    title_bi(s, jp, en if BILINGUAL else None)


def bullets(s, items):
    """items: [(level, 日本語, English)]。日本語のみのときは English を無視する"""
    if BILINGUAL:
        fill_bi(ph(s, 1), items)
    else:
        fill(ph(s, 1), [(lv, jp) for lv, jp, _ in items])


def desc(s, items):
    """図+説明の右本文。items: [(日本語, English, 色)]"""
    if BILINGUAL:
        desc_bi(ph(s, 1), items)
    else:
        para_list(ph(s, 1), [x for jp, _, c in items for x in ((jp, c, 20), ("", TEXT, 8))][:-1])


def put_notes(s, x, y, cx, items, size=20):
    """枠の外の要点・補足。items: [(日本語, English, 色)]"""
    notes(s, x, y, cx, items, size=size, bilingual=BILINGUAL)


def label(jp, en):
    return f"{jp} / {en}" if BILINGUAL else jp


def answer_box(s):
    body = ["次ページに解答がある"] + ([[("Answer on the next page", {"color": EN, "size": 14})]] if BILINGUAL else [])
    add_text(s, 4300000, 5550000, 3600000, 700000 if BILINGUAL else 380000, body, size=18, border="7F7F7F", align=PP_ALIGN.CENTER)


# ---------------------------------------------------------------- 1. 講義表紙（見本 01 / 14）
s = d.new("講義表紙")
if BILINGUAL:
    cover_title(s, [("科目名", 30, None), ("Course Name", 26, EN), ("第N回 講義タイトル", 40, None), ("Lecture N: Lecture Title", 34, EN)])
else:
    cover_title(s, [("科目名", 32, None), ("第N回 講義タイトル", 44, None)])
presenters(s)

# ---------------------------------------------------------------- 2. 本日の内容（見本 04 / 16）
s = d.new(L_BODY); title(s, "本日の内容", "Today's Topics")
bullets(s, [(0, "項目1", "Topic 1"), (0, "項目2", "Topic 2"), (0, "項目3", "Topic 3"), (0, "練習問題・課題", "Exercise and assignment")])

# ---------------------------------------------------------------- 3. 中扉（見本 02 / 15）
s = d.new("中扉")
if BILINGUAL:
    bi_title(s, "セクションタイトル", "Section Title", en_size=44)
else:
    s.shapes.title.text = "セクションタイトル"

# ---------------------------------------------------------------- 4. 本文（見本 04 / 16）
s = d.new(L_BODY); title(s, "タイトル", "English Title")
bullets(s, [
    (0, "本文レベル1", "Body level 1"),
    (1, "本文レベル2", "Body level 2"),
    (0, "本文レベル1", "Body level 1"),
    (1, "本文レベル2", "Body level 2"),
])
ps = ph(s, 1).text_frame.paragraphs
for p in ps[-2 if BILINGUAL else -1:]:          # 強調したい行を赤字にする例（最後の項目）
    for r in p.runs:
        r.font.color.rgb = RGBColor.from_string(EMPH)

# ---------------------------------------------------------------- 5. 図 + 右説明（見本 05 / 19）
s = d.new(L_FIG); title(s, "タイトル", "English Title")
# 図は図形で描くか、画像を置く: add_image(s, user_image("xxx.png"), 380000, Y_TOP + 90000, 5500000, 4450000)
add_rect(s, 380000, Y_TOP + 90000, 5500000, 4450000, "F2F2F2", "図・グラフ領域\n（図形で描くか画像を置く）", 20)
desc(s, [
    ("図の説明文", "Description of the figure", TEXT),
    ("要点（赤字）", "Key point", EMPH),
    ("補足（青字）", "Supplementary note", NOTE),
])
add_text(s, 5300000, 6130000, 6500000, 300000, [[("出典: 著者名「資料名」", {"color": "7F7F7F"})]], size=11, align=PP_ALIGN.RIGHT)

# ---------------------------------------------------------------- 6. 表（見本 06）
s = d.new(L_TITLE_ONLY); title(s, "タイトル", "English Title")
add_text(s, 700000, Y_TOP + 90000, 10800000, 450000, ["表の説明文"] + ([[("Description of the table", {"color": EN, "size": 16})]] if BILINGUAL else []), size=20)
table(s, 700000, Y_TOP + 800000, 10800000, [
    [label("項目", "Item"), label("列見出し1", "Column 1"), label("列見出し2", "Column 2"), label("コード", "Code")],
    [label("行見出し1", "Row 1"), "0.00", "0.00", "code()"],
    [label("行見出し2", "Row 2"), "0.00", "0.00", "code()"],
    [label("行見出し3", "Row 3"), "0.00", "0.00", "code()"],
], col_w=[3000000, 2500000, 2500000, 2800000], size=18, row_h=560000, first_col_bold=True, code_cols=(3,))
put_notes(s, 700000, Y_TOP + 800000 + 4 * 560000 + 400000, 10800000, [("表に関する補足（赤字）", "Note on the table", EMPH)])

# ---------------------------------------------------------------- 7. コード解説（見本 08）
s = d.new(L_TITLE_ONLY); title(s, "タイトル", "English Title")
code_label(s, 380000, Y_CODE, label("コード", "Code"))
code_box(s, 380000, Y_CODE, 7000000, [
    [kw("variable"), (" = function(arg1, arg2)", {})],
    "code_line_2()",
    "code_line_3()",
])
put_notes(s, 380000, Y_CODE + 1500000, 11400000, [("コードの要点（赤字）", "Key point of the code", EMPH), ("補足説明（青字）", "Supplementary note", NOTE)])

# ---------------------------------------------------------------- 8. コード + 実行結果（見本 09 / 18）
s = d.new(L_TITLE_ONLY); title(s, "タイトル", "English Title")
code_label(s, 380000, Y_CODE, label("コード", "Code"))
code_box(s, 380000, Y_CODE, 5500000, [[kw("variable"), (" = function(arg)", {})], "print(variable)", "code_line_3()"])
arrow(s, 6020000, Y_CODE + 450000, 420000, 360000)
code_label(s, 6580000, Y_CODE, label("実行結果", "Output"))
code_box(s, 6580000, Y_CODE, 5230000, ["output_1", "output_2"], fill_hex=OUT_FILL)
put_notes(s, 380000, Y_CODE + 1500000, 11400000, [("出力の読み方（赤字）", "How to read the output", EMPH), ("補足説明（青字）", "Supplementary note", NOTE)])

# ---------------------------------------------------------------- 9. 画像・スクリーンショット
s = d.new(L_TITLE_ONLY); title(s, "タイトル", "English Title")
code_label(s, 380000, Y_CODE, label("実行画面", "Screenshot"))
# 実際は: add_image(s, os.path.join(deck_dir(NAME), "shots", "xxx.png"), 380000, Y_CODE, 9500000, 3000000, align="left")
add_rect(s, 380000, Y_CODE, 9500000, 3000000, "F2F2F2", "スクリーンショット（shots/xxx.png）", 20)
put_notes(s, 380000, Y_CODE + 3200000, 11400000, [("画像の要点（赤字）", "Key point of the image", EMPH)])

# ---------------------------------------------------------------- 10. 練習問題N（見本 12）
s = d.new(L_BODY); title(s, "練習問題N", "Exercise N")
bullets(s, [(0, "問題文", "Problem statement"), (1, "補足・ヒント", "Hint")])
answer_box(s)

# ---------------------------------------------------------------- 11. 練習問題N 解答（見本 09 の配置）
s = d.new(L_TITLE_ONLY); title(s, "練習問題N 解答", "Exercise N: Answer")
code_label(s, 380000, Y_CODE, label("解答例", "Sample answer"))
code_box(s, 380000, Y_CODE, 7000000, ["code_line_1()", "code_line_2()"])
put_notes(s, 380000, Y_CODE + 1200000, 11400000, [("解答の要点（赤字）", "Key point of the answer", EMPH), ("補足説明（青字）", "Supplementary note", NOTE)])

# ---------------------------------------------------------------- 12. 課題N（見本 13 / 17）
s = d.new(L_BODY); title(s, "課題N", "Assignment N")
bullets(s, [
    (0, "課題文（使用するデータの説明と、求めること）", "Assignment statement (data and what to do)"),
    (1, "設問1", "Question 1"), (1, "設問2", "Question 2"), (1, "設問3", "Question 3"),
])
body = ph(s, 1)
if BILINGUAL:
    numbered_bi(body)
else:
    for p in body.text_frame.paragraphs[1:]:
        numbered(p)

d.save(OUT, year=YEAR)
