#!/usr/bin/env python3
"""assets/iniad-template.pptx から references/slide-details/slideNN.md を生成する（再ベイク用）。

使い方: .venv/bin/python scripts/dump_slide_details.py assets/iniad-template.pptx references/slide-details
"""
import os
import sys

from pptx import Presentation

PURPOSE = {
    1: ("講義表紙（担当者つき）", "科目名 + 「第N回 講義タイトル」を中央に、下の短い 3 本線（両端ドット）の上に担当者名・英語表記/所属・連絡先を青字 #0033CC 太字で。各回の 1 枚目。"),
    2: ("中扉（1行）", "セクション・ノートブック名（`xxx.ipynb`）・「課題N」「演習課題」「出席確認」などの扉。中央タイトル + 中央の水色ライン。"),
    3: ("中扉（2行タイトル）", "章・節の区切り。タイトル + 括弧書きの補足など 2 行まで。"),
    4: ("標準: タイトル + 箇条書き", "最も基本の本文スライド。lvl0 ●(#C00000) / lvl1 ■(#FF9900) / lvl2 ・。重要語は赤字 #FF0000。"),
    5: ("図 + 右説明", "左半分に図・グラフ・画像（自由配置）、右に説明文。出典は右下に小さく。"),
    6: ("表", "濃灰ヘッダ（#404040・白字太字）+ 1行おきに #F2F2F2 の表。上に説明文、下に赤字の補足。評価指標のまとめ・混同行列など。"),
    7: ("数式・自由キャンバス", "タイトルのみ。数式の展開や図解を自由に配置。"),
    8: ("コード解説", "コード枠 1 つ + 下に要点（赤字）と補足説明（青字）。コードの意味を説明するときの基本形。"),
    9: ("コード + 実行結果（左右）", "左にコード枠（白地）、矢印、右に実行結果枠（#F2F2F2 地）。下に出力の読み方（赤字）と補足（青字）。print や DataFrame の出力に。"),
    10: ("コード + 図", "左にコード枠と要点・補足、右に実行結果の図（画像）。図中の注目点は青枠の吹き出しで注釈。グラフ描画コードに。"),
    11: ("長いコード（タイトルなし）", "白紙レイアウトにコード枠を全幅で（Consolas 16pt・10〜15 行まで）。前スライドのコードの続きや、ノートブックのセルをまとめて見せるときに。"),
    12: ("練習問題（解答案内つき）", "タイトルは「練習問題N」（N は番号。課題N と同じ表記）。問題文を箇条書きで。下部に「次ページに解答がある」の枠。"),
    13: ("課題", "タイトルは「課題N」（授業名・科目コードは入れない）。lvl0 ● に課題文（データの説明と求めること）、lvl1 に設問をオレンジの自動番号（1. 2. 3.、`buAutoNum arabicPeriod` + buClr #FF9900）で。"),
    14: ("講義表紙（日英）", "01 の日英併記版。科目名・講義タイトルの各行の直後に英語を青字 #0033CC で（日本語より一段小さく）。"),
    15: ("中扉（日英）", "02/03 の日英併記版。日本語タイトル（黒）の次の行に英語（青字 #0033CC）。"),
    16: ("本文（日英）", "タイトル（日本語）+ 英語タイトル（idx=13・青字 24pt）+ 箇条書き。各箇条の直後に、同じ字下げ・記号なし・一段小さい青字の英語段落。"),
    17: ("課題（日英）", "タイトル「課題N」+ 英語タイトル「Assignment N」。課題文・設問それぞれの直後に英語行。英語行を挟むと自動番号がリセットされるので、2 問目以降は `startAt` で番号を指定する。"),
    18: ("コード + 実行結果（日英）", "「タイトル（日英）のみ」レイアウト。英語タイトルの分だけ中身を下げ、コード枠は y=1900000 から。ラベルは「コード / Code」「実行結果 / Output」、要点・補足の直後に一段小さい英語行（同じ色）。日英併記のコード・表・図・数式スライドはこの配置を基準にする。"),
    19: ("図 + 右説明（日英）", "「図+説明（日英）」レイアウト。英語タイトル欄つき、図は左半分（y=1650000 から）、右の説明は日本語 18pt → 英語 14pt（青）で英語行は日本語に詰める。「見出し / Heading:」は太字。iniad_deck の desc_bi で埋める。"),
}


def emu(v):
    return f"{int(v)}"


def main(src, outdir):
    prs = Presentation(src)
    os.makedirs(outdir, exist_ok=True)
    for i, slide in enumerate(prs.slides, 1):
        name, use = PURPOSE.get(i, ("", ""))
        lines = [f"# slide{i:02d}: {name}", "",
                 f"- 使用レイアウト: {slide.slide_layout.name}",
                 f"- サムネイル: `assets/thumbnails/slide-{i:02d}.jpg`",
                 f"- 向いている内容: {use}", "",
                 "## 編集対象の枠", "",
                 "| shape 名 | 種別 (ph type, idx) | サンプル文言 | 位置(off x,y EMU) | サイズ(ext cx,cy EMU) |",
                 "|---|---|---|---|---|"]
        for sh in slide.shapes:
            if sh.is_placeholder:
                pf = sh.placeholder_format
                kind = f"placeholder ({str(pf.type).split('.')[-1].split(' ')[0].lower()}, {pf.idx})"
            else:
                kind = sh.shape_type and str(sh.shape_type).split('.')[-1].split(' ')[0].lower() or "shape"
            text = ""
            if sh.has_text_frame:
                text = " / ".join(p.text for p in sh.text_frame.paragraphs if p.text)[:60].replace("|", "\\|")
            lines.append(f"| {sh.name} | {kind} | {text} | {emu(sh.left)}, {emu(sh.top)} | {emu(sh.width)}, {emu(sh.height)} |")
        lines += ["", "## 装飾（マスター/レイアウトから継承・編集不要）", "",
                  "- 上ライン + 両端ドット、「情報連携学部」、INIAD ロゴ、下ライン + 両端ドット、Copyright 表記、ページ番号はマスターにある。スライド側で描き足さない。"]
        if slide.slide_layout.name.startswith("INIAD 中扉"):
            lines.append("- 中央ライン（y=3753000）はレイアウト側にある。タイトルは 1〜2 行で、ラインより上に収める。")
        with open(os.path.join(outdir, f"slide{i:02d}.md"), "w") as f:
            f.write("\n".join(lines) + "\n")
    print("wrote", len(prs.slides), "files to", outdir)


if __name__ == "__main__":
    main(*sys.argv[1:3])
