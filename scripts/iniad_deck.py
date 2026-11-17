#!/usr/bin/env python3
"""INIAD テンプレでデッキを組むための部品（python-pptx）。

テンプレ（assets/iniad-template.pptx）を開いて見本スライドを消し、レイアウトから
スライドを追加して、build_template.py と同じ部品（コード枠・表・日英箇条など）で中身を置く。
部品の見た目は見本スライドと同じルールなので、座標と文言を決めるだけでよい。

使い方（デッキ生成スクリプトの冒頭）:
    import sys; sys.path.insert(0, "<スキルのディレクトリ>/scripts")
    from iniad_deck import *
    d = Deck()                              # テンプレを開き、見本スライドを削除
    s = d.new("中扉"); bi_title(s, "セクションタイトル", "Section Title")
    ...
    d.save("out.pptx", year="2026")         # 保存 + Copyright 年の置換

ひな形: examples/deck_skeleton.py
"""
import os
import re
import zipfile
import shutil
import tempfile

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml

from build_template import (  # noqa: F401  （デッキ側からまとめて import できるように再公開）
    add_text, add_rect, code_box, code_label, arrow, callout, table, fill, fill_bi, bi_title, cover_title,
    numbered, EMPH, NOTE, EN, RED, ORANGE, TEXT, LINE, FONT, CODE_FONT, NS, OUT_FILL, BI_MARL,
)

# from iniad_deck import * で公開する名前（内部の os / re / shutil などを利用側に漏らさない）
__all__ = [
    # デッキ
    "Deck", "LAYOUT", "TEMPLATE", "USER_IMAGES", "OUTPUT_ROOT", "deck_dir", "user_image", "add_image", "Y_BODY", "Y_BODY_BI", "Y_CODE_BI", "set_year",
    # 部品
    "ph", "title_bi", "presenters", "set_geometry", "notes", "kw", "shape_box", "connector", "para_list", "desc_bi", "numbered_bi",
    "add_text", "add_rect", "code_box", "code_label", "arrow", "callout", "table", "fill", "fill_bi", "bi_title",
    "cover_title", "numbered",
    # 色・フォント
    "EMPH", "NOTE", "EN", "RED", "ORANGE", "TEXT", "LINE", "FONT", "CODE_FONT", "NS", "OUT_FILL", "BI_MARL",
    # python-pptx でよく使うもの
    "Emu", "Pt", "RGBColor", "MSO_SHAPE", "MSO_CONNECTOR", "PP_ALIGN", "MSO_ANCHOR", "parse_xml",
]

_HERE = os.path.dirname(os.path.abspath(__file__))   # import * で利用側の変数を上書きしないよう _ 始まり
TEMPLATE = os.path.join(_HERE, "..", "assets", "iniad-template.pptx")
USER_IMAGES = os.path.normpath(os.path.join(_HERE, "..", "user-images"))   # ユーザーが画像を置くフォルダ
OUTPUT_ROOT = os.path.normpath(os.path.join(_HERE, "..", "output"))        # 生成物の出力先（固定）

# 短い名前 → レイアウト名
LAYOUT = {
    "講義表紙": "INIAD 講義表紙（担当者つき）",
    "中扉": "INIAD 中扉（中央タイトル）",
    "本文": "INIAD タイトル+本文",
    "本文日英": "INIAD タイトル（日英）+本文",
    "図説明": "INIAD 図+説明（右本文）",
    "タイトルのみ": "INIAD タイトルのみ",
    "タイトルのみ日英": "INIAD タイトル（日英）のみ",
    "図説明日英": "INIAD 図+説明（日英）",
    "白紙": "INIAD 白紙（装飾のみ）",
}

# 日英併記で英語タイトルがあるスライドは、本文・コード枠を y=1560000〜1900000 から始める
Y_BODY = 1180000        # 英語タイトルなし
Y_BODY_BI = 1560000     # 英語タイトルあり（本文プレースホルダの位置）
Y_CODE_BI = 1900000     # 英語タイトルあり（コード枠。ラベルがその 380000 上に入る）


class Deck:
    def __init__(self, template=TEMPLATE):
        self.prs = Presentation(template)
        lst = self.prs.slides._sldIdLst
        for sldId in list(lst):            # 見本スライドをすべて削除（マスターとレイアウトは残る）
            self.prs.part.drop_rel(sldId.rId)
            lst.remove(sldId)
        self.layouts = {l.name: l for l in self.prs.slide_layouts}

    def new(self, kind):
        """kind は LAYOUT のキー（"本文" など）かレイアウト名そのもの"""
        return self.prs.slides.add_slide(self.layouts[LAYOUT.get(kind, kind)])

    def save(self, out, year):
        """保存して Copyright の YYYY を year に置換する（year は必ずユーザーに確認した値）"""
        self.prs.save(out)
        n = set_year(out, year)
        if n == 0:
            raise RuntimeError("Copyright 表記が見つかりません")
        print(f"saved {out} ({len(self.prs.slides)} slides, Copyright {year})")


def set_year(path, year):
    pat = re.compile(r"Copyright © (?:YYYY|\d{4}(?:[-–]\d{4})?) by INIAD")
    total = 0
    fd, tmp = tempfile.mkstemp(suffix=".pptx"); os.close(fd)
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/slideMasters/") and item.filename.endswith(".xml"):
                xml, n = pat.subn(f"Copyright © {year} by INIAD", data.decode("utf-8"))
                data = xml.encode("utf-8"); total += n
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return total


def ph(slide, idx):
    """idx でプレースホルダを取る（本文=1、英語タイトル=13、講義表紙の担当者欄=20〜22）"""
    return [p for p in slide.placeholders if p.placeholder_format.idx == idx][0]


def title_bi(slide, jp, en=None):
    """タイトルを入れる。en があれば英語タイトル枠（idx=13）へ。枠が無いレイアウトでは同じ位置にテキストボックスを置く"""
    slide.shapes.title.text = jp
    if en is None:
        return
    try:
        ph(slide, 13).text = en
    except IndexError:
        add_text(slide, 300000, 960000, 11000000, 450000, [[(en, {"color": EN})]], size=24)


def presenters(slide, name="担当者名", name_en="Presenter Name", contact="連絡先（メールアドレス）"):
    for i, t in enumerate((name, name_en, contact)):
        ph(slide, 20 + i).text = t


def set_geometry(shape, x, y, cx, cy):
    """位置・サイズの変更。継承プレースホルダは top/height だけ変えると left/width が 0 になって消えるので、必ず 4 つとも指定する"""
    shape.left, shape.top, shape.width, shape.height = Emu(x), Emu(y), Emu(cx), Emu(cy)


def notes(slide, x, y, cx, items, size=20, bilingual=False):
    """枠の外の要点・補足。items: [(日本語, English or None, 色)]。色は EMPH（要点）/ NOTE（補足）/ TEXT。
    bilingual=True なら各行の直後に英語を一段小さく入れる（要点・補足の英語は同じ色、黒文字の英語は EN の青）"""
    lines = []
    for jp, en, color in items:
        lines.append([(jp, {"color": color})])
        if bilingual and en:
            lines.append([(en, {"color": EN if color == TEXT else color, "size": size - 4})])
    return add_text(slide, x, y, cx, 400000 * len(lines), lines, size=size)


def kw(text):
    """コード枠内の注目箇所（#C00000）。code_box の行に [kw("for"), (" i in x:", {})] のように使う"""
    return (text, {"color": RED})


def shape_box(slide, kind, x, y, cx, cy, text, en=None, size=18, fill_hex="FFFFFF", line_hex=LINE):
    """フロー図などの図形（白地・水色枠 2.25pt・中央揃え）。en を渡すと 2 行目に英語（青・一段小さく）"""
    sh = slide.shapes.add_shape(kind, Emu(x), Emu(y), Emu(cx), Emu(cy))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    sh.line.color.rgb = RGBColor.from_string(line_hex); sh.line.width = Pt(2.25)
    for ref in sh._element.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}effectRef'):
        ref.set('idx', '0')                         # テーマ既定の影を消す
    tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, (t, color, sz) in enumerate([(text, TEXT, size)] + ([(en, EN, size - 4)] if en else [])):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = t
        r.font.size = Pt(sz); r.font.name = FONT; r.font.color.rgb = RGBColor.from_string(color)
    return sh


def connector(slide, x1, y1, x2, y2, color="7F7F7F", width=2):
    """矢印付きの直線（フロー図の矢印）"""
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(x1), Emu(y1), Emu(x2), Emu(y2))
    c.line.color.rgb = RGBColor.from_string(color); c.line.width = Pt(width)
    c.line._get_or_add_ln().append(parse_xml(f'<a:tailEnd {NS} type="triangle" w="med" len="med"/>'))
    return c


def para_list(ph_or_shape, items):
    """箇条記号なしで段落を並べる（図+説明の右本文など）。items: [(text, color, size_pt)]。空文字は段落の間隔用"""
    tf = ph_or_shape.text_frame
    tf.text = ""
    for i, (t, color, size) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = t
        r.font.size = Pt(size); r.font.color.rgb = RGBColor.from_string(color)


def desc_bi(body, items):
    """「図説明日英」の右本文を日英で埋める。items: [(日本語 or [(text, color, bold)...], English, 色)]。
    日本語 18pt → 直後に英語 14pt（色は要点・補足ならその色、黒文字なら EN の青）。英語行は日本語に詰める"""
    tf = body.text_frame
    tf.text = ""
    first = True
    for jp, en, color in items:
        for is_en in (False, True):
            if is_en and not en:
                continue
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            if is_en:
                p._p.get_or_add_pPr().append(parse_xml(f'<a:spcBef {NS}><a:spcPts val="0"/></a:spcBef>'))
                runs = [(en, EN if color == TEXT else color, False)]
            else:
                runs = jp if isinstance(jp, list) else [(jp, color, False)]
            for text, c, bold in runs:
                r = p.add_run(); r.text = text
                r.font.size = Pt(14 if is_en else 18); r.font.bold = bold
                r.font.color.rgb = RGBColor.from_string(c)


def deck_dir(name):
    """デッキの出力先 output/<name>/ を返す（shots/ も作る）。
    ここに spec.md（仕様書）・build.py（生成スクリプト）・<name>.pptx・shots/（スクリーンショット）を置く"""
    path = os.path.join(OUTPUT_ROOT, name)
    os.makedirs(os.path.join(path, "shots"), exist_ok=True)
    return path


def user_image(name):
    """user-images/ 内の画像のパスを返す（サブフォルダも可: "ai-10/svm.png"）。無ければ一覧を添えてエラー"""
    path = os.path.join(USER_IMAGES, name)
    if not os.path.exists(path):
        found = sorted(os.path.relpath(os.path.join(r, f), USER_IMAGES)
                       for r, _, fs in os.walk(USER_IMAGES) for f in fs if f != "README.md")
        raise FileNotFoundError(f"{path} がありません。user-images にある画像: {found}")
    return path


def add_image(slide, path, x, y, max_w, max_h, border=True, align="center"):
    """画像を (x, y, max_w, max_h) の枠に縦横比を保って収めて置く。
    border=True で #BFBFBF 0.75pt の細枠（スクリーンショットの輪郭をはっきりさせる。元資料の Jupyter 画像と同じ扱い）。
    align: "center"（枠の中央）/ "left"（左上寄せ）。戻り値は picture shape"""
    from PIL import Image
    with Image.open(path) as im:
        w, h = im.size
    k = min(max_w / w, max_h / h)
    cw, ch = int(w * k), int(h * k)
    px = x + (max_w - cw) // 2 if align == "center" else x
    py = y + (max_h - ch) // 2 if align == "center" else y
    pic = slide.shapes.add_picture(path, Emu(px), Emu(py), Emu(cw), Emu(ch))
    if border:
        pic.line.color.rgb = RGBColor.from_string("BFBFBF")
        pic.line.width = Pt(0.75)
    return pic


def numbered_bi(body):
    """fill_bi で入れた課題の設問（lvl1）を 1. 2. 3. にする。英語行を挟んでも番号が続くよう startAt を付ける"""
    n = 0
    for p in body.text_frame.paragraphs:
        pPr = p._p.pPr
        is_en = pPr is not None and pPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}buNone') is not None
        if p.level == 1 and not is_en:
            n += 1
            numbered(p, start=n if n > 1 else None)
        elif p.level == 1 and is_en:
            pPr.set("marL", "800100")   # 番号の後ろの字下げに合わせる
