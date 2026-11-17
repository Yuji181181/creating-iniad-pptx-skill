#!/usr/bin/env python3
"""INIAD テンプレート pptx（マスター/レイアウト + 19 枚の見本スライド）を生成する。

使い方:
  .venv/bin/python scripts/prepare_logo.py <新しいロゴ.jpg> assets/iniad-logo.png   # ロゴを差し替えるときだけ
  .venv/bin/python scripts/build_template.py assets/iniad-template.pptx   # Copyright は YYYY のまま
"""
import os
import sys

from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "..", "assets", "iniad-logo.png")

# ---- ブランドトークン（brand-tokens.md と同期させること） ----
SW, SH = 12192000, 6858000
LINE = "5FB7E3"        # 水色ライン
HEADER = "31859C"      # 「情報連携学部」
RED = "C00000"         # 第1階層 ●
ORANGE = "FF9900"      # 第2階層 ■
EMPH = "FF0000"        # 強調（赤字）
NOTE = "0070C0"        # 注釈（青字）
EN = "0033CC"          # 日英併記の英語・講義表紙の担当者名
TEXT = "000000"
FOOT = "BFBFBF"
FONT = "Meiryo UI"
CODE_FONT = "Consolas"
LINE_W = 44450         # 3.5pt
DOT = 125000           # ライン端ドット直径

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def font_xml(face=FONT):
    return f'<a:latin typeface="{face}"/><a:ea typeface="{face}"/><a:cs typeface="{face}"/>'


# ---------------------------------------------------------------- 装飾 XML
def line_xml(sid, x1, x2, y):
    return f'''<p:cxnSp {NS}><p:nvCxnSpPr><p:cNvPr id="{sid}" name="INIAD Line {sid}"/><p:cNvCxnSpPr/><p:nvPr userDrawn="1"/></p:nvCxnSpPr>
<p:spPr><a:xfrm><a:off x="{x1}" y="{y}"/><a:ext cx="{x2 - x1}" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>
<a:ln w="{LINE_W}" cap="rnd"><a:solidFill><a:srgbClr val="{LINE}"/></a:solidFill></a:ln></p:spPr></p:cxnSp>'''


def dot_xml(sid, cx, cy):
    return f'''<p:sp {NS}><p:nvSpPr><p:cNvPr id="{sid}" name="INIAD Dot {sid}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{cx - DOT // 2}" y="{cy - DOT // 2}"/><a:ext cx="{DOT}" cy="{DOT}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{LINE}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP"/></a:p></p:txBody></p:sp>'''


def rule(sid, x1, x2, y):
    """両端ドット付きの水平ライン（INIAD の象徴的装飾）"""
    return [line_xml(sid, x1, x2, y), dot_xml(sid + 1, x1, y), dot_xml(sid + 2, x2, y)]


def textbox_xml(sid, name, x, y, cx, cy, runs_xml, algn="l", anchor="t"):
    return f'''<p:sp {NS}><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>
<a:p><a:pPr algn="{algn}"/>{runs_xml}</a:p></p:txBody></p:sp>'''


def master_decor(logo_rid, year):
    y_top, y_bot = 356000, 6515000
    out = []
    out += rule(100, 318000, 11420000, y_top)
    out += rule(110, 385000, 11885000, y_bot)
    out.append(textbox_xml(
        120, "INIAD Header", 9700000, 20000, 1640000, 300000,
        f'<a:r><a:rPr lang="ja-JP" sz="1600" dirty="0"><a:solidFill><a:srgbClr val="{HEADER}"/></a:solidFill>{font_xml()}</a:rPr><a:t>情報連携学部</a:t></a:r>',
        algn="r", anchor="b"))
    out.append(textbox_xml(
        121, "INIAD Copyright", 0, 6600000, SW, 200000,
        f'<a:r><a:rPr lang="en-US" sz="1000" b="1" dirty="0"><a:solidFill><a:srgbClr val="{FOOT}"/></a:solidFill>{font_xml("Arial")}</a:rPr><a:t>Copyright © {year} by INIAD</a:t></a:r>',
        algn="ctr"))
    out.append(textbox_xml(
        122, "INIAD Page Number", 11285000, 6600000, 600000, 200000,
        f'<a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum"><a:rPr lang="en-US" sz="1000" b="1" dirty="0"><a:solidFill><a:srgbClr val="{FOOT}"/></a:solidFill>{font_xml("Arial")}</a:rPr><a:t>‹#›</a:t></a:fld>',
        algn="r"))
    out.append(f'''<p:pic {NS}><p:nvPicPr><p:cNvPr id="123" name="INIAD Logo"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>
<p:blipFill><a:blip r:embed="{logo_rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
<p:spPr><a:xfrm><a:off x="11500000" y="25000"/><a:ext cx="620000" cy="833000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>''')
    return out


# ---------------------------------------------------------------- プレースホルダ XML
def ph_xml(sid, name, ph, x, y, cx, cy, body_pr="", lst_style="", text=""):
    para = f'<a:p><a:r><a:rPr lang="ja-JP"/><a:t>{text}</a:t></a:r></a:p>' if text else '<a:p><a:endParaRPr lang="ja-JP"/></a:p>'
    return f'''<p:sp {NS}><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>{ph}</p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr{body_pr}/><a:lstStyle>{lst_style}</a:lstStyle>{para}</p:txBody></p:sp>'''


TITLE = dict(x=260000, y=380000, cx=11100000, cy=560000)
BODY = dict(x=380000, y=1180000, cx=11430000, cy=5150000)
BODY_BI = dict(x=380000, y=1560000, cx=11430000, cy=4770000)   # 英語サブタイトルの下から
COVER_LINES = (4970000, 5460000, 5980000)                      # 講義表紙の短い 3 本線


def title_ph(sid=2, **geo):
    g = {**TITLE, **geo}
    return ph_xml(sid, "Title", '<p:ph type="title"/>', g["x"], g["y"], g["cx"], g["cy"], text="タイトル")


def body_ph(sid, idx, name, geo, text="本文レベル1"):
    return ph_xml(sid, name, f'<p:ph idx="{idx}"/>', geo["x"], geo["y"], geo["cx"], geo["cy"], text=text)


def set_tree(part_el, shapes_xml):
    """layout / master の spTree を nvGrpSpPr, grpSpPr を残して入れ替える"""
    tree = part_el.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree')
    for child in list(tree)[2:]:
        tree.remove(child)
    for s in shapes_xml:
        tree.append(parse_xml(s))


# ---------------------------------------------------------------- マスター
def build_master(prs, year):
    master = prs.slide_master
    _, logo_rid = master.part.get_or_add_image_part(LOGO)
    shapes = [
        ph_xml(2, "Title Placeholder", '<p:ph type="title"/>', TITLE["x"], TITLE["y"], TITLE["cx"], TITLE["cy"],
               body_pr=' vert="horz" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:normAutofit/></a:bodyPr', text="マスタータイトル").replace('<a:bodyPr vert', '<a:bodyPr vert', 1),
        ph_xml(3, "Text Placeholder", '<p:ph type="body" idx="1"/>', BODY["x"], BODY["y"], BODY["cx"], BODY["cy"],
               body_pr=' vert="horz" lIns="0" tIns="0" rIns="0" bIns="0"><a:normAutofit/></a:bodyPr', text="本文レベル1"),
    ]
    # ph_xml は <a:bodyPr{body_pr}/> なので、子要素付き body_pr の場合は閉じタグを整える
    shapes = [s.replace('</a:bodyPr/>', '</a:bodyPr>') for s in shapes]
    set_tree(master._element, shapes + master_decor(logo_rid, year))

    shadow = ('<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl" rotWithShape="0">'
              '<a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw></a:effectLst>')
    title_style = (f'<p:titleStyle {NS}><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
                   f'<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>'
                   f'<a:defRPr sz="4000" kern="1200"><a:solidFill><a:srgbClr val="{TEXT}"/></a:solidFill>{shadow}{font_xml()}</a:defRPr></a:lvl1pPr></p:titleStyle>')

    def lvl(n, marl, indent, sz, bu_char, bu_color, spc_bef):
        bu = (f'<a:buClr><a:srgbClr val="{bu_color}"/></a:buClr><a:buSzPct val="90000"/>'
              f'<a:buFont typeface="Arial"/><a:buChar char="{bu_char}"/>') if bu_char else '<a:buNone/>'
        return (f'<a:lvl{n}pPr marL="{marl}" indent="{indent}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
                f'<a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="{spc_bef}"/></a:spcBef>{bu}'
                f'<a:defRPr sz="{sz}" kern="1200"><a:solidFill><a:srgbClr val="{TEXT}"/></a:solidFill>{font_xml()}</a:defRPr></a:lvl{n}pPr>')

    body_style = (f'<p:bodyStyle {NS}>'
                  + lvl(1, 342900, -342900, 2800, "●", RED, 1200)
                  + lvl(2, 800100, -342900, 2400, "■", ORANGE, 400)
                  + lvl(3, 1200150, -285750, 2000, "・", TEXT, 300)
                  + lvl(4, 1600200, -285750, 1800, "-", TEXT, 300)
                  + lvl(5, 2000250, -285750, 1800, "", TEXT, 300)
                  + '</p:bodyStyle>')
    tx = master._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}txStyles')
    old_title = tx.find('{http://schemas.openxmlformats.org/presentationml/2006/main}titleStyle')
    old_body = tx.find('{http://schemas.openxmlformats.org/presentationml/2006/main}bodyStyle')
    tx.replace(old_title, parse_xml(title_style))
    tx.replace(old_body, parse_xml(body_style))

    # テーマフォントも Meiryo UI に（新規テキストボックスの既定）
    theme_part = [r.target_part for r in master.part.rels.values() if r.reltype.endswith('/theme')][0]
    xml = theme_part.blob.decode('utf-8')
    xml = xml.replace('<a:latin typeface="Calibri Light"', f'<a:latin typeface="{FONT}"').replace('<a:latin typeface="Calibri"', f'<a:latin typeface="{FONT}"')
    xml = xml.replace('<a:ea typeface=""/>', f'<a:ea typeface="{FONT}"/>')
    theme_part._blob = xml.encode('utf-8')


# ---------------------------------------------------------------- レイアウト
def en_title():
    """日英併記レイアウトの英語タイトル枠（body idx=13・青字 24pt）"""
    return ph_xml(3, "English Title", '<p:ph type="body" idx="13"/>', 300000, 960000, 11000000, 450000,
               body_pr=' lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"',
               lst_style=f'<a:lvl1pPr marL="0" indent="0"><a:spcBef><a:spcPts val="0"/></a:spcBef><a:buNone/>'
                         f'<a:defRPr sz="2400"><a:solidFill><a:srgbClr val="{EN}"/></a:solidFill></a:defRPr></a:lvl1pPr>',
               text="English Title")


LAYOUTS = [
    # (name, shapes 関数)
    ("INIAD 中扉（中央タイトル）", lambda: [
        ph_xml(2, "Title", '<p:ph type="title"/>', 600000, 1700000, 10990000, 1950000,
               body_pr=' anchor="b"', lst_style='<a:lvl1pPr algn="ctr"><a:defRPr sz="5400"/></a:lvl1pPr>', text="タイトル"),
        *rule(10, 318000, 11885000, 3753000),
    ]),
    ("INIAD タイトル+本文", lambda: [title_ph(), body_ph(3, 1, "Body", BODY)]),
    ("INIAD 図+説明（右本文）", lambda: [
        title_ph(),
        ph_xml(3, "Body Right", '<p:ph idx="1"/>', 6250000, 1180000, 5560000, 5150000,
               lst_style='<a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="2000"/></a:lvl1pPr>',
               text="図の説明文"),
    ]),
    ("INIAD タイトルのみ", lambda: [title_ph()]),
    ("INIAD 白紙（装飾のみ）", lambda: []),
    ("INIAD 講義表紙（担当者つき）", lambda: [
        ph_xml(2, "Title", '<p:ph type="title"/>', 600000, 700000, 10990000, 3600000,
               body_pr=' anchor="ctr"', lst_style='<a:lvl1pPr algn="ctr"><a:defRPr sz="4000"/></a:lvl1pPr>', text="講義タイトル"),
        *[x for i, y in enumerate(COVER_LINES) for x in rule(10 + 3 * i, 2540000, 9620000, y)],
        *[ph_xml(20 + i, f"Presenter {i + 1}", f'<p:ph type="body" idx="{20 + i}"/>', 2540000, y - 440000, 7080000, 400000,
                 body_pr=' lIns="0" tIns="0" rIns="0" bIns="0" anchor="b"',
                 lst_style=f'<a:lvl1pPr marL="0" indent="0" algn="ctr"><a:spcBef><a:spcPts val="0"/></a:spcBef><a:buNone/>'
                           f'<a:defRPr sz="2000" b="1"><a:solidFill><a:srgbClr val="{EN}"/></a:solidFill></a:defRPr></a:lvl1pPr>',
                 text=t) for i, (y, t) in enumerate(zip(COVER_LINES, ("担当者名", "Presenter Name", "連絡先")))],
    ]),
    ("INIAD タイトル（日英）+本文", lambda: [
        title_ph(),
        en_title(),
        body_ph(4, 1, "Body", BODY_BI),
    ]),
    ("INIAD タイトル（日英）のみ", lambda: [title_ph(), en_title()]),
    ("INIAD 図+説明（日英）", lambda: [
        title_ph(), en_title(),
        ph_xml(4, "Body Right", '<p:ph idx="1"/>', 6250000, 1560000, 5560000, 4750000,
               lst_style='<a:lvl1pPr marL="0" indent="0"><a:spcBef><a:spcPts val="1800"/></a:spcBef><a:buNone/><a:defRPr sz="1800"/></a:lvl1pPr>',
               text="図の説明文"),
    ]),
]


def build_layouts(prs):
    layouts = list(prs.slide_layouts)
    for i, (name, fn) in enumerate(LAYOUTS):
        lay = layouts[i]
        lay.name = name
        shapes = [s.replace('<a:bodyPr/>', '<a:bodyPr/>') for s in fn()]
        set_tree(lay._element, shapes)
    for lay in layouts[len(LAYOUTS):]:
        prs.slide_layouts.remove(lay)


# ---------------------------------------------------------------- 見本スライド
def fill(ph, paras):
    """paras: [(level, text)] もしくは [(level, [(text, {color, bold})...])]"""
    tf = ph.text_frame
    tf.text = ""
    for i, (lv, content) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = lv
        runs = content if isinstance(content, list) else [(content, {})]
        for text, st in runs:
            r = p.add_run()
            r.text = text
            if "color" in st:
                r.font.color.rgb = RGBColor.from_string(st["color"])
            if st.get("bold"):
                r.font.bold = True


def add_text(slide, x, y, cx, cy, lines, size=18, color=TEXT, face=FONT, border=None, align=None):
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(cx), Emu(cy))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if align:
            p.alignment = align
        runs = line if isinstance(line, list) else [(line, {})]
        for text, st in runs:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(st.get("size", size))
            r.font.name = st.get("face", face)
            r.font.color.rgb = RGBColor.from_string(st.get("color", color))
            if st.get("bold"):
                r.font.bold = True
            rpr = r._r.get_or_add_rPr()
            rpr.append(parse_xml(f'<a:ea {NS} typeface="{st.get("face", face)}"/>'))
    if border:
        tb.line.color.rgb = RGBColor.from_string(border)
        tb.line.width = Pt(0.75)
    return tb


def add_rect(slide, x, y, cx, cy, fill_hex, text="", size=16, color="595959"):
    from pptx.enum.shapes import MSO_SHAPE
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(cx), Emu(cy))
    s.fill.solid()
    s.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    s.line.color.rgb = RGBColor.from_string("A6A6A6")
    s.line.width = Pt(0.75)
    s.shadow.inherit = False
    for ref in s._element.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}effectRef'):
        ref.set('idx', '0')   # テーマ既定の影を消す
    if text:
        tf = s.text_frame
        tf.text = text
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                r.font.size = Pt(size)
                r.font.name = FONT
                r.font.color.rgb = RGBColor.from_string(color)
    return s


CODE_SIZE = 18          # コード枠の標準サイズ（長いコードは 16）
OUT_FILL = "F2F2F2"     # 実行結果枠の地色（コード枠は白）
LABEL = "595959"        # 「コード」「実行結果」ラベル


def code_box(slide, x, y, cx, lines, size=CODE_SIZE, fill_hex=None):
    """コード枠: 白地 + #7F7F7F 細枠 + Consolas。実行結果枠は fill_hex=OUT_FILL。
    注目箇所は (text, {"color": RED}) で色付けする"""
    h = int(Pt(size).emu * 1.25 * len(lines) + 120000)
    tb = add_text(slide, x, y, cx, h, lines, size=size, face=CODE_FONT, border="7F7F7F")
    if fill_hex:
        tb.fill.solid()
        tb.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    return tb


def code_label(slide, x, y, text):
    """コード枠・実行結果枠の直上に置く小見出し（枠の y から 380000 上に置く）"""
    return add_text(slide, x, y - 380000, 6000000, 340000, [[(text, {"bold": True, "color": LABEL})]], size=16)


def arrow(slide, x, y, cx, cy, fill_hex="7F7F7F"):
    from pptx.enum.shapes import MSO_SHAPE
    a = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Emu(x), Emu(y), Emu(cx), Emu(cy))
    a.fill.solid(); a.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    a.line.fill.background()
    for ref in a._element.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}effectRef'):
        ref.set('idx', '0')
    return a


def callout(slide, x, y, cx, cy, text, tail=(-0.8, 0.9)):
    """図中の注釈吹き出し（白地・青枠・青字）。tail は尾の先端（枠中心からの相対位置、幅/高さ比）"""
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR
    c = slide.shapes.add_shape(MSO_SHAPE.RECTANGULAR_CALLOUT, Emu(x), Emu(y), Emu(cx), Emu(cy))
    c.adjustments[0], c.adjustments[1] = tail
    c.fill.solid(); c.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
    c.line.color.rgb = RGBColor.from_string(NOTE); c.line.width = Pt(1)
    for ref in c._element.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}effectRef'):
        ref.set('idx', '0')
    tf = c.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.size = Pt(16); r.font.name = FONT; r.font.color.rgb = RGBColor.from_string(NOTE)
    r._r.get_or_add_rPr().append(parse_xml(f'<a:ea {NS} typeface="{FONT}"/>'))
    return c


def numbered(para, color=ORANGE, start=None):
    """段落を「1. 2. 3.」の自動番号（オレンジ）にする。課題の設問用。
    間に英語行など番号なし段落を挟むと番号がリセットされるので、その場合は start で開始番号を指定する"""
    pPr = para._p.get_or_add_pPr()
    for tag in ("buClr", "buSzPct", "buFont", "buChar", "buNone", "buAutoNum"):
        for el in pPr.findall(f"{{http://schemas.openxmlformats.org/drawingml/2006/main}}{tag}"):
            pPr.remove(el)
    pPr.set("marL", "800100"); pPr.set("indent", "-457200")
    for xml in (f'<a:buClr {NS}><a:srgbClr val="{color}"/></a:buClr>',
                f'<a:buFont {NS} typeface="+mj-lt"/>',
                f'<a:buAutoNum {NS} type="arabicPeriod"' + (f' startAt="{start}"' if start else '') + '/>'):
        el = parse_xml(xml)
        # buClr → buFont → buAutoNum の順で、lnSpc/spcBef/spcAft より後、defRPr より前に置く
        ref = pPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr")
        if ref is not None:
            ref.addprevious(el)
        else:
            pPr.append(el)


def table(slide, x, y, cx, rows, col_w=None, size=16, row_h=380000, header=True, first_col_bold=False, code_cols=()):
    """濃灰ヘッダ（#404040・白字）+ 1行おきに #F2F2F2 の表。
    code_cols に指定した列（演算子・書き方・例など）は見出し行以外を Consolas・太字なしにする
    （太字の Consolas は環境によって「<」などが化けるため）"""
    from pptx.enum.text import MSO_ANCHOR
    nr, nc = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nr, nc, Emu(x), Emu(y), Emu(cx), Emu(row_h * nr))
    tbl = gt.table
    tblPr = tbl._tbl.tblPr
    tblPr.set('firstRow', '0'); tblPr.set('bandRow', '0')   # テーマ既定の縞を無効化して自前で塗る
    if col_w:
        for i, w in enumerate(col_w):
            tbl.columns[i].width = Emu(w)
    for r in range(nr):
        tbl.rows[r].height = Emu(row_h)
        for c in range(nc):
            cell = tbl.cell(r, c)
            is_head = header and r == 0
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string("404040" if is_head else ("F2F2F2" if r % 2 == 0 else "FFFFFF"))
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Emu(80000)
            tf = cell.text_frame
            tf.text = ""
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if (c > 0 or is_head) else PP_ALIGN.LEFT
            run = p.add_run(); run.text = rows[r][c]
            run.font.size = Pt(size); run.font.name = FONT
            is_code = (not is_head) and c in code_cols
            face = CODE_FONT if is_code else FONT
            run.font.name = face
            run.font.bold = (not is_code) and (is_head or (first_col_bold and c == 0))
            run.font.color.rgb = RGBColor.from_string("FFFFFF" if is_head else TEXT)
            run._r.get_or_add_rPr().append(parse_xml(f'<a:ea {NS} typeface="{face}"/>'))
    return gt


BI_MARL = {0: 342900, 1: 800100, 2: 1200150}   # bodyStyle の marL と同期
BI_SIZE = {0: 2400, 1: 2000, 2: 1800}          # 英語行は日本語より一段小さく


def fill_bi(ph, items):
    """日英併記の箇条書き。items: [(level, 日本語, English or None)]。
    日本語の段落（箇条記号つき）の直後に、同じ字下げ・記号なし・青字 #0033CC の英語段落を置く"""
    tf = ph.text_frame
    tf.text = ""
    first = True
    for lv, jp, en in items:
        for text, is_en in ((jp, False), (en, True)):
            if is_en and en is None:      # 英語を持たない行（URL・固有名詞など）は日本語だけ
                continue
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.level = lv
            r = p.add_run(); r.text = text
            if is_en:
                pPr = p._p.get_or_add_pPr()
                pPr.set("marL", str(BI_MARL[lv])); pPr.set("indent", "0")
                pPr.append(parse_xml(f'<a:spcBef {NS}><a:spcPts val="0"/></a:spcBef>'))
                pPr.append(parse_xml(f'<a:buNone {NS}/>'))
                r.font.size = Pt(BI_SIZE[lv] / 100)
                r.font.color.rgb = RGBColor.from_string(EN)


def bi_title(slide, jp, en, jp_size=None, en_size=None):
    """中扉・表紙のタイトルを「日本語（黒）+ 改行 + English（青）」にする"""
    tf = slide.shapes.title.text_frame
    tf.text = ""
    for i, (text, color, size) in enumerate(((jp, None, jp_size), (en, EN, en_size))):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = text
        if color:
            r.font.color.rgb = RGBColor.from_string(color)
        if size:
            r.font.size = Pt(size)


def cover_title(slide, lines):
    """講義表紙のタイトル。lines: [(text, size_pt, color_hex or None)]（科目名・回とタイトル など）"""
    tf = slide.shapes.title.text_frame
    tf.text = ""
    for i, (text, size, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = text
        r.font.size = Pt(size)
        if color:
            r.font.color.rgb = RGBColor.from_string(color)


def build_samples(prs):
    """見本スライド。文言は汎用のダミー（「タイトル」「本文レベル1」等）のみ"""
    by = {l.name: l for l in prs.slide_layouts}
    C, T, FIG, TO, BL = (by["INIAD 中扉（中央タイトル）"], by["INIAD タイトル+本文"], by["INIAD 図+説明（右本文）"],
                         by["INIAD タイトルのみ"], by["INIAD 白紙（装飾のみ）"])
    COVER, TBI = by["INIAD 講義表紙（担当者つき）"], by["INIAD タイトル（日英）+本文"]
    TOBI = by["INIAD タイトル（日英）のみ"]
    FIGBI = by["INIAD 図+説明（日英）"]

    def new(layout):
        return prs.slides.add_slide(layout)

    def ph(slide, idx):
        return [p for p in slide.placeholders if p.placeholder_format.idx == idx][0]

    def presenters(slide, texts=("担当者名", "Presenter Name", "連絡先（メールアドレス）")):
        for i, t in enumerate(texts):
            ph(slide, 20 + i).text = t

    # 01 講義表紙（科目名 + 回・タイトル + 担当者）
    s = new(COVER); cover_title(s, [("科目名", 32, None), ("第N回 講義タイトル", 44, None)]); presenters(s)
    # 02 中扉（1行）: セクション・ノートブック名・「課題N」などの扉
    s = new(C); s.shapes.title.text = "タイトル"
    # 03 中扉（2行）
    s = new(C); s.shapes.title.text = "セクションタイトル\n(サブタイトル)"
    # 04 タイトル + 箇条書き（強調色の例つき）
    s = new(T); s.shapes.title.text = "タイトル"
    fill(ph(s, 1), [
        (0, "本文レベル1"),
        (0, [("本文レベル1（", {}), ("強調は赤字", {"color": EMPH}), ("）", {})]),
        (1, "本文レベル2"),
        (2, "本文レベル3"),
    ])
    # 05 図 + 右説明
    s = new(FIG); s.shapes.title.text = "タイトル"
    add_rect(s, 380000, 1300000, 5500000, 4900000, "F2F2F2", "図・グラフ領域\n（画像を貼る）", 20)
    add_text(s, 700000, 1500000, 1500000, 400000, [[("ラベルA", {"color": EMPH, "bold": True})]], size=20)
    add_text(s, 4300000, 5300000, 1500000, 400000, [[("ラベルB", {"color": NOTE, "bold": True})]], size=20)
    fill(ph(s, 1), [(0, "図の説明文"),
                    (0, [("説明文（", {}), ("重要語", {"color": EMPH}), ("は赤字）", {})])])
    add_text(s, 9300000, 6120000, 2500000, 300000, ["出典（著者 年）"], size=14, align=PP_ALIGN.RIGHT)
    # 06 表（ヘッダ付きテーブル）
    s = new(TO); s.shapes.title.text = "タイトル"
    add_text(s, 380000, 1200000, 11000000, 450000, ["表の説明文"], size=22)
    table(s, 1200000, 1850000, 9800000, [
        ["項目", "列見出し1", "列見出し2", "列見出し3"],
        ["行見出し1", "0.00", "0.00", "0.00"],
        ["行見出し2", "0.00", "0.00", "0.00"],
        ["行見出し3", "0.00", "0.00", "0.00"],
        ["行見出し4", "0.00", "0.00", "0.00"],
        ["行見出し5", "0.00", "0.00", "0.00"],
    ], col_w=[2900000, 2300000, 2300000, 2300000], size=18, row_h=520000, first_col_bold=True)
    add_text(s, 1200000, 5100000, 9800000, 400000, [[("表に関する補足（赤字）", {"color": EMPH})]], size=18)
    # 07 数式・自由キャンバス
    s = new(TO); s.shapes.title.text = "タイトル"
    add_text(s, 500000, 1300000, 11000000, 4600000, [
        "数式の説明文",
        "f(x) = a₁x₁ + a₂x₂ + … + aₙxₙ",
        [("数式に関する強調（赤字）", {"color": EMPH})],
    ], size=22, face="Cambria Math")
    # ---- コード関連（08〜11）: 共通ルール
    #   コード枠 = 白地 / 実行結果枠 = #F2F2F2 地（どちらも #7F7F7F 細枠・Consolas 18pt）
    #   枠の直上に「コード」「実行結果」ラベル、要点は赤字 #FF0000、補足説明は青字 #0070C0
    #   コード中の注目箇所は #C00000 で色付け
    key = [("variable", {"color": RED}), (" = ", {}), ("function", {"color": RED}), ("(arg1, arg2)", {})]

    # 08 コード解説（コード + 要点 + 補足説明）
    s = new(TO); s.shapes.title.text = "タイトル"
    code_label(s, 380000, 1600000, "コード")
    code_box(s, 380000, 1600000, 8000000, [key, [("variable", {"color": RED}), (".method(x)", {})], "code_line_3()"])
    add_text(s, 380000, 3000000, 11400000, 1800000, [
        [("コードの要点（赤字）", {"color": EMPH})],
        [("コードの補足説明（青字）", {"color": NOTE})],
    ], size=20)

    # 09 コード + 実行結果（左右）
    s = new(TO); s.shapes.title.text = "タイトル"
    code_label(s, 380000, 1600000, "コード")
    code_box(s, 380000, 1600000, 5500000, [key, "print(variable)", "code_line_3()", "code_line_4()"])
    arrow(s, 6020000, 2100000, 420000, 360000)
    code_label(s, 6580000, 1600000, "実行結果")
    code_box(s, 6580000, 1600000, 5230000, ["output_1", "output_2", "output_3", "output_4"], fill_hex=OUT_FILL)
    add_text(s, 380000, 3500000, 11400000, 1800000, [
        [("出力の読み方（赤字）", {"color": EMPH})],
        [("補足説明（青字）", {"color": NOTE})],
    ], size=20)

    # 10 コード + 図（実行結果の図 + 注釈吹き出し）
    s = new(TO); s.shapes.title.text = "タイトル"
    code_label(s, 380000, 1600000, "コード")
    code_box(s, 380000, 1600000, 5400000, ["code_line_1()", "plot_function(x, y)", "plt.show()"])
    add_text(s, 380000, 2900000, 5400000, 2000000, [
        [("図の要点（赤字）", {"color": EMPH})],
        [("図の見方の補足説明（青字）", {"color": NOTE})],
    ], size=20)
    code_label(s, 6100000, 1600000, "実行結果")
    add_rect(s, 6100000, 1600000, 5710000, 4400000, OUT_FILL, "実行結果の図\n（画像を貼る）", 20)
    callout(s, 8700000, 4700000, 2000000, 600000, "図中の注釈", tail=(-0.75, -1.1))

    # 11 長いコード（タイトルなし・コードを全面に。前スライドの続きなど）
    s = new(BL)
    code_label(s, 380000, 1000000, "コード（続き）")
    code_box(s, 380000, 1000000, 11430000, ["code_line_1()", "code_line_2()", "code_line_3()", "code_line_4()",
                                            [("important_line()", {"color": RED})], "code_line_6()", "code_line_7()",
                                            "code_line_8()", "code_line_9()", "code_line_10()"], size=16)
    add_text(s, 380000, 4300000, 11430000, 1500000, [[("補足説明（青字）", {"color": NOTE})]], size=20)

    # 12 練習問題（解答案内つき）
    s = new(T); s.shapes.title.text = "練習問題N"
    fill(ph(s, 1), [(0, "問題文"), (1, "補足・ヒント")])
    add_text(s, 4700000, 5600000, 3000000, 380000, ["次ページに解答がある"], size=18, border="7F7F7F", align=PP_ALIGN.CENTER)
    # 13 課題（「課題N」+ 課題文 + 番号付きの設問）
    s = new(T); s.shapes.title.text = "課題N"
    body = ph(s, 1)
    fill(body, [(0, "課題文（使用するデータの説明と、求めること）"),
                (1, "設問1"), (1, "設問2"), (1, "設問3")])
    for para in body.text_frame.paragraphs[1:]:
        numbered(para)

    # ---- 日英併記（14〜17）: 日本語の直後に英語を青字 #0033CC で。英語は一段小さく
    # 14 講義表紙（日英）
    s = new(COVER)
    cover_title(s, [("科目名", 30, None), ("Course Name", 26, EN), ("第N回 講義タイトル", 40, None), ("Lecture N: Lecture Title", 34, EN)])
    presenters(s, ("担当者名", "Presenter Name", "連絡先（メールアドレス）"))
    # 15 中扉（日英）
    s = new(C); bi_title(s, "セクションタイトル", "Section Title", en_size=44)
    # 16 本文（日英）: タイトル日英 + 箇条書き日英
    s = new(TBI); s.shapes.title.text = "タイトル"
    ph(s, 13).text = "English Title"
    fill_bi(ph(s, 1), [(0, "本文レベル1", "Body level 1"), (1, "本文レベル2", "Body level 2"), (0, "本文レベル1", "Body level 1")])
    # 17 課題（日英）
    s = new(TBI); s.shapes.title.text = "課題N"
    ph(s, 13).text = "Assignment N"
    body = ph(s, 1)
    fill_bi(body, [(0, "課題文（使用するデータの説明と、求めること）", "Assignment statement (data description and what to do)"),
                   (1, "設問1", "Question 1"), (1, "設問2", "Question 2")])
    ps = body.text_frame.paragraphs
    numbered(ps[2]); numbered(ps[4], start=2)
    for i in (3, 5):   # 英語行は番号の後ろの字下げに合わせる
        ps[i]._p.get_or_add_pPr().set("marL", "800100")

    # 18 コード + 実行結果（日英）: タイトル（日英）のみレイアウト。英語タイトルの分だけ全体を下げる（y=1900000 から）
    s = new(TOBI); s.shapes.title.text = "タイトル"
    ph(s, 13).text = "English Title"
    code_label(s, 380000, 1900000, "コード / Code")
    code_box(s, 380000, 1900000, 5500000, [key, "print(variable)", "code_line_3()"])
    arrow(s, 6020000, 2350000, 420000, 360000)
    code_label(s, 6580000, 1900000, "実行結果 / Output")
    code_box(s, 6580000, 1900000, 5230000, ["output_1", "output_2", "output_3"], fill_hex=OUT_FILL)
    add_text(s, 380000, 3500000, 11400000, 1800000, [
        [("出力の読み方（赤字）", {"color": EMPH})], [("How to read the output", {"color": EMPH, "size": 16})],
        [("補足説明（青字）", {"color": NOTE})], [("Supplementary note", {"color": NOTE, "size": 16})],
    ], size=20)

    # 19 図 + 右説明（日英）: 図は左半分（y=1560000 から）、右の説明は日本語 18pt → 英語 14pt（青）
    s = new(FIGBI); s.shapes.title.text = "タイトル"
    ph(s, 13).text = "English Title"
    add_rect(s, 380000, 1650000, 5500000, 4550000, "F2F2F2", "図・グラフ領域\n（画像を貼る）", 20)
    tf = ph(s, 1).text_frame
    tf.text = ""
    for i, (runs) in enumerate([
            [("図の説明文", TEXT, 18, False)], [("Description of the figure", EN, 14, False)],
            [("見出し / Heading: ", TEXT, 18, True), ("要点（赤字）", EMPH, 18, False)], [("Key point", EMPH, 14, False)],
            [("補足（青字）", NOTE, 16, False)], [("Supplementary note", NOTE, 14, False)]]):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if i % 2 == 1:   # 英語行は直前の日本語に詰める
            para._p.get_or_add_pPr().append(parse_xml(f'<a:spcBef {NS}><a:spcPts val="0"/></a:spcBef>'))
        for text, color, size, bold in runs:
            r = para.add_run(); r.text = text
            r.font.size = Pt(size); r.font.bold = bold
            r.font.color.rgb = RGBColor.from_string(color)


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("out", nargs="?", default="iniad-template.pptx")
    ap.add_argument("--year", default="YYYY",
                    help="Copyright 年。同梱テンプレは未確定を示す YYYY のまま生成し、デッキ作成時にユーザーに聞いて置換する")
    a = ap.parse_args()
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(SW), Emu(SH)
    build_master(prs, a.year)
    build_layouts(prs)
    build_samples(prs)
    prs.save(a.out)
    print("saved", a.out)


if __name__ == "__main__":
    main()
