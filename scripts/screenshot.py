#!/usr/bin/env python3
"""スライドに貼るスクリーンショットを Playwright（Chromium）で撮る。

使い方（スキルのディレクトリで。依存は requirements-screenshot.txt）:
  # Web ページ・ローカル HTML
  .venv/bin/python scripts/screenshot.py url https://example.com -o output/<デッキ名>/shots/page.png
  .venv/bin/python scripts/screenshot.py url page.html -o output/<デッキ名>/shots/part.png --selector "#main" --width 1280
  # 記事内の画像（遅延読み込み）の 3 番目を撮る
  .venv/bin/python scripts/screenshot.py url https://… -o output/<デッキ名>/shots/fig3.png --selector ".entry-content img" --nth 3 --scroll
  # 撮れる候補の一覧（番号・alt・大きさ）を表示する
  .venv/bin/python scripts/screenshot.py url https://… -o - --selector ".entry-content img" --scroll --list

  # Jupyter Notebook（Jupyter の見た目のまま撮る）
  .venv/bin/python scripts/screenshot.py notebook work.ipynb -o output/<デッキ名>/shots/cells.png --cells 2-3
  .venv/bin/python scripts/screenshot.py notebook work.ipynb -o output/<デッキ名>/shots/cell.png --cells 4 --execute   # 実行してから撮る

  # Python コードから Notebook を作り、実行して撮る（「# %%」でセルを区切る）
  .venv/bin/python scripts/screenshot.py pycode demo.py -o output/<デッキ名>/shots/demo.png --cells 1-2

共通オプション:
  --width     ブラウザの幅（px。既定 800）。スライドに貼ったときの文字の大きさは「貼る幅 ÷ この幅」で決まる。
              横に長い出力なら 1000 程度に。notebook / pycode は 760 未満にすると Jupyter が狭い画面用の表示
              （In [n]: がコードの上に出る）に切り替わるので、760 以上にする
  --scale     解像度の倍率（既定 2。スライドで拡大してもぼやけないように）
  --wait      描画を待つ時間（ms。既定 500）
  --full-page ページ全体（url モードで --selector なしのとき）
  --scroll    撮る前にページの最後までスクロールする（遅延読み込みの画像を表示させる）
  --nth       --selector に当てはまる要素のうち N 番目（1 始まり）を撮る。「5-8,12」のように複数指定するときは
              -o に {n} を入れる（例: -o output/<デッキ名>/shots/fig{n}.png）
  --list      撮らずに、--selector に当てはまる要素の一覧を表示する

出力は PNG。スライドへの配置は iniad_deck.add_image() を使う。
"""
import argparse
import os
import re
import sys
import tempfile


def parse_cells(spec, n):
    """"2-3,5" → [1, 2, 4]（0 始まりの添字）。None ならすべて"""
    if not spec:
        return list(range(n))
    out = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            out += list(range(int(a) - 1, int(b)))
        else:
            out.append(int(part) - 1)
    bad = [i + 1 for i in out if not 0 <= i < n]
    if bad:
        sys.exit(f"セル番号が範囲外です: {bad}（このノートブックのセル数は {n}）")
    return out


def notebook_to_html(nb, no_input=False, no_output=False):
    from nbconvert import HTMLExporter
    exp = HTMLExporter(template_name="lab")
    exp.exclude_input = no_input
    exp.exclude_output = no_output
    exp.exclude_input_prompt = False
    body, _ = exp.from_notebook_node(nb)
    return body


def execute(nb, cwd):
    from nbclient import NotebookClient
    NotebookClient(nb, timeout=120, kernel_name="python3", resources={"metadata": {"path": cwd}}).execute()


def py_to_notebook(path):
    """「# %%」区切りの .py からノートブックを作る（区切りが無ければ全体を 1 セル）"""
    import nbformat
    src = open(path, encoding="utf-8").read()
    chunks = [c.strip("\n") for c in re.split(r"^# %%.*$", src, flags=re.M)]
    chunks = [c for c in chunks if c.strip()]
    nb = nbformat.v4.new_notebook()
    nb.cells = [nbformat.v4.new_code_cell(c) for c in chunks]
    nb.metadata["kernelspec"] = {"name": "python3", "display_name": "Python 3", "language": "python"}
    return nb


def shoot(target, out, width, scale, wait, selector=None, full_page=False, cells=None, scroll=False, nth=1, list_only=False):
    """target（URL か HTML ファイル）を開いて撮る。cells があれば Jupyter のセル要素の範囲だけを 1 枚に"""
    from playwright.sync_api import sync_playwright
    if os.path.exists(target):
        target = "file://" + os.path.abspath(target)
    if not list_only:
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": width, "height": 800}, device_scale_factor=scale)
        try:
            page.goto(target, wait_until="networkidle", timeout=45000)
        except Exception:        # 広告などで networkidle にならないサイトは読み込み完了で進める
            page.goto(target, wait_until="load", timeout=60000)
        if scroll:
            h = page.evaluate("document.body.scrollHeight")
            for y in range(0, h, 600):
                page.evaluate(f"window.scrollTo(0, {y})"); page.wait_for_timeout(120)
            page.evaluate("window.scrollTo(0, 0)")
        page.wait_for_timeout(wait)
        if cells is not None:
            els = page.query_selector_all(".jp-Cell")
            idx = parse_cells(cells, len(els))
            boxes = [els[i].bounding_box() for i in idx]
            pad = 8
            x0 = min(b["x"] for b in boxes) - pad
            y0 = min(b["y"] for b in boxes)          # 上は余白なし（前のセルの端が写り込まないように）
            x1 = max(b["x"] + b["width"] for b in boxes) + pad
            y1 = max(b["y"] + b["height"] for b in boxes)    # 下も余白なし（次のセルの端が写り込まないように）
            page.screenshot(path=out, clip={"x": max(x0, 0), "y": max(y0, 0), "width": x1 - x0, "height": y1 - y0},
                            full_page=True)
        elif selector:
            els = [e for e in page.query_selector_all(selector) if (e.bounding_box() or {}).get("width", 0) > 1]
            if list_only:
                for i, e in enumerate(els, 1):
                    b = e.bounding_box()
                    alt = e.get_attribute("alt") or e.inner_text()[:40] if e else ""
                    print(f"{i:3d}  {int(b['width'])}x{int(b['height'])}  y={int(b['y'])}  {alt!r}")
                browser.close()
                return
            if not els:
                sys.exit(f"要素が見つかりません: {selector}")
            targets = [i + 1 for i in parse_cells(str(nth), len(els))]
            if len(targets) > 1 and "{n}" not in out:
                sys.exit("複数の --nth を撮るときは -o に {n} を入れてください（例: output/<デッキ名>/shots/fig{n}.png）")
            for n in targets:
                path = out.replace("{n}", str(n))
                els[n - 1].scroll_into_view_if_needed(); page.wait_for_timeout(300)
                els[n - 1].screenshot(path=path)
                print("saved", path)
            browser.close()
            return
        else:
            page.screenshot(path=out, full_page=full_page)
        browser.close()
    print("saved", out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("mode", choices=["url", "notebook", "pycode"])
    ap.add_argument("source", help="URL / HTML ファイル / .ipynb / .py")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--selector", help="url モード: この CSS セレクタの要素だけを撮る")
    ap.add_argument("--full-page", action="store_true")
    ap.add_argument("--cells", help="notebook / pycode モード: 撮るセル番号（1 始まり。例 2-3,5）。省略時は全セル")
    ap.add_argument("--execute", action="store_true", help="notebook モード: 撮る前に実行する（pycode は常に実行）")
    ap.add_argument("--no-input", action="store_true", help="コードを隠して出力だけ撮る")
    ap.add_argument("--no-output", action="store_true", help="出力を隠してコードだけ撮る")
    ap.add_argument("--scroll", action="store_true")
    ap.add_argument("--nth", default="1")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--width", type=int, default=800)
    ap.add_argument("--scale", type=float, default=2)
    ap.add_argument("--wait", type=int, default=500)
    a = ap.parse_args()

    if a.mode == "url":
        shoot(a.source, a.out, a.width, a.scale, a.wait, selector=a.selector, full_page=a.full_page,
              scroll=a.scroll, nth=a.nth, list_only=a.list)
        return

    import nbformat
    if a.mode == "notebook":
        nb = nbformat.read(a.source, as_version=4)
        if a.execute:
            execute(nb, os.path.dirname(os.path.abspath(a.source)))
    else:
        nb = py_to_notebook(a.source)
        execute(nb, os.path.dirname(os.path.abspath(a.source)))
    html = notebook_to_html(nb, a.no_input, a.no_output)
    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as f:
        f.write(html)
        tmp = f.name
    try:
        shoot(tmp, a.out, a.width, a.scale, a.wait, cells=a.cells or "")
    finally:
        os.unlink(tmp)


if __name__ == "__main__":
    main()
