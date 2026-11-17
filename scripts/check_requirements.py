#!/usr/bin/env python3
"""
creating-iniad-pptx のシステム要件チェッカー。

このスキルは pptx 操作の機構を /pptx（document-skills:pptx）スキルに依存しており、
さらに /pptx のワークフロー（画像化 QA など）が外部ツールに依存する。
本スクリプトはそれらが揃っているかを確認し、足りないものを列挙する。

各項目の意味と準備手順は references/system-requirements.md に丁寧に書いてある。
不足が出たら、その項目名で同リファイルを引いてユーザーを案内すること。

終了コード:
  0 = 必須要件をすべて満たす（任意要件の不足は警告のみ）
  1 = 必須要件に不足がある

使い方（Python 依存はスキル直下の venv に入れる。システムの pip には入れない）:
  python3 -m venv .venv && .venv/bin/pip install -r requirements.txt   # 初回のみ
  .venv/bin/python scripts/check_requirements.py
"""

import importlib.util
import shutil
import subprocess
import sys
from pathlib import Path

# macOS で LibreOffice が PATH に入っていない場合の既定インストール先
MAC_SOFFICE = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
# Windows で LibreOffice が PATH に入っていない場合の既定インストール先
WIN_SOFFICE = [Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
               Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe")]
IS_WIN = sys.platform.startswith("win")


def find_executable(name, extra_paths=()):
    """PATH と追加候補から実行ファイルを探し、見つかればパスを返す。"""
    found = shutil.which(name)
    if found:
        return found
    for cand in extra_paths:
        if Path(cand).exists():
            return str(cand)
    return None


def python_module_available(module_name):
    """import 可能なら True。"""
    try:
        return importlib.util.find_spec(module_name) is not None
    except (ImportError, ValueError):
        return False


def check_python():
    v = sys.version_info
    ok = (v.major, v.minor) >= (3, 9)
    detail = f"{v.major}.{v.minor}.{v.micro}"
    return ok, detail, "references/system-requirements.md#python"


SKILL_DIR = Path(__file__).resolve().parent.parent
VENV_BIN = SKILL_DIR / ".venv" / ("Scripts" if IS_WIN else "bin")   # Windows は Scripts


def check_venv():
    # Python 依存（Pillow / python-pptx / markitdown）は venv + pip で管理する
    in_venv = sys.prefix != sys.base_prefix
    exists = (VENV_BIN / ("python.exe" if IS_WIN else "python")).exists()
    if in_venv:
        return True, f"venv 内で実行中 ({sys.prefix})", "references/system-requirements.md#venv"
    if exists:
        return False, ".venv はあるが venv 外で実行中（.venv/bin/python、Windows は .venv\\Scripts\\python で実行し直す）", "references/system-requirements.md#venv"
    return False, ".venv 未作成", "references/system-requirements.md#venv"


def check_pillow():
    ok = python_module_available("PIL")
    detail = "import PIL OK" if ok else "import PIL 失敗"
    return ok, detail, "references/system-requirements.md#pillow"


def check_soffice():
    path = find_executable("soffice", extra_paths=[MAC_SOFFICE, *WIN_SOFFICE])
    return bool(path), (path or "未検出"), "references/system-requirements.md#libreoffice"


def check_pdftoppm():
    path = find_executable("pdftoppm")
    return bool(path), (path or "未検出"), "references/system-requirements.md#poppler"


def check_markitdown():
    # CLI かモジュールのどちらかがあれば QA のプレースホルダ検出に使える
    cli = find_executable("markitdown", extra_paths=[VENV_BIN / ("markitdown.exe" if IS_WIN else "markitdown")])
    mod = python_module_available("markitdown")
    ok = bool(cli) or mod
    detail = (cli or ("python module" if mod else "未検出"))
    return ok, detail, "references/system-requirements.md#markitdown"


def check_pptx_tooling():
    # /pptx スキルが提供する extract-text の存在は、機構ツールが入っている強いシグナル。
    # ただしスキル本体の有無はスクリプトからは確証できないため任意扱い（AI が available skills で確認する）。
    path = find_executable("extract-text")
    return bool(path), (path or "未検出（/pptx スキルの読み込みを AI 側で確認）"), \
        "references/system-requirements.md#pptx-skill"


def check_python_pptx():
    # Copyright 年の置換・テンプレ再生成（build_template.py / dump_slide_details.py）に使う。
    try:
        import pptx  # noqa: F401
        return True, f"インストール済み ({pptx.__version__})", "references/system-requirements.md#python-pptx"
    except ImportError:
        return False, "未検出", "references/system-requirements.md#python-pptx"


def check_playwright():
    # スクリーンショット（scripts/screenshot.py）を撮るときだけ必要
    ref = "references/system-requirements.md#playwright"
    if not python_module_available("playwright"):
        return False, "未導入（スクリーンショットを撮るときだけ必要）", ref
    cache = Path.home() / "Library" / "Caches" / "ms-playwright"
    if sys.platform.startswith("linux"):
        cache = Path.home() / ".cache" / "ms-playwright"
    elif IS_WIN:
        cache = Path.home() / "AppData" / "Local" / "ms-playwright"
    has_browser = cache.exists() and any(p.name.startswith("chromium") for p in cache.iterdir())
    return has_browser, ("Chromium 導入済み" if has_browser else "Chromium 未導入（playwright install chromium）"), ref


# (ラベル, チェック関数, 必須かどうか)
CHECKS = [
    ("Python 3.9+", check_python, True),
    ("venv (.venv)", check_venv, True),
    ("Pillow (PIL)", check_pillow, True),
    ("LibreOffice (soffice)", check_soffice, True),
    ("Poppler (pdftoppm)", check_pdftoppm, True),
    ("markitdown", check_markitdown, True),
    ("/pptx ツール (extract-text)", check_pptx_tooling, False),
    ("python-pptx", check_python_pptx, True),
    ("playwright + Chromium (任意)", check_playwright, False),
]


def main():
    print("creating-iniad-pptx システム要件チェック")
    print("=" * 44)

    missing_required = []
    missing_optional = []

    for label, fn, required in CHECKS:
        ok, detail, ref = fn()
        tag = "[必須]" if required else "[任意]"
        mark = "OK  " if ok else "NG  "
        print(f"{tag} {label:<28} {mark} {detail}")
        if not ok:
            (missing_required if required else missing_optional).append((label, ref))

    print("-" * 44)

    if missing_required:
        print("必須要件に不足があります。次を参照して準備を案内してください:")
        for label, ref in missing_required:
            print(f"  - {label}: {ref}")
        if missing_optional:
            print("（参考）任意要件の不足:")
            for label, ref in missing_optional:
                print(f"  - {label}: {ref}")
        print("\n結果: NG（必須要件に不足あり）")
        return 1

    print("必須要件はすべて満たしています。")
    if missing_optional:
        print("（参考）任意要件の不足 — 該当作業をする場合のみ準備:")
        for label, ref in missing_optional:
            print(f"  - {label}: {ref}")
    print("\n結果: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
