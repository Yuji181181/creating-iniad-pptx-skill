#!/usr/bin/env python3
"""スライドマスターの「Copyright © YYYY by INIAD」の年を置き換える。

年は資料ごとに変わるため、同梱テンプレは YYYY のままにしてある。デッキ作成時に
ユーザーへ年を確認し、このスクリプトで確定させる（unpack 前のコピーに対して実行する）。

使い方:
  .venv/bin/python scripts/set_copyright_year.py deck.pptx 2026
  .venv/bin/python scripts/set_copyright_year.py unpacked_dir/ 2026   # unpack 済みディレクトリにも可
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

PATTERN = re.compile(r"Copyright © (?:YYYY|\d{4}(?:[-–]\d{4})?) by INIAD")


def replace(xml, year):
    new, n = PATTERN.subn(f"Copyright © {year} by INIAD", xml)
    return new, n


def on_dir(root, year):
    total = 0
    mdir = os.path.join(root, "ppt", "slideMasters")
    for name in sorted(os.listdir(mdir)):
        if not name.endswith(".xml"):
            continue
        path = os.path.join(mdir, name)
        with open(path, encoding="utf-8") as f:
            xml, n = replace(f.read(), year)
        with open(path, "w", encoding="utf-8") as f:
            f.write(xml)
        total += n
    return total


def on_pptx(path, year):
    total = 0
    fd, tmp = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/slideMasters/") and item.filename.endswith(".xml"):
                xml, n = replace(data.decode("utf-8"), year)
                data = xml.encode("utf-8")
                total += n
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return total


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    target, year = sys.argv[1], sys.argv[2]
    if not re.fullmatch(r"\d{4}(?:[-–]\d{4})?", year):
        sys.exit(f"年の形式が不正です: {year}（例: 2026 / 2025-2026）")
    n = on_dir(target, year) if os.path.isdir(target) else on_pptx(target, year)
    if n == 0:
        sys.exit("Copyright 表記が見つかりませんでした")
    print(f"Copyright の年を {year} に置換しました（{n} 箇所）")


if __name__ == "__main__":
    main()
