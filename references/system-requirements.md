# システム要件と準備手順

このスキルは「INIAD 固有の知識（テンプレ・ブランド値・レイアウトカタログ）」だけを持ち、
pptx を操作する機構そのものは `/pptx`（document-skills:pptx）スキルに依存している。
さらに `/pptx` のワークフロー（特に画像化してのビジュアル QA）は外部ツールに依存する。

`scripts/check_requirements.py` で充足状況をまとめて確認できる。**不足が出たら、本ファイルの
該当セクション（`#python` などのアンカー）を読み、ユーザーに合わせて準備を丁寧に案内すること。**
ユーザー環境を勝手に書き換えず、何を・なぜ入れるのかを説明してから実行する。

## 案内の進め方

1. まず venv を用意し（下記 [venv](#venv)）、`.venv/bin/python scripts/check_requirements.py` を実行して結果を見せる
2. `[必須]` に NG があれば、該当セクションの手順で導入を案内（環境＝macOS / Linux / claude.ai サンドボックスを見極める）
3. `[任意]` の NG は、その作業をするときだけ案内すればよい（理由を添える）
4. 再度チェッカーを実行して `結果: OK` を確認してから本編ワークフローに進む

> 環境の見分け方: `soffice` などが既に揃っていて `pip install` がネットワーク制限で通りにくい場合は claude.ai サンドボックスの可能性が高い（多くは最初から導入済み）。ローカルの macOS なら Homebrew、Linux なら apt が基本線。

---

<a id="pptx-skill"></a>
## /pptx スキル本体（最重要・必須）

**役割:** unpack / add_slide / clean / pack / thumbnail / extract-text と QA の機構。このスキルの
ワークフローはすべてこの機構の上で動く。これが無いと成立しない。

**確認:** スクリプトは `extract-text` コマンドの有無で間接的に判定するだけ。スキル本体の有無は
スクリプトからは確証できないため、**AI 自身が「available skills」一覧に `document-skills:pptx`
（または `pptx`）があるかを確認する**こと。

**準備:**
- Claude Code: pptx スキルを含むプラグイン（document-skills / anthropic-skills など）を導入する
- claude.ai: ファイル作成（コード実行）環境に標準で備わっており、通常は意識不要
- これが見当たらない環境では、このスキルは利用できない旨をユーザーに伝える

---

<a id="python"></a>
## Python 3.9+（必須）

**役割:** `/pptx` の各スクリプト（unpack/pack/thumbnail 等）と本チェッカーの実行基盤。

**確認:** `python3 --version`

**準備:**
- macOS: `brew install python` もしくは mise / pyenv 等で導入
- Debian/Ubuntu: `sudo apt-get install -y python3 python3-venv`
- claude.ai サンドボックス: 標準で利用可能

---

<a id="venv"></a>
## venv + pip（必須）

**役割:** このスキルの Python 依存（Pillow / python-pptx / markitdown）は、スキル直下の仮想環境 `.venv` に
`pip` で入れて使う。**システムの Python に `pip install` しない**（Homebrew の Python は PEP 668 で拒否されることもある）。
スクリプトはすべて `.venv/bin/python scripts/xxx.py` の形で実行する。

**準備（スキルのディレクトリで、初回のみ）:**
```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

**確認:** `.venv/bin/python scripts/check_requirements.py` の `venv (.venv)` が OK になること。

> `/pptx` スキル側のスクリプト（unpack/pack/thumbnail 等）も、この `.venv/bin/python` で実行してよい（依存は requirements.txt で揃う）。

---

<a id="pillow"></a>
## Pillow (PIL)（必須）

**役割:** `/pptx` のサムネイルグリッド生成（`thumbnail.py`）。

**確認:** `.venv/bin/python -c "import PIL; print(PIL.__version__)"`

**準備:**
- 共通: `.venv/bin/pip install -r requirements.txt`（venv に入れる。システムの pip には入れない）
- claude.ai サンドボックス: 通常は導入済み

---

<a id="libreoffice"></a>
## LibreOffice (soffice)（必須）

**役割:** ビジュアル QA のために pptx → PDF へ変換する。`/pptx` の `scripts/office/soffice.py`
がサンドボックス環境向けに自動設定する前提。

**確認:** `soffice --version`（macOS で PATH に無い場合は
`/Applications/LibreOffice.app/Contents/MacOS/soffice --version`）

**準備:**
- macOS: `brew install --cask libreoffice`（または公式サイトから .app を導入）
- Debian/Ubuntu: `sudo apt-get install -y libreoffice`
- claude.ai サンドボックス: 標準で利用可能（`soffice.py` が自動設定）

> QA の画像検査を省略するなら必須ではなくなるが、仕上がり確認の品質が落ちるため原則導入を勧める。

---

<a id="poppler"></a>
## Poppler (pdftoppm)（必須）

**役割:** 変換した PDF を 1 枚ずつの画像（jpg）にして、視覚検査に回す。

**確認:** `pdftoppm -v`

**準備:**
- macOS: `brew install poppler`
- Debian/Ubuntu: `sudo apt-get install -y poppler-utils`
- claude.ai サンドボックス: 通常は導入済み

---

<a id="markitdown"></a>
## markitdown（必須）

**役割:** QA でプレースホルダ残り（`xxxx` / lorem ipsum / TODO 等）を検出するためのテキスト化。
CLI かモジュールのどちらかがあればよい。

**確認:** `.venv/bin/markitdown --help`

**準備:**
- 共通: `.venv/bin/pip install -r requirements.txt`（`markitdown[pptx]` を含む）
- claude.ai サンドボックス: 通常は導入済み

> 代替として `/pptx` の `extract-text` でも本文抽出は可能。markitdown が用意できない場合は
> `extract-text` ベースでプレースホルダ検出に切り替えてよい。

---

<a id="python-pptx"></a>
## python-pptx（必須）

**役割:** テンプレ `assets/iniad-template.pptx` を**再生成**する（`scripts/build_template.py`）ときと、
`references/slide-details/` を再ベイクする（`scripts/dump_slide_details.py`）ときに使う。
Copyright 年の置換（`scripts/set_copyright_year.py` は標準ライブラリのみだが venv で実行する）と合わせ、venv に入れておく。

**確認:** `.venv/bin/python -c "import pptx; print(pptx.__version__)"`

**準備:**
- 共通: `.venv/bin/pip install -r requirements.txt`

---

<a id="playwright"></a>
## Playwright + Chromium（任意・スクリーンショット用）

**役割:** `scripts/screenshot.py` で Web ページ・HTML・Jupyter Notebook のスクリーンショットを撮る。
Notebook を実行してから撮る場合は nbclient / ipykernel と、ノートブックが使うライブラリ（pandas など）も要る。
スクリーンショットを使わないなら **不要**。

**準備（スキルのディレクトリで）:**
```bash
.venv/bin/pip install -r requirements-screenshot.txt
.venv/bin/python -m playwright install chromium    # ブラウザ本体（約 100MB）をダウンロード
```

**確認:** `.venv/bin/python scripts/check_requirements.py` の「playwright + Chromium」が OK になること。

---

## 環境別サマリ

| 要件 | macOS (Homebrew) | Windows (winget) | Debian/Ubuntu (apt) | claude.ai サンドボックス |
|---|---|---|---|---|
| /pptx スキル | プラグイン導入 | プラグイン導入 / スキルを配置 | プラグイン導入 | 標準 |
| Python 3.9+ | `brew install python` | `winget install Python.Python.3.12` | `apt install python3 python3-venv` | 標準 |
| venv | `python3 -m venv .venv` | `py -m venv .venv` | `python3 -m venv .venv` | `python3 -m venv .venv` |
| Pillow / python-pptx / markitdown | `.venv/bin/pip install -r requirements.txt` | `.venv\Scripts\pip install -r requirements.txt` | macOS と同じ | macOS と同じ |
| LibreOffice | `brew install --cask libreoffice` | `winget install TheDocumentFoundation.LibreOffice` | `apt install libreoffice` | 標準（自動設定） |
| Poppler | `brew install poppler` | `winget install oschwartz10612.Poppler` | `apt install poppler-utils` | 通常導入済み |
| Playwright + Chromium（任意） | `.venv/bin/pip install -r requirements-screenshot.txt` → `.venv/bin/python -m playwright install chromium` | `.venv\Scripts\pip install -r requirements-screenshot.txt` → `.venv\Scripts\python -m playwright install chromium` | macOS と同じ | ネットワーク制限に注意 |

> Windows では、コマンド中の `.venv/bin/python` を `.venv\Scripts\python` に読み替える。LibreOffice は既定では PATH に入らないが、要件チェッカーは既定のインストール先（`C:\Program Files\LibreOffice\program\soffice.exe`）も探す。winget で入れたあとは、新しいターミナルを開き直すと PATH が反映される。
