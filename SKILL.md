---
name: creating-iniad-pptx
description: INIAD（東洋大学 情報連携学部）の講義スライドテンプレートに基づいて PowerPoint / pptx を作るときに使う。INIAD 形式の講義資料・演習資料・課題スライドを作る、既存の内容を INIAD スライド化する場合（「INIAD のテンプレでスライドにして」「INIAD の講義資料っぽい pptx を作って」）。ブランド指定なしの汎用 pptx は素の pptx スキルを使う。
---

# Creating INIAD pptx（INIAD テンプレで pptx を作る）

> **【最重要】ユーザーへの応答は必ず日本語で行うこと。** このスキルは英語で書かれた `/pptx`（document-skills:pptx / anthropic-skills:pptx）スキルに依存しており、その内容を読むと応答が英語に引きずられがちです。`/pptx` をロードした後も、進捗報告・確認・説明・最終回答のすべてを **日本語** で書くこと。

> **【必須】スライドを生成する前に、1 ページずつの仕様書（`references/spec-template.md` の形式）を書いてユーザーの承認を得ること。** plan モード・仕様駆動開発と同じく、承認前に生成スクリプトを書いたり pptx を出力したりしない。

## Overview

INIAD の講義資料のデザインを実測して再構築したテンプレ（16:9・マスター/レイアウト 9 種 + 見本 19 スライド）を**同梱**し、構造・レイアウト・ブランド値を**事前ベイク済み**。毎回のテンプレ分析を省き、`scripts/iniad_deck.py`（python-pptx）の部品で pptx を組む。QA の機構（validate・画像化）は `/pptx` スキルに依存する。

**Goal:** 「INIAD のテンプレでスライド化して」で INIAD 講義資料らしい pptx が出ること。

## 他の AI エージェント・Windows で使う場合
このスキルは Claude Code を前提に書いているが、Codex などの Agent Skills 形式に対応したエージェントでも使える。その場合は次のように読み替える。
- **`AskUserQuestion` が無いエージェント**: 同じ質問を、選択肢つきでチャットに示して回答を待つ。仕様書の承認も同じ（承認の返事があるまで生成しない）
- **`/pptx` スキルが無いエージェント**: Anthropic の skills リポジトリの pptx スキルを同じスキルの置き場所に入れるか、少なくとも `soffice --headless --convert-to pdf` と `pdftoppm` で全スライドを画像にして目で確認する QA は省略しない
- **Windows**: コマンド中の `.venv/bin/python` を `.venv\Scripts\python`（pip・markitdown も同様）に読み替える。準備は `references/system-requirements.md` の環境別サマリ（winget）を参照

## When to use / 棲み分け
- INIAD 形式の講義・演習・課題スライドを作る・既存内容を INIAD スライド化する → このスキル
- ブランド指定なしの汎用 pptx → 素の `/pptx`

## 同梱物
- `assets/iniad-template.pptx` — テンプレ本体（編集の土台）。**装飾はすべてスライドマスターにある**
- `assets/thumbnails/slide-NN.jpg` — 見本 17 スライドの見た目（LibreOffice 描画。フォントは代替表示）
- `assets/iniad-logo.png` — ユーザー提供の公式ロゴを、余白を切り詰めて白背景を透過したもの（テンプレに埋め込み済み）
- `references/brand-tokens.md` — 配色・フォント・サイズ・座標の確定値（厳守）
- `references/slide-details/slideNN.md` — 各見本スライドの枠・座標（使うスライドの分だけ読む）
- `references/spec-template.md` — **生成前に書く仕様書の形式と承認の取り方**（必読）
- `references/system-requirements.md` — システム要件と環境別の準備手順
- `scripts/iniad_deck.py` — **デッキを組む部品**（テンプレを開いて見本を消す `Deck`、コード枠・表・日英箇条・フロー図など）
- `examples/deck_skeleton.py` — `iniad_deck.py` で組む生成スクリプトの**ひな形**（文言はすべてダミー。主要なスライドの型を一通り含み、`BILINGUAL` で日本語のみ / 日英併記を切り替え）。新しいデッキはこれをコピーして中身を差し替える
- `scripts/screenshot.py` — **スクリーンショット撮影**（Playwright。Web ページ・HTML・Jupyter Notebook）
- `user-images/` — **ユーザーがスライドに使いたい画像を置くフォルダ**（中身は Git 管理外）
- `output/` — **生成物の出力先**（デッキごとに `output/<デッキ名>/`。中身は Git 管理外）
- `scripts/check_requirements.py` — システム要件チェッカー（着手前に実行）
- `scripts/set_copyright_year.py` — Copyright の年（テンプレでは `YYYY`）を置換
- `requirements.txt` — Python 依存（venv に pip で入れる）
- `requirements-screenshot.txt` — スクリーンショット用の追加依存（使うときだけ）
- `scripts/build_template.py` / `scripts/prepare_logo.py` / `scripts/dump_slide_details.py` — テンプレ再生成・再ベイク用

## ブランド要約（詳細は references/brand-tokens.md）
- 16:9 / 12192000×6858000 EMU、背景は白
- 装飾（マスター）: 上下の水色ライン #5FB7E3（両端ドット付き）、右上「情報連携学部」#31859C + INIAD ロゴ、下部中央「Copyright © YYYY by INIAD」#BFBFBF（**年はテンプレでは `YYYY`。毎回ユーザーに確認して置換**）、右下ページ番号
- 文字: 黒 #000000、強調の赤字 #FF0000、注釈の青字 #0070C0、日英併記の英語 #0033CC（注釈の青とは別色）
- 箇条書き: 第1階層 ● #C00000（28pt）/ 第2階層 ■ #FF9900（24pt）/ 第3階層 ・（20pt）
- フォント: Meiryo UI（和欧とも）、コードは Consolas、数式は Cambria Math。タイトルは 40pt・左上・影付き

## レイアウト（スライドマスター配下）
| レイアウト名 | プレースホルダ |
|---|---|
| INIAD 講義表紙（担当者つき） | title（中央・科目名と講義タイトルを段落で）/ body idx=20・21・22（担当者名・英語表記/所属・連絡先。青字 #0033CC 太字 20pt）+ 短い 3 本線 |
| INIAD 中扉（中央タイトル） | title（中央 54pt）+ 中央ライン |
| INIAD タイトル+本文 | title / body idx=1 |
| INIAD タイトル（日英）+本文 | title / body idx=13（英語タイトル・青字 #0033CC 24pt）/ body idx=1 |
| INIAD タイトル（日英）のみ | title / body idx=13（英語タイトル）。日英併記のコード・表・図・数式スライド用。中身は y=1560000（コード枠は 1900000）から |
| INIAD 図+説明（右本文） | title / body idx=1（右半分・箇条なし 20pt）。左半分は図を自由配置 |
| INIAD 図+説明（日英） | title / body idx=13（英語タイトル）/ body idx=1（右半分・y=1560000 から・箇条なし 18pt）。英語行 14pt を足しても収まる大きさ |
| INIAD タイトルのみ | title |
| INIAD 白紙（装飾のみ） | なし |

## 見本スライドカタログ（19 スライド）
> 番号の入る文言は「第N回」「練習問題N」「課題N」のように **N で表記**している。実際のデッキでは N を数字に置き換える。

| # | 用途 | レイアウト | 詳細 |
|---|---|---|---|
| 01 | 講義表紙（科目名・第N回タイトル・担当者・連絡先） | 講義表紙 | [slide01.md](references/slide-details/slide01.md) |
| 02 | 中扉（1行）。セクション・ノートブック名・「課題N」・「出席確認」などの扉 | 中扉 | [slide02.md](references/slide-details/slide02.md) |
| 03 | 中扉（2行タイトル・章区切り） | 中扉 | [slide03.md](references/slide-details/slide03.md) |
| 04 | 標準: タイトル + 箇条書き（赤字強調の例） | タイトル+本文 | [slide04.md](references/slide-details/slide04.md) |
| 05 | 図 + 右説明（出典つき） | 図+説明 | [slide05.md](references/slide-details/slide05.md) |
| 06 | 表（濃灰ヘッダ + 縞） | タイトルのみ | [slide06.md](references/slide-details/slide06.md) |
| 07 | 数式・自由キャンバス | タイトルのみ | [slide07.md](references/slide-details/slide07.md) |
| 08 | コード解説（コード + 要点 + 補足） | タイトルのみ | [slide08.md](references/slide-details/slide08.md) |
| 09 | コード + 実行結果（左右） | タイトルのみ | [slide09.md](references/slide-details/slide09.md) |
| 10 | コード + 図（吹き出し注釈） | タイトルのみ | [slide10.md](references/slide-details/slide10.md) |
| 11 | 長いコード（タイトルなし・全幅） | 白紙 | [slide11.md](references/slide-details/slide11.md) |
| 12 | 練習問題（「練習問題N」+「次ページに解答がある」枠） | タイトル+本文 | [slide12.md](references/slide-details/slide12.md) |
| 13 | 課題（「課題N」+ 課題文 + オレンジ番号の設問） | タイトル+本文 | [slide13.md](references/slide-details/slide13.md) |
| 14 | 【日英】講義表紙 | 講義表紙 | [slide14.md](references/slide-details/slide14.md) |
| 15 | 【日英】中扉 | 中扉 | [slide15.md](references/slide-details/slide15.md) |
| 16 | 【日英】本文（英語タイトル + 日英の箇条書き） | タイトル（日英）+本文 | [slide16.md](references/slide-details/slide16.md) |
| 17 | 【日英】課題（課題N / Assignment N） | タイトル（日英）+本文 | [slide17.md](references/slide-details/slide17.md) |
| 18 | 【日英】コード + 実行結果（コード・表・数式スライドの日英版の基準） | タイトル（日英）のみ | [slide18.md](references/slide-details/slide18.md) |
| 19 | 【日英】図 + 右説明（説明は日本語 18pt → 英語 14pt） | 図+説明（日英） | [slide19.md](references/slide-details/slide19.md) |

## 言語モード（着手前に選ぶ）
INIAD の講義資料には日本語のみ・日英併記・英語のみがある。**方向性の確認で必ず言語を選んでもらい**、以下に従う。

| モード | 使うスライド | ルール |
|---|---|---|
| 日本語のみ | 01〜13 | 既定。英語は用語の補足程度 |
| 日英併記 | 14〜19 を優先し、無い型は 01〜13 を「タイトル（日英）のみ」レイアウトで作って英語行を足す（18 が基準）。図 + 説明は 19 を使う | 日本語 → 直後に英語。英語は青字 #0033CC・日本語より一段小さく（本文 lvl0 24pt / lvl1 20pt / lvl2 18pt、英語タイトル 24pt）。箇条記号は日本語側だけ。英語段落は同じ字下げ（marL）・indent=0・buNone。表・図のラベルは「日本語 / English」を 1 セルに 2 行で。コード枠のラベルは「コード / Code」「実行結果 / Output」 |
| 英語のみ | 01〜13 | 文言をすべて英語に。色は日本語版と同じ（英語でも本文は黒、青字 #0033CC は使わない）。固定文言は下表で置き換える |

英語のみ・日英併記での固定文言:
| 日本語 | 英語 |
|---|---|
| 第N回 | Lecture N |
| 課題N | Assignment N |
| 練習問題N | Exercise N |
| 演習課題 | Exercises |
| 出席確認 | Attendance Check |
| 課題解説 | Assignment Explanation |
| 次ページに解答がある | Answer on the next page |
| コード / 実行結果 / コード（続き） | Code / Output / Code (cont.) |
| 参考文献 | References |

> 右上の「情報連携学部」とロゴはマスターの固定装飾なので、英語のみの資料でも変えない。

### コード関連スライド（08〜11）の使い分けと共通ルール
| 見せたいもの | 使うスライド |
|---|---|
| コードの意味・書き方を説明したい | 08 コード解説 |
| テキスト出力（print・DataFrame・配列）を見せたい | 09 コード + 実行結果 |
| グラフ・画像の出力を見せたい | 10 コード + 図 |
| 1 枚に収まらない長いコード・前スライドの続き | 11 長いコード |

- コード枠 = 白地、実行結果枠 = #F2F2F2 地（どちらも #7F7F7F 0.75pt の細枠・Consolas 18pt。長いコードのみ 16pt）
- 枠の直上に「コード」「実行結果」「コード（続き）」のラベル（16pt 太字 #595959）
- コード中の注目箇所は #C00000、枠の外の要点は赤字 #FF0000、補足説明は青字 #0070C0
- 図中の注釈は白地・青枠・青字の吹き出し
- 表の中のコード（演算子・書き方・例の列）は Consolas・**太字にしない**（`table(..., code_cols=(0, 3))`）。太字の Consolas は代替フォントで「<」などが化ける
- 1 つのコード枠は 10〜15 行まで。超える場合は 11 を続けて分割する

## 出力先（固定）
生成物はすべて、スキル直下の **`output/<デッキ名>/`** に置く（作業ディレクトリやスキル外には置かない）。デッキ名は英小文字とアンダースコア（例: `dijkstra`, `binary_search_tree`）。

| ファイル | 内容 |
|---|---|
| `output/<デッキ名>/spec.md` | 仕様書（承認を得たもの。QA での変更も反映する） |
| `output/<デッキ名>/build.py` | 生成スクリプト（冒頭で `sys.path.insert(0, os.path.join(HERE, "..", "..", "scripts"))`） |
| `output/<デッキ名>/<デッキ名>.pptx` | 完成したスライド |
| `output/<デッキ名>/shots/` | スクリーンショット |
| `output/<デッキ名>/*_demo.py` など | スクリーンショット用のコード（あれば） |

- フォルダは `deck_dir("<デッキ名>")` で作る（`shots/` も作られる）。同じ名前のデッキを作り直すときは、既存のフォルダの内容を確認してから上書きする
- 画像化などの一時ファイルは一時フォルダに出し、`output/` には残さない
- `output/` の中身は Git 管理外（講義資料・引用画像を含むため）。README だけ管理する

## 画像とスクリーンショット
INIAD の資料には Jupyter Notebook の実行画面や Web ページのスクリーンショット、写真・書影が多く使われる。画像は次の 2 通りで用意し、`add_image()` で置く。

### ユーザーが持っている画像（`user-images/`）
- ユーザーは画像をスキル直下の `user-images/`（サブフォルダ可）に入れ、「user-images の ○○.png を使って」と指示する
- AI は **着手時に `ls -R user-images` で中身を確認** し、使う画像を仕様書の該当ページに「画像: user-images/…」と書いて承認を得る
- 生成スクリプトでは `add_image(s, user_image("ai-10/svm.png"), x, y, max_w, max_h)`。`user_image()` は無いファイル名を渡すとフォルダの一覧つきでエラーにする
- ユーザーが画像を別の場所で渡した場合（チャットに貼る・パスを指定する）は、`user-images/` にコピーしてから使う

### AI が撮るスクリーンショット（`scripts/screenshot.py`）
- 準備（初回のみ）: `.venv/bin/pip install -r requirements-screenshot.txt` → `.venv/bin/python -m playwright install chromium`
- Jupyter の画面: 説明したいコードを「# %%」でセルに区切った .py を書き、`screenshot.py pycode output/<デッキ名>/demo.py -o output/<デッキ名>/shots/demo.png --cells 1-2` で **実行してから** 撮る。既存の .ipynb は `notebook work.ipynb --cells 3 [--execute]`
- Web ページ・HTML: `screenshot.py url <URL or file> -o output/<デッキ名>/shots/x.png [--selector "#main"]`
- Web ページ内の図（記事の画像など）: まず `--list` で候補（番号・alt・大きさ）を確認し、`--nth` で撮る。遅延読み込みのページは `--scroll` を付ける
  - `screenshot.py url <URL> -o - --selector "article img" --scroll --list`
  - `screenshot.py url <URL> --selector "article img" --nth 5,7,8 --scroll -o output/<デッキ名>/shots/fig{n}.png`（複数は `{n}` 入りのファイル名で一度に）
  - 他サイトの図を使うときは、スライドの右下に出典、最後に参考ページ（URL）を置く。学外配布の可否はユーザーに確認を促す
- 保存先は `output/<デッキ名>/shots/`（固定。下の「出力先」を参照）
- `--width` は既定の 800px のままでよい（notebook / pycode は 760px 未満にすると Jupyter が狭い画面用の表示に切り替わる）。横に長い出力だけ 1000px 程度に広げる
- **Jupyter のセル画像（コード + グラフ）は、横幅 7000000 EMU 以上で貼る。** 「図+説明」レイアウトの左半分（5500000）に入れるとコードが読めなくなるので、「タイトルのみ」レイアウトで画像を左に大きく、説明を右か下に置く
- コードそのものを読ませたいときは、スクリーンショットに頼らず `code_box` で書き起こす
- 撮るものは仕様書に書き、承認を得てから撮る。撮ったら画像を目で確認する（出力が意図どおりか・個人情報が写っていないか）

### 配置のルール
- `add_image(slide, path, x, y, max_w, max_h, border=True, align="center"|"left")` で、枠の中に縦横比を保って収める（引き伸ばさない）
- スクリーンショットは細いグレー枠（既定）。写真・イラストは `border=False`
- コードのスクリーンショットの上には `code_label()` で「実行画面」などのラベル、下に要点（赤字）・補足（青字）
- 画像の文字を読ませたいときは、画像を大きく（横幅 8000000 EMU 以上）して説明は下に置く。小さくしか置けないなら、コード枠（`code_box`）で書き起こすほうを選ぶ
- Web の画像・他人の図を使うときは出典を右下に小さく書く（見本 05 と同じ）

## INIAD 講義資料の典型的な構成
元資料の各回は以下の流れで組まれている。講義資料を作るときはこの型に沿うと INIAD らしくなる:
1. **講義表紙**（slide01）: 科目名・第N回タイトル・担当者・連絡先
2. **導入・理論**（slide04 → 05/06/07 の繰り返し。区切りは 02/03）: 原理は図 + 説明、数値のまとめは表、数式は自由キャンバス
3. **演習**: ノートブック扉（slide02 にノートブック名）→ コード関連スライド（08〜11）でコードと出力を順に説明
4. **練習問題**: 問題文（slide12）→ 解答（08〜11）
5. **課題**: 扉（slide02 に「課題N」）→ 課題文（slide13）
- 日英併記のときは 01/02/04/13 をそれぞれ 14/15/16/17 に、05 は 19 に、06〜11 は「タイトル（日英）のみ」レイアウト版（18 が基準）に置き換える

## 最初に: デザインの方向性を確認する（必須）
**作業着手前に、必ず `AskUserQuestion` ツールで方向性をユーザーに尋ねること。** 選択式で 1 回にまとめて聞き（1 回 4 問まで）、回答をスライドマッピング・edit に反映する。デフォルトの観点:

- **言語（必須・省略禁止）**: 「日本語のみ」「日英併記（日本語 → 英語を青字で）」「英語のみ」。上の「言語モード」に従って使うスライドと文言を決める。内容から推測して勝手に決めない
- **Copyright の年（必須・省略禁止）**: 年は資料ごとに変わるため、テンプレは `YYYY` のままになっている。選択肢は「今年（現在の日付から判断した西暦）」を先頭に、「昨年」「年度表記（例: 2025-2026）」などを並べ、それ以外は「Other」で入力してもらう。勝手に今年の年を入れて済ませない
- **ビジュアル表現の度合い**: シンプル（元資料どおり箇条書き中心） / 中間（要所で図・表・コード枠） / リッチ（図解・対比を多用）
- **情報密度**: 余白多め（lvl0 28pt 中心） / 標準 / 情報量多め（lvl1 を 18pt まで下げて詰める）
- 資料の種類（講義・演習・課題）は、内容から明らかでなければ 2 回目の `AskUserQuestion` で聞く

> 言語と年以外の観点は、コンテンツにふさわしくなければ差し替えてよいが、「選択式で着手前に方向性を確認する」流れは必ず守る。

## システム要件の確認（初回・環境が変わったとき）
**Python は必ず venv + pip で使う（システムの Python に pip install しない）。** スキルのディレクトリで初回のみ:
```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```
以降、このスキルと `/pptx` のスクリプトはすべて `.venv/bin/python scripts/xxx.py` で実行する（`markitdown` も `.venv/bin/markitdown`）。

**作業着手前に `.venv/bin/python scripts/check_requirements.py` を実行する。**
- `[必須]` に NG があれば `references/system-requirements.md` の該当セクションを読み、環境に合わせて準備を案内する（勝手に環境を書き換えない）
- `/pptx` スキル本体の有無は AI 自身が available skills で確認する
- 同一環境で一度 OK を確認できていれば毎回の実行は省略してよい

## ワークフロー（毎回の分析をスキップする）
0. **言語・Copyright の年・方向性を確認**（上記。`AskUserQuestion` で選択式に。内容のテーマや対象が曖昧なら、ここで一緒に確認する）
1. **`/pptx` スキルを呼ぶ**（QA の機構 `validate.py`・画像化をロード）。応答は引き続き日本語
2. **スライドマッピング**: 上のカタログ表 + `assets/thumbnails/` で各コンテンツに使う見本スライドを選ぶ（テンプレ分析はしない）。必要な見本の `references/slide-details/slideNN.md` だけ読む
3. **仕様書を書いて承認を得る（必須・省略禁止）**: `references/spec-template.md` の形式で、生成予定の **全ページ** について見本番号・レイアウト・載せる文言・コード・表・図の構成を書き、`output/<デッキ名>/spec.md` として保存する（出力先は固定。下の「出力先」を参照）。構成の一覧をチャットに示し、`AskUserQuestion` で「承認して生成する / 修正したい / 構成から見直す」を聞く
   - **承認されるまで生成スクリプトを書かない・実行しない**（plan モードと同じ）
   - 修正の指示があれば仕様書を直し、変更点を要約して再度承認を聞く
4. **生成スクリプトを書く**（承認済みの仕様書どおりに。仕様書のページ番号とスクリプトのコメント番号を一致させる）: `examples/deck_skeleton.py` を `output/<デッキ名>/build.py` にコピーし（`NAME`・`BILINGUAL`・`YEAR` と `sys.path` の `"..", "..", "scripts"` を書き換える）、`scripts/iniad_deck.py` の部品で中身を置く（テンプレは直接編集しない）
   ```python
   from iniad_deck import *
   d = Deck()                                   # テンプレを開き、見本スライドを削除
   s = d.new("本文日英"); title_bi(s, "タイトル", "English Title")
   fill_bi(ph(s, 1), [(0, "本文", "Body")])
   d.save(os.path.join(deck_dir("<デッキ名>"), "<デッキ名>.pptx"), year="2026")   # 保存 + Copyright 年の置換（年はユーザーの回答）
   ```
   - レイアウト名は短い名前で指定する: 講義表紙 / 中扉 / 本文 / 本文日英 / 図説明 / 図説明日英 / タイトルのみ / タイトルのみ日英 / 白紙
   - 日英の図 + 説明は `d.new("図説明日英")` + `desc_bi(ph(s, 1), [(日本語, English, 色), ...])`
   - 位置は見本スライドの `slideNN.md` の座標を使う。日英併記で英語タイトルがあるスライドは中身を `Y_BODY_BI`（コード枠は `Y_CODE_BI`）から始める
5. **実行**: `.venv/bin/python output/<デッキ名>/build.py`
6. **QA**: `/pptx` 標準 QA
   - `.venv/bin/python <pptx スキル>/scripts/office/validate.py output/<デッキ名>/<デッキ名>.pptx --original assets/iniad-template.pptx`
   - `soffice --headless --convert-to pdf --outdir <一時フォルダ> output/<デッキ名>/<デッキ名>.pptx` → `pdftoppm -r 80 -jpeg` で画像化して全スライドを目視（`/pptx` 同梱の soffice ラッパーで日本語が豆腐になる環境では soffice を直接呼ぶ）
   - `markitdown` の出力で、見本のダミー文言（「タイトル」「本文レベル1」「code_line_1()」など）と `YYYY` が残っていないか確認する
   - macOS では Meiryo UI / Consolas が無く、プレビューは代替フォントになる。字幅が変わるため、はみ出し判定は余裕を持って行う
   - 直したら生成スクリプトを直して再実行する（出力 pptx を手で直さない）
   - QA でページの分割・統合や文言の変更が必要になったら、**先に仕様書を直してから** 生成し直す
7. **報告**: `output/<デッキ名>/` のパス（スライド・仕様書）と、仕様書から変えた点（あれば）を伝える

### 中身を書くときのルール
- 箇条書きは `fill` / `fill_bi` に階層（0/1/2）を渡すだけ（`fill_bi` の英語に `None` を渡すと日本語だけの行になる。URL・固有名詞向け）。●/■ の色・サイズはマスターの bodyStyle から継承されるので、**段落に buChar や色を直書きしない**
- 色は `EMPH`（赤・要点）/ `NOTE`（青・補足）/ `EN`（日英併記の英語）/ `RED`（コード中の注目箇所、`kw()`）の定数を使う（`schemeClr` は使わない）
- **ライン・ロゴ・ヘッダ・Copyright・ページ番号をスライド側に描き足さない**（マスターにある。二重になる）
- タイトルは 1 行に収める（40pt で全角 25 文字程度まで）。長い場合は短く言い換える
- 本文は下ライン（y=6515000）より上、y≈6330000 までに収める
- 課題の設問番号は `numbered`（日英併記なら `numbered_bi`。英語行を挟むと番号がリセットされるため startAt を自動で付ける）

### python-pptx の注意点
- **継承プレースホルダの位置を変えるときは `set_geometry(shape, x, y, cx, cy)` で 4 つとも指定する。** `top` と `height` だけ代入すると `left` / `width` が 0 で書き出され、枠が消える
- `code_box` の各行は「文字列」か「`(文字列, 書式)` の組のリスト」で渡す。`["文字列"]` のように文字列 1 つをリストに入れるとエラーになる
- 表のコード列は `code_cols` で指定する（Consolas・太字なし）
- 表の行はフォントや折り返しで `row_h` より伸びることがある。表の下に置く注記は「表の上端 + 行数 × row_h」より 300000〜500000 下げて置く
- `from iniad_deck import *` で入る名前は `__all__` に限定してある。生成スクリプト側で自由に変数名を付けてよい
- 画像は `add_image()` で置く（`slide.shapes.add_picture` に幅と高さを両方渡すと縦横比が崩れる）
- `text_frame.text = ...` は書式を消すので、書式つきの文字は部品（`add_text` / `fill_bi` / `para_list`）を使う

## キレイに作る原則
- 元資料の「白地 + 水色ライン + 黒文字 + 赤●/橙■」の簡素さを保つ。色面の背景・独自色・装飾過多は避ける（INIAD らしさが消える）
- 素のタイトル＋箇条書きの単調な連続を避け、図 + 説明 / 表 / コード解説 / 中扉を織り交ぜる
- 1 枚の情報量を先に見積もる: 薄ければ図を大きく・隣と統合、濃ければ 2 枚に割る（元資料は lvl0 が 3〜4 項目程度）
- 図の 2 クラス表現は赤 #FF0000 と青 #0070C0（brand-tokens.md 参照）
- 科目によってタイトル位置や線の配置に小さな差があるが、このテンプレに統一してよい
- 最後に `/pptx` 標準 QA で仕上げ確認

## テンプレ更新時の再ベイク手順
デザイン（色・座標・見本スライド）を変えるときは、テンプレを手で直さず生成スクリプトを直して再生成する:
1. `scripts/build_template.py` 冒頭のトークンを修正（`references/brand-tokens.md` と同期）
2. ロゴ差し替え時: 白背景の新しいロゴ画像を用意し、`.venv/bin/python scripts/prepare_logo.py <新しいロゴ.jpg> assets/iniad-logo.png`（余白カット + 白背景の透過）
3. `.venv/bin/python scripts/build_template.py assets/iniad-template.pptx`（Copyright は `YYYY` のまま生成される。年を入れない）
   - 見本スライドの文言は「タイトル」「本文レベル1」などの汎用ダミーに留め、特定の講義内容の単語を入れない
4. `.venv/bin/python scripts/dump_slide_details.py assets/iniad-template.pptx references/slide-details`
5. サムネイル: `soffice --headless --convert-to pdf` → `pdftoppm -r 48 -jpeg <pdf> assets/thumbnails/slide`
6. このファイルのカタログ表を更新
