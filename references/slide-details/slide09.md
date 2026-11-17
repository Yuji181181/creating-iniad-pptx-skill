# slide09: コード + 実行結果（左右）

- 使用レイアウト: INIAD タイトルのみ
- サムネイル: `assets/thumbnails/slide-09.jpg`
- 向いている内容: 左にコード枠（白地）、矢印、右に実行結果枠（#F2F2F2 地）。下に出力の読み方（赤字）と補足（青字）。print や DataFrame の出力に。

## 編集対象の枠

| shape 名 | 種別 (ph type, idx) | サンプル文言 | 位置(off x,y EMU) | サイズ(ext cx,cy EMU) |
|---|---|---|---|---|
| Title 1 | placeholder (title, 0) | タイトル | 260000, 380000 | 11100000, 560000 |
| TextBox 2 | text_box | コード | 380000, 1220000 | 6000000, 340000 |
| TextBox 3 | text_box | variable = function(arg1, arg2) / print(variable) / code_lin | 380000, 1600000 | 5500000, 1263000 |
| Right Arrow 4 | auto_shape |  | 6020000, 2100000 | 420000, 360000 |
| TextBox 5 | text_box | 実行結果 | 6580000, 1220000 | 6000000, 340000 |
| TextBox 6 | text_box | output_1 / output_2 / output_3 / output_4 | 6580000, 1600000 | 5230000, 1263000 |
| TextBox 7 | text_box | 出力の読み方（赤字） / 補足説明（青字） | 380000, 3500000 | 11400000, 1800000 |

## 装飾（マスター/レイアウトから継承・編集不要）

- 上ライン + 両端ドット、「情報連携学部」、INIAD ロゴ、下ライン + 両端ドット、Copyright 表記、ページ番号はマスターにある。スライド側で描き足さない。
