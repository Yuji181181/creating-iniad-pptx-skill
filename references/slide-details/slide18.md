# slide18: コード + 実行結果（日英）

- 使用レイアウト: INIAD タイトル（日英）のみ
- サムネイル: `assets/thumbnails/slide-18.jpg`
- 向いている内容: 「タイトル（日英）のみ」レイアウト。英語タイトルの分だけ中身を下げ、コード枠は y=1900000 から。ラベルは「コード / Code」「実行結果 / Output」、要点・補足の直後に一段小さい英語行（同じ色）。日英併記のコード・表・図・数式スライドはこの配置を基準にする。

## 編集対象の枠

| shape 名 | 種別 (ph type, idx) | サンプル文言 | 位置(off x,y EMU) | サイズ(ext cx,cy EMU) |
|---|---|---|---|---|
| Title 1 | placeholder (title, 0) | タイトル | 260000, 380000 | 11100000, 560000 |
| Text Placeholder 2 | placeholder (body, 13) | English Title | 300000, 960000 | 11000000, 450000 |
| TextBox 3 | text_box | コード / Code | 380000, 1520000 | 6000000, 340000 |
| TextBox 4 | text_box | variable = function(arg1, arg2) / print(variable) / code_lin | 380000, 1900000 | 5500000, 977250 |
| Right Arrow 5 | auto_shape |  | 6020000, 2350000 | 420000, 360000 |
| TextBox 6 | text_box | 実行結果 / Output | 6580000, 1520000 | 6000000, 340000 |
| TextBox 7 | text_box | output_1 / output_2 / output_3 | 6580000, 1900000 | 5230000, 977250 |
| TextBox 8 | text_box | 出力の読み方（赤字） / How to read the output / 補足説明（青字） / Supplementa | 380000, 3500000 | 11400000, 1800000 |

## 装飾（マスター/レイアウトから継承・編集不要）

- 上ライン + 両端ドット、「情報連携学部」、INIAD ロゴ、下ライン + 両端ドット、Copyright 表記、ページ番号はマスターにある。スライド側で描き足さない。
