# slide13: 課題

- 使用レイアウト: INIAD タイトル+本文
- サムネイル: `assets/thumbnails/slide-13.jpg`
- 向いている内容: タイトルは「課題N」（授業名・科目コードは入れない）。lvl0 ● に課題文（データの説明と求めること）、lvl1 に設問をオレンジの自動番号（1. 2. 3.、`buAutoNum arabicPeriod` + buClr #FF9900）で。

## 編集対象の枠

| shape 名 | 種別 (ph type, idx) | サンプル文言 | 位置(off x,y EMU) | サイズ(ext cx,cy EMU) |
|---|---|---|---|---|
| Title 1 | placeholder (title, 0) | 課題N | 260000, 380000 | 11100000, 560000 |
| Content Placeholder 2 | placeholder (object, 1) | 課題文（使用するデータの説明と、求めること） / 設問1 / 設問2 / 設問3 | 380000, 1180000 | 11430000, 5150000 |

## 装飾（マスター/レイアウトから継承・編集不要）

- 上ライン + 両端ドット、「情報連携学部」、INIAD ロゴ、下ライン + 両端ドット、Copyright 表記、ページ番号はマスターにある。スライド側で描き足さない。
