# slide17: 課題（日英）

- 使用レイアウト: INIAD タイトル（日英）+本文
- サムネイル: `assets/thumbnails/slide-17.jpg`
- 向いている内容: タイトル「課題N」+ 英語タイトル「Assignment N」。課題文・設問それぞれの直後に英語行。英語行を挟むと自動番号がリセットされるので、2 問目以降は `startAt` で番号を指定する。

## 編集対象の枠

| shape 名 | 種別 (ph type, idx) | サンプル文言 | 位置(off x,y EMU) | サイズ(ext cx,cy EMU) |
|---|---|---|---|---|
| Title 1 | placeholder (title, 0) | 課題N | 260000, 380000 | 11100000, 560000 |
| Text Placeholder 2 | placeholder (body, 13) | Assignment N | 300000, 960000 | 11000000, 450000 |
| Content Placeholder 3 | placeholder (object, 1) | 課題文（使用するデータの説明と、求めること） / Assignment statement (data descript | 380000, 1560000 | 11430000, 4770000 |

## 装飾（マスター/レイアウトから継承・編集不要）

- 上ライン + 両端ドット、「情報連携学部」、INIAD ロゴ、下ライン + 両端ドット、Copyright 表記、ページ番号はマスターにある。スライド側で描き足さない。
