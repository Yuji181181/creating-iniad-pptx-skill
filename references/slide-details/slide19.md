# slide19: 図 + 右説明（日英）

- 使用レイアウト: INIAD 図+説明（日英）
- サムネイル: `assets/thumbnails/slide-19.jpg`
- 向いている内容: 「図+説明（日英）」レイアウト。英語タイトル欄つき、図は左半分（y=1650000 から）、右の説明は日本語 18pt → 英語 14pt（青）で英語行は日本語に詰める。「見出し / Heading:」は太字。iniad_deck の desc_bi で埋める。

## 編集対象の枠

| shape 名 | 種別 (ph type, idx) | サンプル文言 | 位置(off x,y EMU) | サイズ(ext cx,cy EMU) |
|---|---|---|---|---|
| Title 1 | placeholder (title, 0) | タイトル | 260000, 380000 | 11100000, 560000 |
| Text Placeholder 2 | placeholder (body, 13) | English Title | 300000, 960000 | 11000000, 450000 |
| Content Placeholder 3 | placeholder (object, 1) | 図の説明文 / Description of the figure / 見出し / Heading: 要点（赤字） /  | 6250000, 1560000 | 5560000, 4750000 |
| Rectangle 4 | auto_shape | 図・グラフ領域 / （画像を貼る） | 380000, 1650000 | 5500000, 4550000 |

## 装飾（マスター/レイアウトから継承・編集不要）

- 上ライン + 両端ドット、「情報連携学部」、INIAD ロゴ、下ライン + 両端ドット、Copyright 表記、ページ番号はマスターにある。スライド側で描き足さない。
