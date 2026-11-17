# user-images（スライドに使いたい画像の置き場所）

スライドに入れたい画像（写真・図・スクリーンショット・書影など）をこのフォルダに入れてください。
AI に「user-images の ○○.png を 3 枚目に使って」のように指示すると、スライドに配置されます。

- 対応形式: PNG / JPEG / GIF / BMP（SVG は PowerPoint に直接貼れないため PNG に変換してから）
- ファイル名は内容がわかる名前にすると指示しやすい（例: `iris_scatter.png`, `jupyter_home.png`）
- サブフォルダで講義ごとに分けてもよい（例: `user-images/ai-10/svm_margin.png`）
- このフォルダの画像は Git の管理対象外（`.gitignore`）。講義資料の画像が誤って公開されないようにするため

AI 向け: 画像を使う前に、このフォルダの一覧を確認し（`ls -R user-images`）、使う画像を仕様書の該当ページに
「画像: user-images/…」として書いてユーザーの承認を得る。配置は `iniad_deck.add_image()` を使う。
