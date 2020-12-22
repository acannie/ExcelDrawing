# Drawing_on_Excel

画像を読み込み、Excel のセルの色でその画像を再現するプログラムです。

詳細は以下の Qiita の記事にまとめました。

[画像を読み込んでExcel上でセルの背景色で描画するプログラム作ってみた](https://qiita.com/acannie/items/1e516944da65dda47259)

# 使い方

(読み込みたい画像ファイルの名前が `image.jpg` の場合)

1. 
    `image.jpg` を `figure/` 下に置きます。
    
2. 
    下記コマンドを実行します。

    ```
    $ python3 main.py image.jpg
    ```
3. 
    `result/` 下に同名の .xlsx ファイル (`image.xlsx`) が出力されます。

## 注意点

入力ファイルは .jpg 形式のみ対応しております。
.png は失敗する可能性があります。