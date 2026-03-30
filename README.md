## 1. エクセルファイルから、時間割のjsonファイルを作る。

* プログラム：excel2json.py
* 作成ファイル：docs/schedule.json, docs/info.json

excel2json.pyで、エクセルファイル名と、各学年のシート名を指定してあります。
「2026年(1年前期)」ようなシートのみ読み込みます。


## 2. jsonファイルを読み込んでホームページを表示する
* ファイル：docs/index.html

内部ではjQueryを使ってます。
