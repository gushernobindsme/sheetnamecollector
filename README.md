SheetNameCollector
============

## 1.概要
エクセルファイルのシート名集計ツール。  

指定したフォルダからエクセルファイルを再帰的に検索し、シート名を集計してテキストファイルに出力します。

## 2.使い方
### (1)jarファイル作成
ソースをチェックアウトし、

    mvn package

コマンドでjarファイルを作成してください。

### (2)バッチファイル作成
適用な.batファイルを作成し、

    java -jar sheetnamecollector-0.0.1-SNAPSHOT-jar-with-dependencies.jar ${検索対象ディレクトリ} ${検索結果の出力先ディレクトリ} 
    pause

と記載し、同ディレクトリにjarファイルを配置してください。

### (3)バッチファイルを実行
バッチファイルを実行すると、指定したディレクトリに「result.csv」が作成されます。  
すでに作成済みの場合、ファイルは上書きされます。  
