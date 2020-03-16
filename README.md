# COVID19 Scraping Script for Osaka

## What's this?
大阪府の公開する情報を[大阪府 新型コロナウイルス対策サイト](https://covid19-osaka.info/)向けに整形し、
json形式にまとめ、出力するスクリプトです。

## Make date
まずは患者に関する情報のExcelファイル([例](https://github.com/codeforosaka/covid19/files/4336743/default.xlsx))を`patients.xlsx`、
検査数に関する情報のExcelファイル([例](https://github.com/codeforosaka/covid19/files/4336742/default.xlsx))を`inspections.xlsx`とリネームし、
このファイルがある階層に置きます。  
(TODO: ファイルの公開場所が決定次第、自動で取得し生成するようにする)  
その後、以下のスクリプトを実行すると`/data/data.json`が生成されます。
```shell script
pip install -r requirements.txt
python3 main.py
```

## License
このスクリプトは[MITライセンス](LICENSE)で公開されています。
