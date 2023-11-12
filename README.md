# XLCELL-PICKER
* 作成日：2023.11.11
* 作成者：kawasaki8108
* Python 3.10.9
* そのほかのモジュールは以下の通り(requirements.txtで定義)
```
altgraph==0.17.4
et-xmlfile==1.1.0
numpy==1.26.1
openpyxl==3.1.2
packaging==23.2
pandas==2.1.3
pefile==2023.2.7
pip==23.3.1
pyinstaller==6.2.0
pyinstaller-hooks-contrib==2023.10
python-dateutil==2.8.2
pytz==2023.3.post1
pywin32-ctypes==0.2.2
setuptools==68.2.2
six==1.16.0
tzdata==2023.3
```

## 概要（使用シーン）
* 同じフォーマットの複数のExcelファイルに対して、特定のセルから値を一括抽出・一覧化したいときを想定して作成しています。
* 一覧化では、デフォルトで読み込んだファイル名とそのハッシュ値をSHA1を出力します。
* 出力ファイルはExcel.xlsx形式で以下の形となります。user定義カラム名1~nがuserで定義する列です。

|(index)|ファイル名|SHA1ハッシュ値|user定義カラム名1|user定義カラム名n|
|:---|:---|:---|:---|:---|
|1|Excelファイル1|4b890444dca5a14d76c1908ab6a143a65dc71be0|●●●●|■■■■|
|n|Excelファイルn|db459148c65fb09315ba279090c445bf20218aa7|○○○○|□□□□|

## 使用前の準備
* 「.exeファイル」と「定義ファイル.xlsx」をローカルにダウンロード(「<>code」>「Download ZIP」からまるごとDLでok)
* 任意の場所に上記2点を同階層で格納
* 定義ファイルを開き黄色セルに必要情報入力・上書き保存
* └読み込むシート名、出力するファイルのファイル名、カラム名、セル番地を定義

## つかいかた
* .exeファイルをWクリックして起動。（ポップアップが生じた場合は必要に応じて実行を認める）
* input：以下の階層で親フォルダを読込先ディレクトリとして指定する
```
親フォルダ
    ├── Excelファイル1.xlsx
    ├── Excelファイル2.xlsx
    └── Excelファイル3.xlsx
```
* output：任意のディレクトリ
* outputはPythonでデータフレームから調整なしでExcelに書き出していますので、罫線や列幅調整は出力後user側で必要に応じて実施ください。
  
### 以上
