# Excel VBA Comparison Array and Collection.

Excel VBAにおける（動的）配列とCollectionの処理速度を比較です。<BR>
(Excel VBA Comparison of Collection and Dynamic Array processing speed.)

## ファイルの説明（主なもの）

- Comparison_Array_and_Collection.xlsm<BR>

実験用のExcelブック（マクロ含む）です。

Excelファイル形式のままだと、VBAのコードを直接開くことができませんので、<BR>
使用したコードを含む主なモジュールは、ExportModulesディレクトリに書き出しています。

配列またはCollectionに読み込むデータは、事前に用意したcsvファイルを使用しています。

- [ExportModules/Sheet1.cls](https://github.com/NobuyukiInoue/Comparison_Array_and_Collection/blob/main/ExportModules/Sheet1.cls)<br>
  Sheet1（操作用シート）のコード部です。<BR>
  （動的）配列およびCollectionへの読み込みおよび、イミディエイトウインドウへの書き出し処理を呼び出すボタンが配置されています。<BR>
  処理時間にはばらつきがあるため、（デフォルトでは）同じ処理を５回繰り返します。<BR>
  計測した処理時間についても、このシートに書き出されます。

- [ExportModules/M_OperateArray.bas](https://github.com/NobuyukiInoue/Comparison_Array_and_Collection/blob/main/ExportModules/M_OperateArray.bas)<BR>
（動的）配列を格納する構造体および各種操作関数群です。

- [ExportModules/ClassArray.cls](https://github.com/NobuyukiInoue/Comparison_Array_and_Collection/blob/main/ExportModules/ClassArray.cls)<BR>
上記の（動的）配列処理をクラス化したものです。

- [M_OperateCollection.bas](https://github.com/NobuyukiInoue/Comparison_Array_and_Collection/blob/main/ExportModules/M_OperateCollection.bas)<BR>
Collection使用時の読み込み・書き出し処理をまとめたものです。

- sample.csv<BR>
実験用のサンプルファイルです。<BR>
ファイルサイズは、22,085,930 Byte（約21MByte）あります。<BR>
（当初は半分の10.5MByteで実験していましたが、処理時間が短かったので同じデータを２回繰り返しています）

---

## 検証環境

|項目|値|
|--|--|
|CPU|Inten(R) Core(TM) i7-8559U CPU @ 2.70GHz|
|メモリ|16GB(LPDDR3/2133MHz)|
|OS|Windows10 Pro 21H2
|Excel|Microsoft Excel 2019|
|Strage|Apple APPLE SSD AP1024 SCSI Disk Device|

---

## 結果１（読み込み・読み出し）


配列(Struct)と配列(Class)とでは、読み込みと読み出し時間はほぼ変わらず。

配列と比較すると、Collectionは要素の取り出しがかなり遅いようです。

Collectionについては、読み込み自体は配列と同じ処理時間ですが、要素の取り出しについては配列よりかなり処理時間がかかっています。

- Array(Struct) - Load/Read

| |処理時間合計(s)|読み込み処理時間(s)|読み出し処理時間(s)|
|:----|:----|:----|:----|
|1回目|0.781|0.688|0.093|
|2回目|0.766|0.672|0.094|
|3回目|0.784|0.690|0.094|
|4回目|0.765|0.687|0.078|
|5回目|0.781|0.688|0.093|
|平均|0.775|0.685|0.090|

- Array(Class) - Load/Read

| |処理時間合計(s)|読み込み処理時間(s)|読み出し処理時間(s)|
|:----|:----|:----|:----|
|1回目|0.768|0.674|0.094|
|2回目|0.750|0.640|0.110|
|3回目|0.734|0.640|0.094|
|4回目|0.750|0.656|0.094|
|5回目|0.750|0.641|0.109|
|平均|0.750|0.650|0.100|

- Collection - Load/Read

| |処理時間合計(s)|読み込み処理時間(s)|読み出し処理時間(s)|
|:----|:----|:----|:----|
|1回目|15.392|0.703|14.689|
|2回目|15.256|0.656|14.600|
|3回目|15.277|0.656|14.621|
|4回目|15.246|0.657|14.589|
|5回目|15.695|0.640|15.055|
|平均|15.373|0.662|14.711|

---


## 結果２（読み込み・先頭の要素を削除）

先頭要素の削除については、配列(Struct)と配列(Class)でほぼ変わらず。

配列と比較すると、Collectionはかなり速くなっています。

- Array(Struct) -  Load/Remove

| |処理時間合計(s)|読み込み処理時間(s)|先頭削除処理時間(s)|
|:----|:----|:----|:----|
|1回目|0.719|0.703|0.016|
|2回目|0.687|0.671|0.016|
|3回目|0.688|0.672|0.016|
|4回目|0.706|0.690|0.016|
|5回目|0.671|0.671|0.000|
|平均|0.694|0.681|0.013|

- Array(Struct) - Load/Remove

| |処理時間合計(s)|読み込み処理時間(s)|先頭削除処理時間(s)|
|:----|:----|:----|:----|
|1回目|0.688|0.672|0.016|
|2回目|0.661|0.645|0.016|
|3回目|0.656|0.640|0.016|
|4回目|0.656|0.641|0.015|
|5回目|0.641|0.625|0.016|
|平均|0.660|0.645|0.016|

- Collection - Load/Remove

| |処理時間合計(s)|読み込み処理時間(s)|先頭削除処理時間(s)|
|:----|:----|:----|:----|
|1回目|0.690|0.690|0.000|
|2回目|0.672|0.672|0.000|
|3回目|0.656|0.656|0.000|
|4回目|0.656|0.656|0.000|
|5回目|0.657|0.657|0.000|
|平均|0.666|0.666|0.000|
