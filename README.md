# Excel VBA Comparison Array and Collection.

Excel VBAにおける（動的）配列とCollectionの処理速度を比較するExcelブックです。<BR>
(Excel VBA Comparison of Collection and Dynamic Array processing speed.)

## ファイルの説明（主なもの）

- Comparison_Array_and_Collection.xlsm<BR>

実験用のExcelブック（マクロ含む）です。
Excelファイル形式のままだと、VBAのコードを直接開くことができませんので、
使用したコードを含む主なモジュールは、ExportModulesディレクトリに書き出しています。

配列またはCollectionに読み込むデータは、サンプルのcsvファイルを使用します。

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

- slist20220531.csv<BR>
実験用のサンプルファイルです。<BR>
ファイルサイズは、11,042,966 Byte（約11MByte）あります。

---

## 検証環境

|項目|値|
|--|--|
|CPU|Inten(R) Core(TM) i7-8559U CPU @ 2.70GHz|
|メモリ|16GB(LPDDR3/2133MHz)|
|OS|Windows10 Pro 21H2
|Strage|Apple APPLE SSD AP1024 SCSI Disk Device|

---

## 結果１(CPUキャッシュされず)

通常時の実行結果(配列/CollectionデータがCPUキャッシュメモリ上にない場合)

ファイルから配列およびCollectionへの読み込みについては、処理時間はほぼ変わらずでした。
イミディエイトウインドウへの書き出し（配列/Collectionからの取り出し）については、（動的）配列の方が処理時間が短い結果となりました。

### Array(Struct)
| |読み込み処理時間（合計）(s)|読み込み処理時間(s)|書き出し処理時間(s)|
|:--|:--|:--|:--|
|1回目|11.422|0.359|11.063|
|2回目|11.375|0.344|11.031|
|3回目|11.396|0.328|11.068|
|4回目|11.438|0.344|11.094|
|5回目|11.407|0.328|11.079|
|平均|11.408|0.341|11.067|

### Array(Class)
| |読み込み処理時間（合計）(s)|読み込み処理時間(s)|書き出し処理時間(s)|
|:--|:--|:--|:--|
|1回目|11.720|0.328|11.392|
|2回目|11.565|0.312|11.253|
|3回目|11.313|0.312|11.001|
|4回目|11.352|0.312|11.040|
|5回目|11.422|0.313|11.109|
|平均|11.474|0.315|11.159|


### Collection
| |読み込み処理時間（合計）(s)|読み込み処理時間(s)|書き出し処理時間(s)|
|:--|:--|:--|:--|
|1回目|15.360|0.328|15.032|
|2回目|15.343|0.328|15.015|
|3回目|15.407|0.329|15.078|
|4回目|15.267|0.329|14.938|
|5回目|15.078|0.344|14.734|
|平均|15.291|0.332|14.959|

---

## 結果２(CPUキャッシュ時)

配列/CollectionデータがCPUキャッシュメモリ上にあると思われるときの実行結果

ファイルから配列およびCollectionへの読み込みについては、処理時間はほぼ変わらずでした。
イミディエイトウインドウへの書き出し（配列/Collectionからの取り出し）については、（動的）配列の方が処理時間が短い結果となりました。

### Array(Struct)
| |読み込み処理時間（合計）(s)|読み込み処理時間(s)|書き出し処理時間(s)|
|:--|:--|:--|:--|
|1回目|0.984|0.343|0.641|
|2回目|0.953|0.328|0.625|
|3回目|0.953|0.344|0.609|
|4回目|0.969|0.344|0.625|
|5回目|0.937|0.328|0.609|
|平均|0.959|0.337|0.622|

### Array(Class)
| |読み込み処理時間（合計）(s)|読み込み処理時間(s)|書き出し処理時間(s)|
|:--|:--|:--|:--|
|1回目|0.953|0.328|0.625|
|2回目|0.938|0.313|0.625|
|3回目|0.938|0.313|0.625|
|4回目|0.953|0.328|0.625|
|5回目|0.922|0.297|0.625|
|平均|0.941|0.316|0.625|

### Collection
| |読み込み処理時間（合計）(s)|読み込み処理時間(s)|書き出し処理時間(s)|
|:--|:--|:--|:--|
|1回目|4.969|0.625|4.344|
|2回目|4.609|0.328|4.281|
|3回目|4.688|0.329|4.359|
|4回目|4.594|0.329|4.265|
|5回目|4.656|0.344|4.312|
|平均|4.703|0.391|4.312|
