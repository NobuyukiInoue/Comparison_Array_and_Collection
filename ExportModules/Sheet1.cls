VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------------------------------
' 表の書式設定
'------------------------------------------------------------------------------
Private Sub CellsSetting(T_WR As Range, rowsCount As Long, titleStr As String, secondOperationStr As String)
    ' 表のクリア
    T_WR.CurrentRegion.Clear
    T_WR.CurrentRegion.Borders.LineStyle = xlNone
    
    ' 列見出し
    T_WR.Cells(0, 1).value = titleStr
    T_WR.Cells(1, 2).value = "処理時間合計(s)"
    T_WR.Cells(1, 3).value = "読み込み処理時間(s)"
    T_WR.Cells(1, 4).value = secondOperationStr & "処理時間(s)"
    T_WR.Cells(1, 5).value = "開始時刻"
    T_WR.Cells(1, 6).value = "読み込み終了時刻"
    T_WR.Cells(1, 7).value = secondOperationStr & "終了時刻"
    
    ' 最終行に平均値を書き出す
    T_WR.Cells(2 + rowsCount, 1).value = "平均"
    T_WR.Cells(2 + rowsCount, 2).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 2).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + rowsCount - 1, 2).Address, "$", "") & ")"
    T_WR.Cells(2 + rowsCount, 3).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 3).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + rowsCount - 1, 3).Address, "$", "") & ")"
    T_WR.Cells(2 + rowsCount, 4).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 4).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + rowsCount - 1, 4).Address, "$", "") & ")"

    ' 書式設定
    T_WR.Range(Cells(2, 2).Address & ":" & Cells(2 + rowsCount + 1, 2).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 3).Address & ":" & Cells(2 + rowsCount + 1, 3).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 4).Address & ":" & Cells(2 + rowsCount + 1, 4).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 5).Address & ":" & Cells(2 + rowsCount, 5).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 6).Address & ":" & Cells(2 + rowsCount, 6).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 7).Address & ":" & Cells(2 + rowsCount, 7).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"

    ' 罫線を引く（タイトル行）
    drawBorder T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous

    ' 罫線を引く（データ行）
    drawBorder T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + rowsCount - 1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + rowsCount - 1, 7).Address), xlContinuous
    drawHorizontal T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + rowsCount - 1, 7).Address), xlDot
    
    ' 罫線を引く（平均値出力行）
    drawBorder T_WR.Range(Cells(1 + rowsCount + 1, 1).Address & ":" & Cells(1 + rowsCount + 1, 4).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1 + rowsCount + 1, 1).Address & ":" & Cells(1 + rowsCount + 1, 4).Address), xlContinuous
End Sub

'------------------------------------------------------------------------------
' [読み込み・取り出し(配列(Struct))]ボタンをクリック
'------------------------------------------------------------------------------
Private Sub CommandButton11_Click()
    Dim filePath As String
    Dim fileName As String
    Dim fileNamePath As String
    Dim loopCount As Long

    filePath = Replace(Range("C2").value, ".", ActiveWorkbook.Path)
    fileName = Range("C3").value
    fileNamePath = filePath & "\" & fileName
    loopCount = Range("C4").value
    Application.Cursor = xlWait
    
    Dim T_WR As Range
    Set T_WR = Range("B8")
        
    CellsSetting T_WR, loopCount, "Array(Struct) - Load/Read", "読み出し"
    
    Dim lines As ReadArray
    Dim i As Long
    For i = 1 To loopCount
        T_WR.Cells(1 + i, 1).value = i & "回目"
        Application.StatusBar = T_WR.Cells(1 + i, 1).value & "処理中..."
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 5).value = MyGetSystemTime
        
        ' ファイル読み込み
        lines = ArrayFileLoad(fileNamePath)
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 6).value = MyGetSystemTime
        
        ' 読み込み処理時間を出力
        T_WR.Cells(1 + i, 3).value = (T_WR.Cells(1 + i, 6).value - T_WR.Cells(1 + i, 5).value) * (60 * 60 * 24#)
        
        ' イミディエイトウインドウに出力
        ArrayPrint lines
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 7).value = MyGetSystemTime
        
        ' 読み出し処理時間を出力
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' 合計時間を出力
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "処理が終わりました"
End Sub

'------------------------------------------------------------------------------
' 「読み込み・先頭要素を削除(配列(Struct))」ボタンをクリック
'------------------------------------------------------------------------------
Private Sub CommandButton12_Click()
    Dim filePath As String
    Dim fileName As String
    Dim fileNamePath As String
    Dim loopCount As Long

    filePath = Replace(Range("C2").value, ".", ActiveWorkbook.Path)
    fileName = Range("C3").value
    fileNamePath = filePath & "\" & fileName
    loopCount = Range("C4").value
    Application.Cursor = xlWait
    
    Dim T_WR As Range
    Set T_WR = Range("J8")

    CellsSetting T_WR, loopCount, "Array(Struct) -  Load/Remove", "先頭削除"
    
    Dim lines As ReadArray
    Dim i As Long
    For i = 1 To loopCount
        T_WR.Cells(1 + i, 1).value = i & "回目"
        Application.StatusBar = T_WR.Cells(1 + i, 1).value & "処理中..."
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 5).value = MyGetSystemTime
        
        ' ファイル読み込み
        lines = ArrayFileLoad(fileNamePath)
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 6).value = MyGetSystemTime
        
        ' 読み込み処理時間を出力
        T_WR.Cells(1 + i, 3).value = (T_WR.Cells(1 + i, 6).value - T_WR.Cells(1 + i, 5).value) * (60 * 60 * 24#)
        
        ' 先頭の要素を削除
        RemoveItem lines, 0
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 7).value = MyGetSystemTime
        
        ' 読み出し処理時間を出力
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' 合計時間を出力
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "処理が終わりました"
End Sub

'------------------------------------------------------------------------------
' [読み込み・取り出し(配列(Class))]ボタンをクリック
'------------------------------------------------------------------------------
Private Sub CommandButton21_Click()
    Dim filePath As String
    Dim fileName As String
    Dim fileNamePath As String
    Dim loopCount As Long

    filePath = Replace(Range("C2").value, ".", ActiveWorkbook.Path)
    fileName = Range("C3").value
    fileNamePath = filePath & "\" & fileName
    loopCount = Range("C4").value
    Application.Cursor = xlWait
    
    Dim T_WR As Range
    Set T_WR = Range("B18")
    
    CellsSetting T_WR, loopCount, "Array(Class) - Load/Read", "読み出し"

    Dim ca As ClassArray
    Dim i As Long
    For i = 1 To loopCount
        T_WR.Cells(1 + i, 1).value = i & "回目"
        Application.StatusBar = T_WR.Cells(1 + i, 1).value & "処理中..."
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 5).value = MyGetSystemTime
        
        ' ファイル読み込み
        Set ca = New ClassArray
        ca.ArrayFileLoad fileNamePath, "UTF-8", vbCrLf
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 6).value = MyGetSystemTime
        
        ' 読み込み処理時間を出力
        T_WR.Cells(1 + i, 3).value = (T_WR.Cells(1 + i, 6).value - T_WR.Cells(1 + i, 5).value) * (60 * 60 * 24#)
        
        ' イミディエイトウインドウに出力
        ca.ArrayPrint
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 7).value = MyGetSystemTime
        
        ' 読み出し処理時間を出力
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' 合計時間を出力
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "処理が終わりました"

End Sub

'------------------------------------------------------------------------------
' 「読み込み・先頭要素を削除(配列(Class))」ボタンをクリック
'------------------------------------------------------------------------------
Private Sub CommandButton22_Click()
    Dim filePath As String
    Dim fileName As String
    Dim fileNamePath As String
    Dim loopCount As Long

    filePath = Replace(Range("C2").value, ".", ActiveWorkbook.Path)
    fileName = Range("C3").value
    fileNamePath = filePath & "\" & fileName
    loopCount = Range("C4").value
    Application.Cursor = xlWait
    
    Dim T_WR As Range
    Set T_WR = Range("J18")

    CellsSetting T_WR, loopCount, "Array(Struct) - Load/Remove", "先頭削除"
    
    Dim ca As ClassArray
    Dim i As Long
    For i = 1 To loopCount
        T_WR.Cells(1 + i, 1).value = i & "回目"
        Application.StatusBar = T_WR.Cells(1 + i, 1).value & "処理中..."
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 5).value = MyGetSystemTime
        
        ' ファイル読み込み
        Set ca = New ClassArray
        ca.ArrayFileLoad fileNamePath, "UTF-8", vbCrLf
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 6).value = MyGetSystemTime
        
        ' 読み込み処理時間を出力
        T_WR.Cells(1 + i, 3).value = (T_WR.Cells(1 + i, 6).value - T_WR.Cells(1 + i, 5).value) * (60 * 60 * 24#)
        
        ' 先頭の要素を削除
        ca.RemoveItem (0)
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 7).value = MyGetSystemTime
        
        ' 読み出し処理時間を出力
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' 合計時間を出力
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "処理が終わりました"
End Sub

'------------------------------------------------------------------------------
' 「読み込み・取り出し(Collection)」ボタンをクリック
'------------------------------------------------------------------------------
Private Sub CommandButton31_Click()
    Dim filePath As String
    Dim fileName As String
    Dim fileNamePath As String
    Dim loopCount As Long

    filePath = Replace(Range("C2").value, ".", ActiveWorkbook.Path)
    fileName = Range("C3").value
    fileNamePath = filePath & "\" & fileName
    loopCount = Range("C4").value
    Application.Cursor = xlWait
    
    Dim T_WR As Range
    Set T_WR = Range("B28")
    
    CellsSetting T_WR, loopCount, "Collection - Load/Read", "読み出し"
    
    Dim lines As Collection
    Dim i As Long
    For i = 1 To loopCount
        T_WR.Cells(1 + i, 1).value = i & "回目"
        Application.StatusBar = T_WR.Cells(1 + i, 1).value & "処理中..."
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 5).value = MyGetSystemTime
        
        ' ファイル読み込み
        Set lines = CollectionFileLoad(fileNamePath)
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 6).value = MyGetSystemTime
        
        ' 読み込み処理時間を出力
        T_WR.Cells(1 + i, 3).value = (T_WR.Cells(1 + i, 6).value - T_WR.Cells(1 + i, 5).value) * (60 * 60 * 24#)
        
        ' イミディエイトウインドウに出力
        CollectionPrint lines
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 7).value = MyGetSystemTime
        
        ' 読み出し処理時間を出力
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' 合計時間を出力
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "処理が終わりました"

End Sub

'------------------------------------------------------------------------------
' 「読み込み・先頭要素を削除(Collection)」ボタンをクリック
'------------------------------------------------------------------------------
Private Sub CommandButton32_Click()
    Dim filePath As String
    Dim fileName As String
    Dim fileNamePath As String
    Dim loopCount As Long

    filePath = Replace(Range("C2").value, ".", ActiveWorkbook.Path)
    fileName = Range("C3").value
    fileNamePath = filePath & "\" & fileName
    loopCount = Range("C4").value
    Application.Cursor = xlWait
    
    Dim T_WR As Range
    Set T_WR = Range("J28")
    
    CellsSetting T_WR, loopCount, "Collection - Load/Remove", "先頭削除"
    
    Dim lines As Collection
    Dim i As Long
    For i = 1 To loopCount
        T_WR.Cells(1 + i, 1).value = i & "回目"
        Application.StatusBar = T_WR.Cells(1 + i, 1).value & "処理中..."
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 5).value = MyGetSystemTime
        
        ' ファイル読み込み
        Set lines = CollectionFileLoad(fileNamePath)
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 6).value = MyGetSystemTime
        
        ' 読み込み処理時間を出力
        T_WR.Cells(1 + i, 3).value = (T_WR.Cells(1 + i, 6).value - T_WR.Cells(1 + i, 5).value) * (60 * 60 * 24#)
        
        ' 先頭の要素を削除
        lines.Remove (1)
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 7).value = MyGetSystemTime
        
        ' 読み出し処理時間を出力
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' 合計時間を出力
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "処理が終わりました"
End Sub

'------------------------------------------------------------------------------
' シートがアクティブになったとき
'------------------------------------------------------------------------------
Private Sub Worksheet_Activate()
    Dim loopCount As Long
    loopCount = Range("C4").value
    
    セルの位置に合わせる CommandButton11, Range("C6:D7")
    セルの位置に合わせる CommandButton21, Range("C6:D7").Offset((5 + loopCount) * 1, 0)
    セルの位置に合わせる CommandButton31, Range("C6:D7").Offset((5 + loopCount) * 2, 0)

    セルの位置に合わせる CommandButton12, Range("K6:L7")
    セルの位置に合わせる CommandButton22, Range("K6:L7").Offset((5 + loopCount) * 1, 0)
    セルの位置に合わせる CommandButton32, Range("K6:L7").Offset((5 + loopCount) * 2, 0)

End Sub
