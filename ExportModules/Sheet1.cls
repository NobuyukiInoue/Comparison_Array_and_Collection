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
' [配列(Struct)に読み込み]ボタンをクリック
'------------------------------------------------------------------------------
Private Sub CommandButton1_Click()
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
    T_WR.CurrentRegion.Clear
    T_WR.CurrentRegion.Borders.LineStyle = xlNone
    
    ' 列見出し
    T_WR.Cells(0, 1).value = "Array(Struct)"
    T_WR.Cells(1, 2).value = "読み込み処理時間（合計）(s)"
    T_WR.Cells(1, 3).value = "読み込み処理時間(s)"
    T_WR.Cells(1, 4).value = "書き出し処理時間(s)"
    T_WR.Cells(1, 5).value = "開始時刻"
    T_WR.Cells(1, 6).value = "読み込み終了時刻"
    T_WR.Cells(1, 7).value = "書き出し終了時刻"
    
    ' 最終行に平均値を書き出す
    T_WR.Cells(2 + loopCount, 1).value = "平均"
    T_WR.Cells(2 + loopCount, 2).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 2).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 2).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 3).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 3).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 3).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 4).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 4).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 4).Address, "$", "") & ")"

    ' 書式設定
    T_WR.Range(Cells(2, 2).Address & ":" & Cells(2 + loopCount + 1, 2).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 3).Address & ":" & Cells(2 + loopCount + 1, 3).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 4).Address & ":" & Cells(2 + loopCount + 1, 4).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 5).Address & ":" & Cells(2 + loopCount, 5).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 6).Address & ":" & Cells(2 + loopCount, 6).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 7).Address & ":" & Cells(2 + loopCount, 7).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"

    ' 罫線を引く（タイトル行）
    drawBorder T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous

    ' 罫線を引く（データ行）
    drawBorder T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawHorizontal T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlDot
    
    ' 罫線を引く（平均値出力行）
    drawBorder T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous

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
        
        ' 書き出し処理時間を出力
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' 合計時間を出力
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "処理が終わりました"

End Sub

'------------------------------------------------------------------------------
' [配列(Class)に読み込み]ボタンをクリック
'------------------------------------------------------------------------------
Private Sub CommandButton2_Click()
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
    T_WR.CurrentRegion.Clear
    T_WR.CurrentRegion.Borders.LineStyle = xlNone
    
    ' 列見出し
    T_WR.Cells(0, 1).value = "Array(Class)"
    T_WR.Cells(1, 2).value = "読み込み処理時間（合計）(s)"
    T_WR.Cells(1, 3).value = "読み込み処理時間(s)"
    T_WR.Cells(1, 4).value = "書き出し処理時間(s)"
    T_WR.Cells(1, 5).value = "開始時刻"
    T_WR.Cells(1, 6).value = "読み込み終了時刻"
    T_WR.Cells(1, 7).value = "書き出し終了時刻"
    
    ' 最終行に平均値を書き出す
    T_WR.Cells(2 + loopCount, 1).value = "平均"
    T_WR.Cells(2 + loopCount, 2).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 2).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 2).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 3).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 3).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 3).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 4).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 4).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 4).Address, "$", "") & ")"

    ' 書式設定
    T_WR.Range(Cells(2, 2).Address & ":" & Cells(2 + loopCount + 1, 2).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 3).Address & ":" & Cells(2 + loopCount + 1, 3).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 4).Address & ":" & Cells(2 + loopCount + 1, 4).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 5).Address & ":" & Cells(2 + loopCount, 5).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 6).Address & ":" & Cells(2 + loopCount, 6).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 7).Address & ":" & Cells(2 + loopCount, 7).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    
    ' 罫線を引く（タイトル行）
    drawBorder T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous

    ' 罫線を引く（データ行）
    drawBorder T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawHorizontal T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlDot
    
    ' 罫線を引く（平均値出力行）
    drawBorder T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous
    
    Dim ca As ClassArray
    Dim i As Long
    For i = 1 To loopCount
        T_WR.Cells(1 + i, 1).value = i & "回目"
        Application.StatusBar = T_WR.Cells(1 + i, 1).value & "処理中..."
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 5).value = MyGetSystemTime
        
        ' ファイル読み込み
        Set ca = New ClassArray
        ca.ArrayFileLoad (fileNamePath)
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 6).value = MyGetSystemTime
        
        ' 読み込み処理時間を出力
        T_WR.Cells(1 + i, 3).value = (T_WR.Cells(1 + i, 6).value - T_WR.Cells(1 + i, 5).value) * (60 * 60 * 24#)
        
        ' イミディエイトウインドウに出力
        ca.ArrayPrint
        
        ' 現在の日付時刻取得
        T_WR.Cells(1 + i, 7).value = MyGetSystemTime
        
        ' 書き出し処理時間を出力
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' 合計時間を出力
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "処理が終わりました"

End Sub

'------------------------------------------------------------------------------
' 「コレクションに読み込み」ボタンをクリック
'------------------------------------------------------------------------------
Private Sub CommandButton3_Click()
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
    T_WR.CurrentRegion.Clear
    T_WR.CurrentRegion.Borders.LineStyle = xlNone
    
    ' 列見出し
    T_WR.Cells(0, 1).value = "Collection"
    T_WR.Cells(1, 2).value = "読み込み処理時間（合計）(s)"
    T_WR.Cells(1, 3).value = "読み込み処理時間(s)"
    T_WR.Cells(1, 4).value = "書き出し処理時間(s)"
    T_WR.Cells(1, 5).value = "開始時刻"
    T_WR.Cells(1, 6).value = "読み込み終了時刻"
    T_WR.Cells(1, 7).value = "書き出し終了時刻"
    
    ' 最終行に平均値を書き出す
    T_WR.Cells(2 + loopCount, 1).value = "平均"
    T_WR.Cells(2 + loopCount, 2).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 2).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 2).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 3).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 3).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 3).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 4).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 4).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 4).Address, "$", "") & ")"

    ' 書式設定
    T_WR.Range(Cells(2, 2).Address & ":" & Cells(2 + loopCount + 1, 2).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 3).Address & ":" & Cells(2 + loopCount + 1, 3).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 4).Address & ":" & Cells(2 + loopCount + 1, 4).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 5).Address & ":" & Cells(2 + loopCount, 5).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 6).Address & ":" & Cells(2 + loopCount, 6).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 7).Address & ":" & Cells(2 + loopCount, 7).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"

    ' 罫線を引く（タイトル行）
    drawBorder T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous

    ' 罫線を引く（データ行）
    drawBorder T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawHorizontal T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlDot
    
    ' 罫線を引く（平均値出力行）
    drawBorder T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous

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
        
        ' 書き出し処理時間を出力
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
    
    セルの位置に合わせる CommandButton1, Range("C6:C7")
    セルの位置に合わせる CommandButton2, Range("C6:C7").Offset((5 + loopCount) * 1, 0)
    セルの位置に合わせる CommandButton3, Range("C6:C7").Offset((5 + loopCount) * 2, 0)
End Sub
