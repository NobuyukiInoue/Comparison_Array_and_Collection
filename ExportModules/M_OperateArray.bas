Attribute VB_Name = "M_OperateArray"
Option Explicit

Type ReadArray
    Data() As String
    Count As Long
    ArraySize As Long
End Type

Private Const BLOCK_SIZE As Long = 4096

'------------------------------------------------------------------------------
' ReadArrayを初期化する
'------------------------------------------------------------------------------
Public Function ArrayInit() As ReadArray
    Dim lines As ReadArray

    ReDim Preserve lines.Data(0 To BLOCK_SIZE - 1)
    lines.Count = 0
    lines.ArraySize = BLOCK_SIZE

    ArrayInit = lines
End Function

'------------------------------------------------------------------------------
' ファイルをReadArrayに読み込む
'------------------------------------------------------------------------------
Public Function ArrayFileLoad(fileNamePath As String) As ReadArray
    Dim lines As ReadArray
    lines = ArrayInit()
    
    Dim fp As Long
    fp = FreeFile()
    
    Open fileNamePath For Input As #fp

    Dim buf As String
    Do While Not EOF(fp)
        Line Input #fp, buf
                
        If lines.Count >= lines.ArraySize Then
            lines.ArraySize = lines.ArraySize + BLOCK_SIZE
            ReDim Preserve lines.Data(0 To lines.ArraySize)
        End If

        If lines.Count Mod 1000 = 0 Then
            Application.StatusBar = "読み込み中 ...(" & lines.Count & "行目)"
            DoEvents
        End If
        
        lines.Data(lines.Count) = buf
        lines.Count = lines.Count + 1
    Loop
    
    ArrayFileLoad = lines
End Function
    
'------------------------------------------------------------------------------
' ReadArrayのDataを出力する
'------------------------------------------------------------------------------
Public Sub ArrayPrint(ByRef lines As ReadArray)
    Dim i As Long
    For i = 0 To lines.Count
        If i Mod 1000 = 0 Then
            Application.StatusBar = "出力中 ...(" & i & "行目)"
            DoEvents
        End If
        Debug.Print lines.Data(i)
    Next
End Sub
