Attribute VB_Name = "M_OperateArray"
Option Explicit

Type ReadArray
    Item() As String
    Count As Long
    ArraySize As Long
End Type

Private Const BLOCK_SIZE As Long = 4096

'------------------------------------------------------------------------------
' ReadArrayを初期化する
'------------------------------------------------------------------------------
Public Function ArrayInit() As ReadArray
    Dim lines As ReadArray

    ReDim Preserve lines.Item(0 To BLOCK_SIZE - 1)
    lines.Count = 0
    lines.ArraySize = BLOCK_SIZE

    ArrayInit = lines
End Function

'------------------------------------------------------------------------------
' ReadArray内の配列に要素を追加する
'------------------------------------------------------------------------------
Public Sub AddItem(ByRef lines As ReadArray, ByRef value As String)
    If lines.Count >= lines.ArraySize Then
        lines.ArraySize = lines.ArraySize + BLOCK_SIZE
        ReDim Preserve lines.Item(0 To lines.ArraySize)
    End If

    lines.Item(lines.Count) = value
    lines.Count = lines.Count + 1
End Sub

'------------------------------------------------------------------------------
' ReadArray内の配列の指定した番号の要素を削除する
'------------------------------------------------------------------------------
Public Sub RemoveItem(ByRef lines As ReadArray, index As Long)
    Dim i As Long
    
    For i = index To lines.Count - 2
        lines.Item(i) = lines.Item(i + 1)
    Next
    lines.Count = lines.Count - 1
End Sub

'------------------------------------------------------------------------------
' ファイルをReadArrayに読み込む(ADODB版)
'
' ☆code(.Chrset)の値
' "SJIS", "UTF-8"
'
' ☆separator(.LineSeparator)の値
' -------+---+-------------------------------
' 定数    値  説明
' -------+---+-------------------------------
' adCR    13  復帰を示します。
' adCRLF  -1  既定値。復帰改行を示します。
' adLF    10  改行を示します。
' -------+-----------------------------------
'
' ☆ .ReadText (NumChars):
' Stream オブジェクトから指定したバイト数または文字数のデータを読み取ります｡
' NumChars : 読み取るバイト数を (Long型) で指定します。（もしくは、以下のEnumを指定）
'
' -----------+---+------------------------------------------------
' 定数        値  説明
' -----------+---+------------------------------------------------
' adReadAll   -1  既定値。現在の位置から EOS マーカー方向に、すべてのバイトをストリームから読み取ります。
' adReadLine  -2  ストリームから次の行を読み取ります (LineSeparator プロパティで指定)。
' -----------+---+------------------------------------------------
'
' ※Enum を使用するには、ツールの参照設定で
' Microsoft ActiveX Data Objects 6.1 Library にチェックを入れる必要あり。
'------------------------------------------------------------------------------
Public Function ArrayFileLoad(fileName As String, code As String, separator As String) As ReadArray
    Dim lines As ReadArray
    lines = ArrayInit()
    
    With CreateObject("ADODB.Stream")
        .Charset = code
    
        Select Case separator
        Case vbLf:
            .LineSeparator = 10
        Case vbCr:
            .LineSeparator = 13
        Case Else:
            .LineSeparator = -1
        End Select
        
        .Open
        .LoadFromFile fileName
        
        Do Until .EOS
            AddItem lines, .ReadText(-2) ' １行取り出す
        Loop
        
        .Close
    End With

    ArrayFileLoad = lines
End Function
    
'------------------------------------------------------------------------------
' ファイルをReadArrayに読み込む（ファイルオープン版）
'------------------------------------------------------------------------------
Public Function ArrayFileLoad_normal(fileNamePath As String) As ReadArray
    Dim lines As ReadArray
    lines = ArrayInit()
    
    Dim fileNum As Long
    fileNum = FreeFile()
    
    Open fileNamePath For Input As #fileNum

    Dim buf As String
    Do While Not EOF(fileNum)
        If lines.Count Mod 1000 = 0 Then
            Application.StatusBar = "読み込み中 ...(" & lines.Count & "行目)"
            DoEvents
        End If
    
        Line Input #fileNum, buf
        AddItem lines, buf
    Loop
    
    ArrayFileLoad = lines
End Function
    
'------------------------------------------------------------------------------
' ReadArrayのDataを出力する
'------------------------------------------------------------------------------
Public Sub ArrayPrint(ByRef lines As ReadArray)
    Dim i As Long
    Dim temp As String
    
    For i = 0 To lines.Count - 1
        If i Mod 1000 = 0 Then
            Application.StatusBar = "出力中 ...(" & i & "行目)"
            DoEvents
        End If

    '   Debug.Print lines.Item(i)
        temp = lines.Item(i)
    Next
End Sub

