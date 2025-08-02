Attribute VB_Name = "M_OperateCollection"
Option Explicit

'------------------------------------------------------------------------------
' ファイルをCollectionに読み込む(ADODB版)
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
'------------------------------------------------------------------------------
Public Function CollectionFileLoad(fileName As String, code As String, separator As String) As Collection
    Dim lines As Collection
    Set lines = New Collection
    
    With CreateObject("ADODB.Stream")
        .Charset = code
    
        Select Case separator
        Case vbLf:
            .lineseparator = 10
        Case vbCr:
            .lineseparator = 13
        Case Else:
            .lineseparator = -1
        End Select
        
        .Open
        .LoadFromFile fileName
        
        Do Until .EOS
            lines.Add .ReadText(-2) ' １行取り出す
        Loop
        
        .Close
    End With

    Set CollectionFileLoad = lines
End Function

'------------------------------------------------------------------------------
' ファイルをCollectionに読み込む（ファイルオープン版）
'------------------------------------------------------------------------------
Public Function CollectionFileLoad_normal(fileNamePath As String) As Collection
    Dim lines As Collection
    Set lines = New Collection
    
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
        lines.Add buf
    Loop
    
    Set CollectionFileLoad = lines
End Function

'------------------------------------------------------------------------------
' CollectionのDataを出力する
'------------------------------------------------------------------------------
Public Sub CollectionPrint(ByRef lines As Collection)
    Dim i As Long
    Dim temp As String
    
    For i = 1 To lines.Count - 1
        If i Mod 1000 = 0 Then
            Application.StatusBar = "出力中 ...(" & i & "行目)"
            DoEvents
        End If
    '   Debug.Print lines.Item(i)
        temp = lines.Item(i)
    Next
End Sub


