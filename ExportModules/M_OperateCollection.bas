Attribute VB_Name = "M_OperateCollection"
Option Explicit

'------------------------------------------------------------------------------
' ファイルをCollectionに読み込む
'------------------------------------------------------------------------------
Public Function CollectionFileLoad(fileNamePath As String) As Collection
    Dim lines As Collection
    Set lines = New Collection
    
    Dim fp As Long
    fp = FreeFile()
    
    Open fileNamePath For Input As #fp

    Dim buf As String
    Do While Not EOF(fp)
        Line Input #fp, buf

        If lines.Count Mod 1000 = 0 Then
            Application.StatusBar = "読み込み中 ...(" & lines.Count & "行目)"
            DoEvents
        End If
        
        lines.Add buf
    Loop
    
    Set CollectionFileLoad = lines
End Function
    
'------------------------------------------------------------------------------
' CollectionのDataを出力する
'------------------------------------------------------------------------------
Public Sub CollectionPrint(ByRef lines As Collection)
    Dim i As Long
    
    For i = 1 To lines.Count
        If i Mod 1000 = 0 Then
            Application.StatusBar = "出力中 ...(" & i & "行目)"
            DoEvents
        End If
        Debug.Print lines.Item(i)
    Next
End Sub

