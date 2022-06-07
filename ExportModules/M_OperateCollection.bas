Attribute VB_Name = "M_OperateCollection"
Option Explicit

'------------------------------------------------------------------------------
' �t�@�C����Collection�ɓǂݍ���
'------------------------------------------------------------------------------
Public Function CollectionFileLoad(fileNamePath As String) As Collection
    Dim lines As Collection
    Set lines = New Collection
    
    Dim fp As Long
    fp = FreeFile()
    
    Open fileNamePath For Input As #fp

    Dim buf As String
    Do While Not EOF(fp)
        If lines.Count Mod 1000 = 0 Then
            Application.StatusBar = "�ǂݍ��ݒ� ...(" & lines.Count & "�s��)"
            DoEvents
        End If
        
        Line Input #fp, buf
        lines.Add buf
    Loop
    
    Set CollectionFileLoad = lines
End Function

'------------------------------------------------------------------------------
' Collection��Data���o�͂���
'------------------------------------------------------------------------------
Public Sub CollectionPrint(ByRef lines As Collection)
    Dim i As Long
    Dim temp As String
    
    For i = 1 To lines.Count - 1
        If i Mod 1000 = 0 Then
            Application.StatusBar = "�o�͒� ...(" & i & "�s��)"
            DoEvents
        End If
    '   Debug.Print lines.Item(i)
        temp = lines.Item(i)
    Next
End Sub

