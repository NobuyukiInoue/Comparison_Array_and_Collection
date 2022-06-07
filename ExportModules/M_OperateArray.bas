Attribute VB_Name = "M_OperateArray"
Option Explicit

Type ReadArray
    Item() As String
    Count As Long
    ArraySize As Long
End Type

Private Const BLOCK_SIZE As Long = 4096

'------------------------------------------------------------------------------
' ReadArray������������
'------------------------------------------------------------------------------
Public Function ArrayInit() As ReadArray
    Dim lines As ReadArray

    ReDim Preserve lines.Item(0 To BLOCK_SIZE - 1)
    lines.Count = 0
    lines.ArraySize = BLOCK_SIZE

    ArrayInit = lines
End Function

'------------------------------------------------------------------------------
' ReadArray���̔z��ɗv�f��ǉ�����
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
' ReadArray���̔z��̎w�肵���ԍ��̗v�f���폜����
'------------------------------------------------------------------------------
Public Sub RemoveItem(ByRef lines As ReadArray, index As Long)
    Dim i As Long
    
    For i = index To lines.Count - 2
        lines.Item(i) = lines.Item(i + 1)
    Next
    lines.Count = lines.Count - 1
End Sub

'------------------------------------------------------------------------------
' �t�@�C����ReadArray�ɓǂݍ���
'------------------------------------------------------------------------------
Public Function ArrayFileLoad(fileNamePath As String) As ReadArray
    Dim lines As ReadArray
    lines = ArrayInit()
    
    Dim fileNum As Long
    fileNum = FreeFile()
    
    Open fileNamePath For Input As #fileNum

    Dim buf As String
    Do While Not EOF(fileNum)
        If lines.Count Mod 1000 = 0 Then
            Application.StatusBar = "�ǂݍ��ݒ� ...(" & lines.Count & "�s��)"
            DoEvents
        End If
    
        Line Input #fileNum, buf
        AddItem lines, buf
    Loop
    
    ArrayFileLoad = lines
End Function
    
'------------------------------------------------------------------------------
' ReadArray��Data���o�͂���
'------------------------------------------------------------------------------
Public Sub ArrayPrint(ByRef lines As ReadArray)
    Dim i As Long
    Dim temp As String
    
    For i = 0 To lines.Count - 1
        If i Mod 1000 = 0 Then
            Application.StatusBar = "�o�͒� ...(" & i & "�s��)"
            DoEvents
        End If

    '   Debug.Print lines.Item(i)
        temp = lines.Item(i)
    Next
End Sub
