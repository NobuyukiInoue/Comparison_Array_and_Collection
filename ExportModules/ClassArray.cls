VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Data() As String
Public Count As Long
Public ArraySize As Long

Private Const BLOCK_SIZE As Long = 4096

'------------------------------------------------------------------------------
' �N���X������������
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    ReDim Preserve Data(0 To BLOCK_SIZE - 1)
    Count = 0
    ArraySize = BLOCK_SIZE
End Sub

'------------------------------------------------------------------------------
' Data()�ɒl��ǉ�����
'------------------------------------------------------------------------------
Public Sub AddData(value As String)
    If Count >= ArraySize Then
        ArraySize = ArraySize + BLOCK_SIZE
        ReDim Preserve Data(0 To ArraySize)
    End If
    
    Data(Count) = value
    Count = Count + 1
End Sub

'------------------------------------------------------------------------------
' Data(i)�ɒl���i�[����
'------------------------------------------------------------------------------
Public Function SetData(i As Long, value As String) As String
    Data(i) = value
End Function

'------------------------------------------------------------------------------
' Data(i)�̒l���擾����
'------------------------------------------------------------------------------
Public Function GetData(i As Long) As String
    GetData = Data(i)
End Function
    
'------------------------------------------------------------------------------
' �t�@�C����ReadArray�ɓǂݍ���
'------------------------------------------------------------------------------
Public Sub ArrayFileLoad(fileNamePath As String)
    Dim fp As Long
    fp = FreeFile()
    
    Open fileNamePath For Input As #fp

    Dim buf As String
    Do While Not EOF(fp)
        Line Input #fp, buf
        
        AddData (buf)
        
        If Count Mod 1000 = 0 Then
            Application.StatusBar = "�ǂݍ��ݒ� ...(" & Count & "�s��)"
            DoEvents
        End If
    Loop
End Sub
    
'------------------------------------------------------------------------------
' Data���o�͂���
'------------------------------------------------------------------------------
Public Sub ArrayPrint()
    Dim i As Long
    For i = 0 To Count
        If i Mod 1000 = 0 Then
            Application.StatusBar = "�o�͒� ...(" & i & "�s��)"
            DoEvents
        End If
        Debug.Print GetData(i)
    Next
End Sub
