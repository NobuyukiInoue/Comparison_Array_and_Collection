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
' [�z��(Struct)�ɓǂݍ���]�{�^�����N���b�N
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
    
    ' �񌩏o��
    T_WR.Cells(0, 1).value = "Array(Struct)"
    T_WR.Cells(1, 2).value = "�ǂݍ��ݏ������ԁi���v�j(s)"
    T_WR.Cells(1, 3).value = "�ǂݍ��ݏ�������(s)"
    T_WR.Cells(1, 4).value = "�����o����������(s)"
    T_WR.Cells(1, 5).value = "�J�n����"
    T_WR.Cells(1, 6).value = "�ǂݍ��ݏI������"
    T_WR.Cells(1, 7).value = "�����o���I������"
    
    ' �ŏI�s�ɕ��ϒl�������o��
    T_WR.Cells(2 + loopCount, 1).value = "����"
    T_WR.Cells(2 + loopCount, 2).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 2).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 2).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 3).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 3).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 3).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 4).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 4).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 4).Address, "$", "") & ")"

    ' �����ݒ�
    T_WR.Range(Cells(2, 2).Address & ":" & Cells(2 + loopCount + 1, 2).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 3).Address & ":" & Cells(2 + loopCount + 1, 3).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 4).Address & ":" & Cells(2 + loopCount + 1, 4).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 5).Address & ":" & Cells(2 + loopCount, 5).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 6).Address & ":" & Cells(2 + loopCount, 6).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 7).Address & ":" & Cells(2 + loopCount, 7).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"

    ' �r���������i�^�C�g���s�j
    drawBorder T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous

    ' �r���������i�f�[�^�s�j
    drawBorder T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawHorizontal T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlDot
    
    ' �r���������i���ϒl�o�͍s�j
    drawBorder T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous

    Dim lines As ReadArray
    Dim i As Long
    For i = 1 To loopCount
        T_WR.Cells(1 + i, 1).value = i & "���"
        Application.StatusBar = T_WR.Cells(1 + i, 1).value & "������..."
        
        ' ���݂̓��t�����擾
        T_WR.Cells(1 + i, 5).value = MyGetSystemTime
        
        ' �t�@�C���ǂݍ���
        lines = ArrayFileLoad(fileNamePath)
        
        ' ���݂̓��t�����擾
        T_WR.Cells(1 + i, 6).value = MyGetSystemTime
        
        ' �ǂݍ��ݏ������Ԃ��o��
        T_WR.Cells(1 + i, 3).value = (T_WR.Cells(1 + i, 6).value - T_WR.Cells(1 + i, 5).value) * (60 * 60 * 24#)
        
        ' �C�~�f�B�G�C�g�E�C���h�E�ɏo��
        ArrayPrint lines
        
        ' ���݂̓��t�����擾
        T_WR.Cells(1 + i, 7).value = MyGetSystemTime
        
        ' �����o���������Ԃ��o��
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' ���v���Ԃ��o��
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "�������I���܂���"

End Sub

'------------------------------------------------------------------------------
' [�z��(Class)�ɓǂݍ���]�{�^�����N���b�N
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
    
    ' �񌩏o��
    T_WR.Cells(0, 1).value = "Array(Class)"
    T_WR.Cells(1, 2).value = "�ǂݍ��ݏ������ԁi���v�j(s)"
    T_WR.Cells(1, 3).value = "�ǂݍ��ݏ�������(s)"
    T_WR.Cells(1, 4).value = "�����o����������(s)"
    T_WR.Cells(1, 5).value = "�J�n����"
    T_WR.Cells(1, 6).value = "�ǂݍ��ݏI������"
    T_WR.Cells(1, 7).value = "�����o���I������"
    
    ' �ŏI�s�ɕ��ϒl�������o��
    T_WR.Cells(2 + loopCount, 1).value = "����"
    T_WR.Cells(2 + loopCount, 2).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 2).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 2).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 3).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 3).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 3).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 4).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 4).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 4).Address, "$", "") & ")"

    ' �����ݒ�
    T_WR.Range(Cells(2, 2).Address & ":" & Cells(2 + loopCount + 1, 2).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 3).Address & ":" & Cells(2 + loopCount + 1, 3).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 4).Address & ":" & Cells(2 + loopCount + 1, 4).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 5).Address & ":" & Cells(2 + loopCount, 5).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 6).Address & ":" & Cells(2 + loopCount, 6).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 7).Address & ":" & Cells(2 + loopCount, 7).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    
    ' �r���������i�^�C�g���s�j
    drawBorder T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous

    ' �r���������i�f�[�^�s�j
    drawBorder T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawHorizontal T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlDot
    
    ' �r���������i���ϒl�o�͍s�j
    drawBorder T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous
    
    Dim ca As ClassArray
    Dim i As Long
    For i = 1 To loopCount
        T_WR.Cells(1 + i, 1).value = i & "���"
        Application.StatusBar = T_WR.Cells(1 + i, 1).value & "������..."
        
        ' ���݂̓��t�����擾
        T_WR.Cells(1 + i, 5).value = MyGetSystemTime
        
        ' �t�@�C���ǂݍ���
        Set ca = New ClassArray
        ca.ArrayFileLoad (fileNamePath)
        
        ' ���݂̓��t�����擾
        T_WR.Cells(1 + i, 6).value = MyGetSystemTime
        
        ' �ǂݍ��ݏ������Ԃ��o��
        T_WR.Cells(1 + i, 3).value = (T_WR.Cells(1 + i, 6).value - T_WR.Cells(1 + i, 5).value) * (60 * 60 * 24#)
        
        ' �C�~�f�B�G�C�g�E�C���h�E�ɏo��
        ca.ArrayPrint
        
        ' ���݂̓��t�����擾
        T_WR.Cells(1 + i, 7).value = MyGetSystemTime
        
        ' �����o���������Ԃ��o��
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' ���v���Ԃ��o��
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "�������I���܂���"

End Sub

'------------------------------------------------------------------------------
' �u�R���N�V�����ɓǂݍ��݁v�{�^�����N���b�N
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
    
    ' �񌩏o��
    T_WR.Cells(0, 1).value = "Collection"
    T_WR.Cells(1, 2).value = "�ǂݍ��ݏ������ԁi���v�j(s)"
    T_WR.Cells(1, 3).value = "�ǂݍ��ݏ�������(s)"
    T_WR.Cells(1, 4).value = "�����o����������(s)"
    T_WR.Cells(1, 5).value = "�J�n����"
    T_WR.Cells(1, 6).value = "�ǂݍ��ݏI������"
    T_WR.Cells(1, 7).value = "�����o���I������"
    
    ' �ŏI�s�ɕ��ϒl�������o��
    T_WR.Cells(2 + loopCount, 1).value = "����"
    T_WR.Cells(2 + loopCount, 2).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 2).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 2).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 3).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 3).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 3).Address, "$", "") & ")"
    T_WR.Cells(2 + loopCount, 4).Formula = "=AVERAGE(" & Replace(T_WR.Cells(2, 4).Address, "$", "") & ":" & Replace(T_WR.Cells(2 + loopCount - 1, 4).Address, "$", "") & ")"

    ' �����ݒ�
    T_WR.Range(Cells(2, 2).Address & ":" & Cells(2 + loopCount + 1, 2).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 3).Address & ":" & Cells(2 + loopCount + 1, 3).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 4).Address & ":" & Cells(2 + loopCount + 1, 4).Address).NumberFormat = "0.000"
    T_WR.Range(Cells(2, 5).Address & ":" & Cells(2 + loopCount, 5).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 6).Address & ":" & Cells(2 + loopCount, 6).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
    T_WR.Range(Cells(2, 7).Address & ":" & Cells(2 + loopCount, 7).Address).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"

    ' �r���������i�^�C�g���s�j
    drawBorder T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1, 1).Address & ":" & Cells(1, 7).Address), xlContinuous

    ' �r���������i�f�[�^�s�j
    drawBorder T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawVertical T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlContinuous
    drawHorizontal T_WR.Range(Cells(2, 1).Address & ":" & Cells(2 + loopCount - 1, 7).Address), xlDot
    
    ' �r���������i���ϒl�o�͍s�j
    drawBorder T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous
    drawVertical T_WR.Range(Cells(1 + loopCount + 1, 1).Address & ":" & Cells(1 + loopCount + 1, 4).Address), xlContinuous

    Dim lines As Collection
    Dim i As Long
    For i = 1 To loopCount
        T_WR.Cells(1 + i, 1).value = i & "���"
        Application.StatusBar = T_WR.Cells(1 + i, 1).value & "������..."
        
        ' ���݂̓��t�����擾
        T_WR.Cells(1 + i, 5).value = MyGetSystemTime
        
        ' �t�@�C���ǂݍ���
        Set lines = CollectionFileLoad(fileNamePath)
        
        ' ���݂̓��t�����擾
        T_WR.Cells(1 + i, 6).value = MyGetSystemTime
        
        ' �ǂݍ��ݏ������Ԃ��o��
        T_WR.Cells(1 + i, 3).value = (T_WR.Cells(1 + i, 6).value - T_WR.Cells(1 + i, 5).value) * (60 * 60 * 24#)
        
        ' �C�~�f�B�G�C�g�E�C���h�E�ɏo��
        CollectionPrint lines
        
        ' ���݂̓��t�����擾
        T_WR.Cells(1 + i, 7).value = MyGetSystemTime
        
        ' �����o���������Ԃ��o��
        T_WR.Cells(1 + i, 4).value = (T_WR.Cells(1 + i, 7).value - T_WR.Cells(1 + i, 6).value) * (60 * 60 * 24#)
        
        ' ���v���Ԃ��o��
        T_WR.Cells(1 + i, 2).value = T_WR.Cells(1 + i, 3).value + T_WR.Cells(1 + i, 4).value
    
    Next

    Application.StatusBar = False
    Application.Cursor = xlDefault
    
    MsgBox "�������I���܂���"

End Sub

'------------------------------------------------------------------------------
' �V�[�g���A�N�e�B�u�ɂȂ����Ƃ�
'------------------------------------------------------------------------------
Private Sub Worksheet_Activate()
    Dim loopCount As Long
    loopCount = Range("C4").value
    
    �Z���̈ʒu�ɍ��킹�� CommandButton1, Range("C6:C7")
    �Z���̈ʒu�ɍ��킹�� CommandButton2, Range("C6:C7").Offset((5 + loopCount) * 1, 0)
    �Z���̈ʒu�ɍ��킹�� CommandButton3, Range("C6:C7").Offset((5 + loopCount) * 2, 0)
End Sub