Attribute VB_Name = "M_ExportModules"
Option Explicit

'--------------------------------------------------------------------------------------------------
' ���Z�L�����e�B�Z���^�[�̐ݒ�
' 1-1. Microsoft Office �{�^�����N���b�N���A[Excel �̃I�v�V����] ���N���b�N���܂��B
' 1-2. [�Z�L�����e�B �Z���^�[] ���N���b�N���܂��
' 1-3. [�Z�L�����e�B �Z���^�[�̐ݒ�] ���N���b�N���܂��
' 1-4. [�}�N���̐ݒ�] ���N���b�N���܂��
' 1-5. [VBA �v���W�F�N�g �I�u�W�F�N�g ���f���ւ̃A�N�Z�X��M������] �`�F�b�N �{�b�N�X���I���ɂ��܂��B
' 1-6. [OK] ���N���b�N���� [Excel �̃I�v�V����] �_�C�A���O �{�b�N�X����܂��B
'
' ��VBE�I�u�W�F�N�g�̎g�p����
' VBA�v���O�����ɂ���ă��[�N�u�b�N��VBA�R�[�h��ύX���邽�߂ɂ�, Application.VBE�I�u�W�F�N�g���g�p���܂��
' VBE�I�u�W�F�N�g���g�p����ɂ�, ����2�̏������K�v�ł��
' 2-1. VBE�ɂ�����Microsoft Visual Basic for Applications Extensibility�ւ̎Q�Ƃ�ǉ�����B
' 2-2.�uVBA�v���W�F�N�g �I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������v�I�v�V�������w�肷��B
'
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' �S���W���[��(VBA�R�[�h)�̃G�N�X�|�[�g
'--------------------------------------------------------------------------------------------------
Public Sub ExportAll()
    Dim module                  As VBComponent      '// ���W���[��
    Dim moduleList              As VBComponents     '// VBA�v���W�F�N�g�̑S���W���[��
    Dim extension                                   '// ���W���[���̊g���q
    Dim sPath As String                             '// �����Ώۃu�b�N�̃p�X
    Dim sFilePath                                   '// �G�N�X�|�[�g�t�@�C���p�X
    Dim TargetBook As Workbook                      '// �����Ώۃu�b�N�I�u�W�F�N�g
    Dim Count As Long

    If Workbooks.Count > 1 Then
        MsgBox "���[�N�u�b�N���Q�ȏ�J����Ă��܂��B", vbOKOnly, "�G���["
        Exit Sub
    End If

    Dim targetPath As String

    '------------------------------------------------------
    ' �t�H���_�̑I���_�C�A���O���J��
    '------------------------------------------------------
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = True
        .InitialFileName = ActiveWorkbook.Path
        .Title = "�G�N�X�|�[�g��̃t�H���_��I��"

        If .Show = True Then
            targetPath = .SelectedItems(1)
        End If
    End With

    If targetPath = "" Then

        ' �t�H���_���I������Ȃ������Ƃ�
        Exit Sub

    End If

    Set TargetBook = ActiveWorkbook
    sPath = ActiveWorkbook.Path

    If Dir(targetPath, vbDirectory) = "" Then
        MsgBox targetPath & " �����݂��܂���B", vbOKOnly, "�G���["
        Exit Sub
    End If

    '// �����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = TargetBook.VBProject.VBComponents

    '// VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
    For Each module In moduleList

        If (module.Type = vbext_ct_ClassModule) Then
            '// �N���X
            extension = "cls"

        ElseIf (module.Type = vbext_ct_MSForm) Then
            '// �t�H�[��
            '// .frx���ꏏ�ɃG�N�X�|�[�g�����
            extension = "frm"

        ElseIf (module.Type = vbext_ct_StdModule) Then
            '// �W�����W���[��
            extension = "bas"

        ElseIf (module.Type = vbext_ct_Document) Then
            '// �h�L�������g�i�V�[�g�j
            extension = "cls"

        ElseIf (module.Type = vbext_ct_ActiveXDesigner) Then
            '// ActiveX�f�U�C�i
            '// �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo CONTINUE

        Else
            '// ���̑�
            '// �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo CONTINUE

        End If

        '// �G�N�X�|�[�g���{
        sFilePath = targetPath & "\" & module.Name & "." & extension
        Application.StatusBar = sFilePath & " ���G�N�X�|�[�g��..."

        Call module.Export(sFilePath)
        Count = Count + 1

        '// �o�͐�m�F�p���O�o��
        Debug.Print sFilePath

CONTINUE:
    Next

    Application.StatusBar = False

    MsgBox "�S���W���[���̃G�N�X�|�[�g���I���܂���" & vbCrLf _
        & vbCrLf _
        & "�o�̓t�@�C���� = " & Count _
        , vbOKOnly, "�G�N�X�|�[�g����"

End Sub

