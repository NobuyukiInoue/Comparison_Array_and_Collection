Attribute VB_Name = "M_OperateDictionary"
Option Explicit

'------------------------------------------------------------------------------
' ファイルをDictionaryに読み込む(ADODB版)
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
Public Function DictionaryFileLoad(fileName As String, code As String, separator As String) As Object
    Dim lines As Object
    Set lines = CreateObject("Scripting.Dictionary")
    
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
        
        Dim i As Long
        i = 1
        
        Do Until .EOS
            lines.Add i, .ReadText(-2) ' １行取り出す
            i = i + 1
        Loop
        
        .Close
    End With

    Set DictionaryFileLoad = lines
End Function

'------------------------------------------------------------------------------
' ファイルをDictionaryに読み込む（ファイルオープン版）
'------------------------------------------------------------------------------
Public Function DictionaryFileLoad_normal(fileNamePath As String) As Object
    Dim lines As Object
    Set lines = CreateObject("Scripting.Dictionary")

    Dim fileNum As Long
    fileNum = FreeFile()
    
    Open fileNamePath For Input As #fileNum

    Dim buf As String
    Dim i As Long
    
    Do While Not EOF(fileNum)
        If lines.Count Mod 1000 = 0 Then
            Application.StatusBar = "読み込み中 ...(" & lines.Count & "行目)"
            DoEvents
        End If
        
        Line Input #fileNum, buf
        lines.Add i, buf
        i = i + 1
    Loop
    
    Set DictionaryFileLoad = lines
End Function

'------------------------------------------------------------------------------
' DictionaryのDataを出力する
'------------------------------------------------------------------------------
Public Sub DictionaryPrint(ByRef lines As Object)
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


