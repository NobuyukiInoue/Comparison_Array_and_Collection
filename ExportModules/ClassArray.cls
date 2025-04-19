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

Private Item() As String
Public Count As Long
Public ArraySize As Long

Private Const BLOCK_SIZE As Long = 4096

'------------------------------------------------------------------------------
' クラスを初期化する
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    ReDim Preserve Item(0 To BLOCK_SIZE - 1)
    Count = 0
    ArraySize = BLOCK_SIZE
End Sub

'------------------------------------------------------------------------------
' Item()に値を追加する
'------------------------------------------------------------------------------
Public Sub AddItem(value As String)
    If Count >= ArraySize Then
        ArraySize = ArraySize + BLOCK_SIZE
        ReDim Preserve Item(0 To ArraySize - 1)
    End If
    
    Item(Count) = value
    Count = Count + 1
End Sub

'------------------------------------------------------------------------------
' Item(i)に値を格納する
'------------------------------------------------------------------------------
Public Function SetItem(i As Long, ByRef value As String) As String
    Item(i) = value
End Function

'------------------------------------------------------------------------------
' Item(i)の値を取得する
'------------------------------------------------------------------------------
Public Function GetItem(i As Long) As String
    GetItem = Item(i)
End Function

'------------------------------------------------------------------------------
' 指定した番号の要素を削除する
'------------------------------------------------------------------------------
Public Sub RemoveItem(index As Long)
    Dim i As Long
    
    For i = index To Count - 2
        Item(i) = Item(i + 1)
    Next
    Count = Count - 1
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
'------------------------------------------------------------------------------
Public Sub ArrayFileLoad(fileName As String, code As String, separator As String)
    With CreateObject("ADODB.Stream")
        .Charset = code
    
        Select Case separator
        Case vbLf:
            .lineSeparator = 10
        Case vbCr:
            .lineSeparator = 13
        Case Else:
            .lineSeparator = -1
        End Select
        
        .Open
        .LoadFromFile fileName
        
        Do Until .EOS
            AddItem .ReadText(-2) ' １行取り出す
        Loop
        
        .Close
    End With
End Sub
    
'------------------------------------------------------------------------------
' ファイルをReadArrayに読み込む（ファイルオープン版）
'------------------------------------------------------------------------------
Public Sub ArrayFileLoad_normal(fileNamePath As String)
    Dim fileNum As Long
    fileNum = FreeFile()
    
    Open fileNamePath For Input As #fileNum

    Dim buf As String
    Do While Not EOF(fileNum)
        If Count Mod 1000 = 0 Then
            Application.StatusBar = "読み込み中 ...(" & Count & "行目)"
            DoEvents
        End If
    
        Line Input #fileNum, buf
        AddItem (buf)
    Loop
End Sub
    
'------------------------------------------------------------------------------
' すべてのItemをイミディエイトウインドウに出力する
'------------------------------------------------------------------------------
Public Sub ArrayPrint()
    Dim i As Long
    Dim temp As String
    
    For i = 0 To Count - 1
        If i Mod 1000 = 0 Then
            Application.StatusBar = "出力中 ...(" & i & "行目)"
            DoEvents
        End If
    
    '   Debug.Print GetItem(i)
        temp = GetItem(i)
    Next
End Sub

