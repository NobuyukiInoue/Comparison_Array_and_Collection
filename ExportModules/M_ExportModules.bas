Attribute VB_Name = "M_ExportModules"
Option Explicit

'--------------------------------------------------------------------------------------------------
' ☆セキュリティセンターの設定
' 1-1. Microsoft Office ボタンをクリックし、[Excel のオプション] をクリックします。
' 1-2. [セキュリティ センター] をクリックします｡
' 1-3. [セキュリティ センターの設定] をクリックします｡
' 1-4. [マクロの設定] をクリックします｡
' 1-5. [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する] チェック ボックスをオンにします。
' 1-6. [OK] をクリックして [Excel のオプション] ダイアログ ボックスを閉じます。
'
' ☆VBEオブジェクトの使用準備
' VBAプログラムによってワークブックのVBAコードを変更するためには, Application.VBEオブジェクトを使用します｡
' VBEオブジェクトを使用するには, 次の2つの準備が必要です｡
' 2-1. VBEにおいてMicrosoft Visual Basic for Applications Extensibilityへの参照を追加する。
' 2-2.「VBAプロジェクト オブジェクトモデルへのアクセスを信頼する」オプションを指定する。
'
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' 全モジュール(VBAコード)のエクスポート
'--------------------------------------------------------------------------------------------------
Public Sub ExportAll()
    Dim module                  As VBComponent      '// モジュール
    Dim moduleList              As VBComponents     '// VBAプロジェクトの全モジュール
    Dim extension                                   '// モジュールの拡張子
    Dim sPath As String                             '// 処理対象ブックのパス
    Dim sFilePath                                   '// エクスポートファイルパス
    Dim TargetBook As Workbook                      '// 処理対象ブックオブジェクト
    Dim Count As Long

    If Workbooks.Count > 1 Then
        MsgBox "ワークブックが２つ以上開かれています。", vbOKOnly, "エラー"
        Exit Sub
    End If

    Dim targetPath As String

    '------------------------------------------------------
    ' フォルダの選択ダイアログを開く
    '------------------------------------------------------
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = True
        .InitialFileName = ActiveWorkbook.Path
        .Title = "エクスポート先のフォルダを選択"

        If .Show = True Then
            targetPath = .SelectedItems(1)
        End If
    End With

    If targetPath = "" Then

        ' フォルダが選択されなかったとき
        Exit Sub

    End If

    Set TargetBook = ActiveWorkbook
    sPath = ActiveWorkbook.Path

    If Dir(targetPath, vbDirectory) = "" Then
        MsgBox targetPath & " が存在しません。", vbOKOnly, "エラー"
        Exit Sub
    End If

    '// 処理対象ブックのモジュール一覧を取得
    Set moduleList = TargetBook.VBProject.VBComponents

    '// VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList

        If (module.Type = vbext_ct_ClassModule) Then
            '// クラス
            extension = "cls"

        ElseIf (module.Type = vbext_ct_MSForm) Then
            '// フォーム
            '// .frxも一緒にエクスポートされる
            extension = "frm"

        ElseIf (module.Type = vbext_ct_StdModule) Then
            '// 標準モジュール
            extension = "bas"

        ElseIf (module.Type = vbext_ct_Document) Then
            '// ドキュメント（シート）
            extension = "cls"

        ElseIf (module.Type = vbext_ct_ActiveXDesigner) Then
            '// ActiveXデザイナ
            '// エクスポート対象外のため次ループへ
            GoTo CONTINUE

        Else
            '// その他
            '// エクスポート対象外のため次ループへ
            GoTo CONTINUE

        End If

        '// エクスポート実施
        sFilePath = targetPath & "\" & module.Name & "." & extension
        Application.StatusBar = sFilePath & " をエクスポート中..."

        Call module.Export(sFilePath)
        Count = Count + 1

        '// 出力先確認用ログ出力
        Debug.Print sFilePath

CONTINUE:
    Next

    Application.StatusBar = False

    MsgBox "全モジュールのエクスポートが終わりました" & vbCrLf _
        & vbCrLf _
        & "出力ファイル数 = " & Count _
        , vbOKOnly, "エクスポート完了"

End Sub

