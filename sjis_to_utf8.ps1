$CURRENT_PATH = Get-Location
$TARGET_PATH = $CURRENT_PATH.Path + "/ExportModules"

Get-ChildItem -Path $TARGET_PATH -Filter *.cls | ForEach-Object {
    $filename = $TARGET_PATH + "/" + $_.Name
    $content = Get-Content $_.FullName
#   $content = $content -replace "`r`n", "`n"  # 改行コードをLFに変換
    [System.IO.File]::WriteAllText($filename, $content, [System.Text.Encoding]::UTF8)
}

Get-ChildItem -Path $TARGET_PATH -Filter *.bas | ForEach-Object {
    $filename = $TARGET_PATH + "/" + $_.Name
    $content = Get-Content $_.FullName
#   $content = $content -replace "`r`n", "`n"  # 改行コードをLFに変換
    [System.IO.File]::WriteAllText($filename, $content, [System.Text.Encoding]::UTF8)
}
