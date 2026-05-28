<#
.SYNOPSIS
    scripts/make_sample.py が出力した .xlsx に VBA を注入して .xlsm を作る.

.DESCRIPTION
    openpyxl は VBA プロジェクトを新規に作れないため、Windows + Excel の
    COM オートメーション経由で VBA を埋め込む。

    動作要件:
      - Windows
      - Microsoft Excel がインストールされていること
      - Excel の信頼センターで「VBA プロジェクト オブジェクト モデルへの
        アクセスを信頼する」が有効になっていること
        (ファイル → オプション → セキュリティセンター → マクロの設定)

    入力:
      scripts/vba/Module1_MonthlyUpdate.bas
      scripts/vba/Module2_Helpers.bas
      scripts/vba/ThisWorkbook.cls
      scripts/vba/Sheet_Settings.cls          (「設定」シートのコードに挿入)
      frontend/public/samples/retail_monthly_ops.xlsx (元データ)

    出力:
      frontend/public/samples/retail_monthly_ops.xlsm

    使い方:
      pwsh -File scripts/inject_vba.ps1
      または
      powershell -File scripts\inject_vba.ps1

.NOTES
    OSS リポジトリではこのスクリプトで生成した .xlsm を git にコミット
    してしまうのが最も楽 (ユーザーが Excel を持っていなくてもサンプルを
    試せる)。再生成が必要なときに開発者がこのスクリプトを叩く。
#>

$ErrorActionPreference = 'Stop'

$repoRoot   = Resolve-Path (Join-Path $PSScriptRoot '..')
$xlsxPath   = Join-Path $repoRoot 'frontend\public\samples\retail_monthly_ops.xlsx'
$xlsmPath   = Join-Path $repoRoot 'frontend\public\samples\retail_monthly_ops.xlsm'
$vbaDir     = Join-Path $PSScriptRoot 'vba'

if (-not (Test-Path $xlsxPath)) {
    Write-Error "Source xlsx not found: $xlsxPath`nRun 'uv run python scripts/make_sample.py' first."
}
foreach ($n in @('Module1_MonthlyUpdate.bas', 'Module2_Helpers.bas', 'ThisWorkbook.cls', 'Sheet_Settings.cls')) {
    if (-not (Test-Path (Join-Path $vbaDir $n))) {
        Write-Error "Missing VBA source: $n in $vbaDir"
    }
}

# ----- Excel 起動 -----------------------------------------------------------
Write-Host 'Launching Excel COM...'
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    Write-Error "Failed to start Excel COM. Is Microsoft Excel installed? `n$_"
}
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening $xlsxPath"
    $wb = $excel.Workbooks.Open($xlsxPath)

    # VBA プロジェクトへのアクセス確認
    try {
        $vbproj = $wb.VBProject
    } catch {
        throw @"
Cannot access VBProject. Enable in Excel:
  File → Options → Trust Center → Trust Center Settings
   → Macro Settings → 'Trust access to the VBA project object model'
"@
    }

    # ----- 標準モジュールを Import (Module1, Module2) ----------------------
    foreach ($basName in @('Module1_MonthlyUpdate.bas', 'Module2_Helpers.bas')) {
        $basPath = Join-Path $vbaDir $basName
        Write-Host "  Importing $basName"
        [void]$vbproj.VBComponents.Import($basPath)
    }

    # ----- ThisWorkbook イベントモジュール: コード本体を貼り付け -----------
    # ThisWorkbook は既に存在する組み込みクラスなので Import ではなく
    # CodeModule に直接書き込む。
    Write-Host '  Writing ThisWorkbook events'
    $tw = $vbproj.VBComponents.Item('ThisWorkbook')
    # 既存コードを全消去 (空のはず)
    $tw.CodeModule.DeleteLines(1, [Math]::Max(1, $tw.CodeModule.CountOfLines))
    # ヘッダ (VERSION / Attribute) は CodeModule に書かない. 純粋な VBA コードだけ.
    $twSource = Get-Content (Join-Path $vbaDir 'ThisWorkbook.cls') -Raw -Encoding UTF8
    $twCode = ($twSource -split "Option Explicit", 2)[1]
    $twCode = "Option Explicit" + $twCode
    $tw.CodeModule.AddFromString($twCode.Trim())

    # ----- 「設定」シートの Worksheet_Change を CodeModule に書き込む ------
    Write-Host '  Writing 設定 sheet code'
    $settingsSheet = $null
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq '設定') { $settingsSheet = $s; break }
    }
    if ($settingsSheet -eq $null) {
        throw "Sheet '設定' not found in workbook"
    }
    $codeName = $settingsSheet.CodeName
    $sheetComp = $vbproj.VBComponents.Item($codeName)
    $sheetComp.CodeModule.DeleteLines(1, [Math]::Max(1, $sheetComp.CodeModule.CountOfLines))
    $sheetSource = Get-Content (Join-Path $vbaDir 'Sheet_Settings.cls') -Raw -Encoding UTF8
    $sheetCode = ($sheetSource -split "Option Explicit", 2)[1]
    $sheetCode = "Option Explicit" + $sheetCode
    $sheetComp.CodeModule.AddFromString($sheetCode.Trim())

    # ----- ダッシュボードシートにフォームコントロール (ボタン) を追加 -------
    Write-Host '  Adding form button on ダッシュボード'
    $dash = $null
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq 'ダッシュボード') { $dash = $s; break }
    }
    if ($dash -ne $null) {
        # ボタンの位置 (px). 概ね A26 セル付近.
        $left   = 10
        $top    = 540
        $width  = 140
        $height = 28
        # msoFormControl ボタンを追加してマクロを割り当てる
        # (Shapes.AddFormControl 第1引数: xlButtonControl = 0)
        $btn = $dash.Buttons().Add($left, $top, $width, $height)
        $btn.Caption = '月次更新を実行'
        $btn.OnAction = 'UpdateMonthlyDashboard'
    }

    # ----- xlsm として保存 -------------------------------------------------
    Write-Host "Saving as $xlsmPath"
    # 既存 xlsm があれば削除 (上書き)
    if (Test-Path $xlsmPath) { Remove-Item $xlsmPath -Force }
    # xlOpenXMLWorkbookMacroEnabled = 52
    $wb.SaveAs($xlsmPath, 52)
    $wb.Close($false)
    Write-Host "Done. Output: $xlsmPath"
    $info = Get-Item $xlsmPath
    Write-Host "Size: $($info.Length) bytes"
} finally {
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
