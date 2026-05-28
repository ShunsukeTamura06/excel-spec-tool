Attribute VB_Name = "MonthlyUpdate"
Option Explicit

' 月次更新ボタンの処理.
' ダッシュボードシートの「月次更新」フォームコントロールから呼び出される。
'
' Excelツール改修支援AI の risk_analyzer が以下を検出することを期待:
'   - ActiveSheet     → runtime_state リスク
'   - Worksheets(変数).Range(変数) → dynamic_vba リスク
'   - Range(変数)     → dynamic_vba リスク

Public Sub UpdateMonthlyDashboard()
    Dim targetSheetName As String
    Dim ws As Worksheet
    Dim lastRow As Long

    ' シナリオ名を取得して、対応する計算シート名を組み立てる (実行時に決まる参照)
    targetSheetName = "シナリオ計算"

    ' Worksheets(変数) は静的解析では追跡できない
    Set ws = ThisWorkbook.Worksheets(targetSheetName)

    ' Range(動的範囲文字列) も同様
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Range("F6:F" & lastRow).Calculate

    ' ActiveSheet は実行時の選択状態に依存 (runtime_state)
    Call WriteLastUpdated(ActiveSheet)

    ' 計算結果を Dashboard に転記
    Call PushKpisToDashboard(ws, lastRow)

    MsgBox "月次ダッシュボードを更新しました。", vbInformation, "完了"
End Sub

' --- 内部 --------------------------------------------------------------------

Private Sub WriteLastUpdated(ByVal targetSheet As Worksheet)
    ' 設定シートの「最終更新」欄に現在時刻を書く
    ThisWorkbook.Worksheets("設定").Range("B8").Value = Format(Now, "yyyy/mm/dd hh:nn")
End Sub

Private Sub PushKpisToDashboard(ByVal calcSheet As Worksheet, ByVal lastRow As Long)
    Dim dash As Worksheet
    Set dash = ThisWorkbook.Worksheets("ダッシュボード")

    ' 計算済の合計値を Dashboard の隠し領域に書き戻す (再計算回避)
    dash.Range("G3").Value = WorksheetFunction.Sum(calcSheet.Range("E6:E" & lastRow))
    dash.Range("G4").Value = Now
End Sub

Public Sub ResetScenarioToStandard()
    ' 開発・テスト用. 設定シートのシナリオを「標準」に戻す.
    ThisWorkbook.Worksheets("設定").Range("B7").Value = "標準"
End Sub
