VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Workbook イベントモジュール.
' risk_analyzer は Workbook_Open / Workbook_BeforeClose を event_macro として検出する.

Private Sub Workbook_Open()
    ' ブックを開いた瞬間に走る. シナリオを「標準」にリセットして
    ' 設定シートの「最終更新」を現在時刻に揃える.
    On Error Resume Next
    Call MonthlyUpdate.ResetScenarioToStandard
    Me.Worksheets("設定").Range("B8").Value = Format(Now, "yyyy/mm/dd hh:nn") & " (open)"
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 閉じる直前. 何もしないが、フックの存在自体がリスクとして検出される.
    ' (Workbook_BeforeClose は黙って実行されるので「閉じるだけ」の操作が
    '  気付かないうちに副作用を引き起こす可能性がある)
End Sub
