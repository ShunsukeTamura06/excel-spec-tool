VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 「設定」シートのモジュール (Worksheet_Change イベント).
' risk_analyzer は event_macro として検出する.
'
' inject_vba.ps1 は対象シートの CodeName を取得して
' このコードを差し込む (シート名「設定」とは別)。

Private Sub Worksheet_Change(ByVal Target As Range)
    ' シナリオを変更したら、シナリオ計算シートを再計算する.
    ' 静的解析では「Target が B7 だったときだけ走る」までは追えるが、
    ' その先で何が起きるかは実行時にしか分からない (= 未解析リスク).
    If Target.Address = "$B$7" Then
        Application.EnableEvents = False
        ThisWorkbook.Worksheets("シナリオ計算").Calculate
        Application.EnableEvents = True
    End If
End Sub
