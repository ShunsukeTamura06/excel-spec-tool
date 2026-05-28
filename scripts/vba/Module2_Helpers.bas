Attribute VB_Name = "Helpers"
Option Explicit

' 共通ユーティリティ. 副作用なしの純粋関数群.
' risk_analyzer は意図的に何も検出しないことを期待 (= 健全な参考実装としての存在).

Public Function FormatJPY(ByVal amount As Double) As String
    ' 通貨整形. 1,234,567 形式.
    FormatJPY = Format(amount, "#,##0") & " 円"
End Function

Public Function IsBusinessDay(ByVal d As Date) As Boolean
    ' 土日を除外. 祝日判定は省略 (実プロジェクトでは祝日マスタ参照).
    Dim w As Integer
    w = Weekday(d, vbMonday)
    IsBusinessDay = (w <= 5)
End Function

Public Function MonthEndDate(ByVal anyDateInMonth As Date) As Date
    ' 月末日を返す. DateSerial(年, 月+1, 0) で月末.
    MonthEndDate = DateSerial(Year(anyDateInMonth), Month(anyDateInMonth) + 1, 0)
End Function

Public Function ClampPercent(ByVal v As Double) As Double
    ' 0.0〜1.0 にクランプ.
    If v < 0 Then
        ClampPercent = 0
    ElseIf v > 1 Then
        ClampPercent = 1
    Else
        ClampPercent = v
    End If
End Function

Public Function SafeDivide(ByVal numerator As Double, ByVal denominator As Double) As Double
    ' ゼロ除算ガード.
    If denominator = 0 Then
        SafeDivide = 0
    Else
        SafeDivide = numerator / denominator
    End If
End Function
