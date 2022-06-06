Attribute VB_Name = "M_Common"
Option Explicit

Public Sub セルの位置に合わせる(T_Btn As CommandButton, T_WR As Range)
    T_Btn.Left = T_WR.Left
    T_Btn.Top = T_WR.Top
    T_Btn.Width = T_WR.Width
    T_Btn.Height = T_WR.Height
End Sub

Public Sub 対象セルの書式を指定書式にする(T_WR As Range, formatStr)
    T_WR.NumberFormatLocal = formatStr
End Sub
