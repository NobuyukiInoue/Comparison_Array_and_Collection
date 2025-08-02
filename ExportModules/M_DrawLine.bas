Attribute VB_Name = "M_DrawLine"
Option Explicit

' xlContinuous  実線
' xlDot 破線

'------------------------------------------------------------------------------
' 外枠を引く
'------------------------------------------------------------------------------
Sub drawBorder(T_WR As Range, argLineStyle As Long)
    With T_WR.Borders(xlEdgeLeft)
        .LineStyle = argLineStyle
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With T_WR.Borders(xlEdgeTop)
        .LineStyle = argLineStyle
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With T_WR.Borders(xlEdgeBottom)
        .LineStyle = argLineStyle
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With T_WR.Borders(xlEdgeRight)
        .LineStyle = argLineStyle
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

'------------------------------------------------------------------------------
' 縦線を引く
'------------------------------------------------------------------------------
Sub drawVertical(T_WR As Range, argLineStyle As Long)
    With T_WR.Borders(xlInsideVertical)
        .LineStyle = argLineStyle
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

'------------------------------------------------------------------------------
' 横線を引く
'------------------------------------------------------------------------------
Sub drawHorizontal(T_WR As Range, argLineStyle As Long)
    With T_WR.Borders(xlInsideHorizontal)
        .LineStyle = argLineStyle
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub


