Attribute VB_Name = "M_DrawLine"
Option Explicit

' xlContinuous  é¿ê¸
' xlDot îjê¸

'------------------------------------------------------------------------------
' äOògÇà¯Ç≠
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
' ècê¸Çà¯Ç≠
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
' â°ê¸Çà¯Ç≠
'------------------------------------------------------------------------------
Sub drawHorizontal(T_WR As Range, argLineStyle As Long)
    With T_WR.Borders(xlInsideHorizontal)
        .LineStyle = argLineStyle
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

