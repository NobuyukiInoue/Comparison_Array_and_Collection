Attribute VB_Name = "M_DrawLine"
Option Explicit

' xlContinuous  ����
' xlDot �j��

'------------------------------------------------------------------------------
' �O�g������
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
' �c��������
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
' ����������
'------------------------------------------------------------------------------
Sub drawHorizontal(T_WR As Range, argLineStyle As Long)
    With T_WR.Borders(xlInsideHorizontal)
        .LineStyle = argLineStyle
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

