Attribute VB_Name = "Module6"
Sub Test03()
    Dim myRange As range
    Set myRange = range("G61:G63")
    
    With myRange.Borders
        .LineStyle = xlContinuous
        .colorIndex = xlAutomatic
        .
    End With
End Sub
Sub BorderLook()
Attribute BorderLook.VB_Description = "look at code border"
Attribute BorderLook.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' BorderLook Macro
' look at code border
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    ActiveWindow.Zoom = 55
    ActiveWindow.Zoom = 85
    ActiveWindow.SmallScroll Down:=2
    selection.Borders(xlDiagonalDown).LineStyle = xlNone
    selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    selection.Borders(xlInsideVertical).LineStyle = xlNone
    selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveWindow.SmallScroll Down:=9
    range("C60:C61").Select
    selection.Borders(xlDiagonalDown).LineStyle = xlNone
    selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    selection.Borders(xlInsideVertical).LineStyle = xlNone
    selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    range("C60").Select
    selection.Borders(xlDiagonalDown).LineStyle = xlNone
    selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlDash
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    selection.Borders(xlInsideVertical).LineStyle = xlNone
    selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    range("E62").Select
    ActiveWindow.SmallScroll Down:=2
    range("C63:C65").Select
    selection.Borders(xlDiagonalDown).LineStyle = xlNone
    selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    selection.Borders(xlInsideVertical).LineStyle = xlNone
    selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    range("C64").Select
    selection.Borders(xlDiagonalDown).LineStyle = xlNone
    selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlDash
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlDash
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    selection.Borders(xlInsideVertical).LineStyle = xlNone
    selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveWindow.SmallScroll Down:=10
    range("C67:C70").Select
    ActiveWindow.SmallScroll Down:=-10
    range("E57:E60").Select
    selection.Borders(xlDiagonalDown).LineStyle = xlNone
    selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    selection.Borders(xlInsideVertical).LineStyle = xlNone
    selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    range("E58").Select
    selection.Borders(xlDiagonalDown).LineStyle = xlNone
    selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlDash
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlDash
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    selection.Borders(xlInsideVertical).LineStyle = xlNone
    selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    range("E59").Select
    selection.Borders(xlDiagonalDown).LineStyle = xlNone
    selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlDash
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlDash
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    selection.Borders(xlInsideVertical).LineStyle = xlNone
    selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    range("G59").Select
End Sub
