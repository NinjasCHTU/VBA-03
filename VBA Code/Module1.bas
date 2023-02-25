Attribute VB_Name = "Module1"
Sub TurnBlack()
Attribute TurnBlack.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' TurnBlack Macro
'
' Keyboard Shortcut: Ctrl+b
'
    With selection.Font
        .colorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub
