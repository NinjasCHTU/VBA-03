Attribute VB_Name = "Module2"
Sub TurnWhite()
Attribute TurnWhite.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' TurnWhite Macro
'
' Keyboard Shortcut: Ctrl+w
'
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
End Sub
