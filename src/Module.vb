Option Explicit

Dim oRibbon As IRibbonUI

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set oRibbon = ribbon
End Sub

Public Sub ButtonInvalidate_OnAction(button As IRibbonControl)
    oRibbon.InvalidateControl("btnToggle")
End Sub



Public Sub Button_OnAction(button As IRibbonControl, pressed As Boolean)
    ' IRibbonControl.Context: Represents the active window containing the Ribbon user interface that triggers a callback procedure. Read-only.
    MsgBox "OnAction button " & button.Id &" in window " & button.Context & "(" & TypeName(button.Context) &")"
End Sub

Public Sub Button_GetPressed(button As IRibbonControl, ByRef returnIsPressed)
    ' GetPressed event will be called for each presentation opened in PowerPoint
    ' the button.Context will always reference the active presentation window
    ' according to the documentation, the button.Context is a window with the instance of the button
    MsgBox "GetPressed button " & button.Id &" in window " & button.Context & "(" & TypeName(button.Context) &")"
    returnIsPressed = Not returnIsPressed
End Sub
