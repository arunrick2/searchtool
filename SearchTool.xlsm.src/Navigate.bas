Attribute VB_Name = "Navigate"
Option Explicit

Dim Rib As IRibbonUI
Public MyTag As String
Dim PressedState As Boolean

'Callback for customUI.onLoad
Sub TSSRibbonOnLoad2(ribbon As IRibbonUI)
    Set Rib = ribbon
    PressedState = False
    'If you want to run a macro below when you open the workbook
    'you can call the macro like this :
    'Call EnableControlsWithCertainTag3
End Sub

Sub SendCToTSSRib(control As IRibbonControl)
    SendCToTSS
End Sub

Sub SendAToTSSRib(control As IRibbonControl)
    SendAToTSS
End Sub
