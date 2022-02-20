VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   " Progress Indicator"
   ClientHeight    =   468
   ClientLeft      =   75
   ClientTop       =   270
   ClientWidth     =   2295
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub code()

Dim i As Integer, j As Integer, pctCompl As Single

Sheet1.Cells.Clear

For i = 1 To 100
    For j = 1 To 1000
        Cells(i, 1).Value = j
    Next j
    pctCompl = i
    progress pctCompl
Next i

End Sub
Public Sub progress(pctCompl As Single)

UserForm1.Text.Caption = pctCompl & "% Completed"
UserForm1.Bar.Width = pctCompl * 2

DoEvents

End Sub
Private Sub UserForm_Activate()

'code

End Sub

Private Sub UserForm_Click()

End Sub
