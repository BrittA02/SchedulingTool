VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Schedule2 
   Caption         =   "Schedule Appointment v1001"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13515
   OleObjectBlob   =   "Schedule2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Schedule2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnAO_Click()
    If SetFlags("Skill", "AO") Then
        Schedule3.Show
        Schedule2.hide
    End If
End Sub

Private Sub btnAOOT_Click()
    If SetFlags("Skill", "AO OT") Then
        Schedule3.Show
        Schedule2.hide
    End If
End Sub

Private Sub btnCancel_Click()
    Schedule2.hide
End Sub

Private Sub btnCM_Click()
    If SetFlags("Skill", "CM") Then
        Schedule3.Show
        Schedule2.hide
    End If
End Sub

Private Sub btnOT_Click()
    If SetFlags("Skill", "OT") Then
        Schedule3.Show
        Schedule2.hide
    End If
End Sub
