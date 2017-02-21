VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Schedule1 
   Caption         =   "Schedule Appointment v1001"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13515
   OleObjectBlob   =   "Schedule1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Schedule1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnClose_Click()
    Schedule1.hide
End Sub

Private Sub btnNext_Click()
    If LocationFilter() Then
        
        Schedule2.Show
        Schedule1.hide
    Else
        MsgBox "There are no people that cover this area.", vbCritical
    End If
End Sub

Public Function LocationFilter() As Boolean
    'Set the flags to only pass assessors in the right area
    On Error GoTo CannotFilter
    Dim PCLeft As String
    
    'Check to see if there is a GP entered
    If cbGP = "" Then
        'Go on the postcode
        If Len(tbPostcode) > 4 Then
            If Len(Replace(tbPostcode, " ", "")) = 7 Then
                PCLeft = Left(tbPostcode, 4)
            Else
                PCLeft = Left(tbPostcode, 3)
            End If
        Else
            PCLeft = tbPostcode
        End If
        If Not SetFlags("Location", PCLeft) Then GoTo CannotFilter
    Else
        If Not SetFlags("Location", cbGP) Then GoTo CannotFilter
    End If
    
    LocationFilter = True
Exit Function

CannotFilter:
    LocationFilter = False
End Function

Private Sub UserForm_Initialize()
    'Set the combobox with GPs
    SetSchedule1GPs
End Sub
