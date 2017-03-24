VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Schedule3
   Caption         =   "Schedule Appointment v1001"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13515
   OleObjectBlob   =   "Schedule3.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Schedule3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Slot1 As Slot
Dim Slot2 As Slot

Private Sub btnReplace1_Click()
    If Slot1 Is Nothing Then Exit Sub
    Slot1.Reject
    SortSlots
    ShowSlots
End Sub

Private Sub btnReplace2_Click()
    If Slot2 Is Nothing Then Exit Sub
    Slot2.Reject
    SortSlots
    ShowSlots
End Sub

Private Sub btnResetScores_Click()
    RecalcSlotScores
    SortSlots
    ShowSlots
End Sub

Private Sub btnUse1_Click()
    If Slot1 Is Nothing Then Exit Sub
    If Slot1.Score > 1000 Then Exit Sub

    If MsgBox("Are you sure you wish to use this slot?", vbYesNo) = vbNo Then Exit Sub
    UseSlot Slot1
End Sub

Private Sub btnUse2_Click()
    If Slot2 Is Nothing Then Exit Sub
    If Slot2.Score > 1000 Then Exit Sub

    If MsgBox("Are you sure you wish to use this slot?", vbYesNo) = vbNo Then Exit Sub
    UseSlot Slot2
End Sub

Private Sub UserForm_Activate()
    ResetForm
End Sub

Public Sub ResetForm()
    'Fetch and calculate all the data we need
    RetrieveCalItems
    ShowSlots
End Sub

Private Sub ShowSlots()
    RetrieveBestSlots

    'Set the labels on the form
    If Slot1 Is Nothing Then
        lAssessor1.Caption = "Not available"
        lType1.Caption = ""
        lLocation1.Caption = ""
        lDate1.Caption = "Add referral to waiting list"
        lTime1.Caption = ""
    Else
        If Slot1.Rejected Then
            lAssessor1.Caption = "Not available"
            lType1.Caption = ""
            lLocation1.Caption = ""
            lDate1.Caption = "Add referral to waiting list"
            lTime1.Caption = ""
        Else
            lAssessor1.Caption = Slot1.AssessorName
            lType1.Caption = Slot1.AssessorType
            lLocation1.Caption = Slot1.AssessorLocation
            lDate1.Caption = Format(Slot1.SlotDate, "ddd dd/mm/yyyy")
            lTime1.Caption = Format(Slot1.SlotDate, "hh:mm AMPM")
        End If
    End If

    If Slot2 Is Nothing Then
        lAssessor2.Caption = "Not available"
        lType2.Caption = ""
        lLocation2.Caption = ""
        lDate2.Caption = "Add referral to waiting list"
        lTime2.Caption = ""
    Else
        If Slot2.Rejected Then
            lAssessor2.Caption = "Not available"
            lType2.Caption = ""
            lLocation2.Caption = ""
            lDate2.Caption = "Add referral to waiting list"
            lTime2.Caption = ""
        Else
            lAssessor2.Caption = Slot2.AssessorName
            lType2.Caption = Slot2.AssessorType
            lLocation2.Caption = Slot2.AssessorLocation
            lDate2.Caption = Format(Slot2.SlotDate, "ddd dd/mm/yyyy")
            lTime2.Caption = Format(Slot2.SlotDate, "hh:mm AMPM")
        End If
    End If
End Sub

Private Sub RetrieveBestSlots()
    Set Slot1 = GetSlot(1)
    Set Slot2 = GetSlot(2)
End Sub
