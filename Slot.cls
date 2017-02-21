VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Slot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Rejected As Boolean 'Whether it has been rejected by the scheduler
Public oSlot As AppointmentItem 'The appointmentitem which represents the slot
Public AssessorName As String
Public AssessorType As String
Public AssessorLocation As String
Public AssessorUtilisation As Single
Public SlotDate As Date 'Holds both date and time together - the start time of the slot
Public Score As Single 'The score of the slot (based on utilisation, and time to slot)
Public ParentID As Variant 'StoreID of the parent calendar

Public Sub Reject()
    Rejected = True
    Score = 999999
End Sub

Public Sub Initialise(ByRef oAptmt As AppointmentItem, ByRef oAssessor As Assessor, StoreID As Variant)
    
    'Set all the internal variables based on the assessor and the appointmentitem

    AssessorName = oAssessor.EmailName
    If oAssessor.Qualified Then
        If oAssessor.OT Then
            AssessorType = "OT"
        Else
            AssessorType = "CM"
        End If
    Else
        If oAssessor.OT Then
            AssessorType = "AO OT"
        Else
            AssessorType = "AO"
        End If
    End If
    AssessorLocation = oAssessor.Location
    AssessorUtilisation = oAssessor.GetUtilisation
    ParentID = StoreID
    SlotDate = oAptmt.Start
    Rejected = False
    Set oSlot = oAptmt

    CalcScore
    
End Sub

Public Sub CalcScore()
    'Give the slot a score based on it's information
    'Make it a bit random as well to spread between staff
    
    If SlotDate > Now + 1 Then
        Score = (0.01 + AssessorUtilisation) * CSng(SlotDate - Now) + (Rnd() / 1000)
        Rejected = False
    Else
        Reject
    End If
End Sub
