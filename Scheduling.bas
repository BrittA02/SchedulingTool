Attribute VB_Name = "Scheduling"
Dim oAs As New Collection 'For holding all the assessors - see Assessor Class Module
Dim Locations As New Collection
Dim Slots As New Collection
Dim PostcodeCount As Integer
Dim GPCount As Integer

Public Sub OpenScheduler()
    'Make sure we have a decent list of assessors, then show the first form
    LoadParameters
    Schedule1.Show
End Sub

Private Sub LoadParameters()
    Dim oNS As NameSpace
    Dim oXL As Object 'Excel.Application
    Dim oWB As Object 'Excel.Workbook
    Dim oWSAssess As Object 'Excel.Worksheet
    Dim oWSLookup As Object 'Excel.Worksheet

    'Start from scratch
    ClearCollections

    'Set up namespace for contacts
    Set oNS = Application.GetNamespace("MAPI")

    'Open up the assessor list file
    Set oXL = CreateObject("Excel.Application")
    Set oWB = oXL.Workbooks.Open("PATH_TO\Assessor List.xls")
    Set oWSAssess = oWB.Worksheets("Assessors")
    Set oWSLookup = oWB.Worksheets("Lookup")

    PopulateAssessors oWSAssess, oNS
    PopulateLocations oWSLookup

    'Shut everything down
    oWB.Close savechanges:=False
    oXL.Quit
    Set oWS = Nothing
    Set oWB = Nothing
    Set oXL = Nothing
End Sub

Private Sub PopulateAssessors(oWSAssess As Object, oNS As NameSpace)
    Dim AssessName As String
    Dim AssessOT As Boolean
    Dim AssessQual As Boolean
    Dim AssessLoc As String
    Dim oNewAs As Assessor
    Dim n As Integer

    'On Error Resume Next

    'Loop through the excel file listing the staff and their roles
    For n = 1 To oWSAssess.Range("AssessorCount")
        'Pull parameters
        AssessName = oWSAssess.Range("FirstEntry").offset(n - 1, 0)
        AssessQual = oWSAssess.Range("FirstEntry").offset(n - 1, 1)
        AssessOT = oWSAssess.Range("FirstEntry").offset(n - 1, 2)
        AssessLoc = oWSAssess.Range("FirstEntry").offset(n - 1, 3)

        'Create assessor
        Set oNewAs = New Assessor
        If oNewAs.SetAssessor(AssessName, AssessQual, AssessOT, AssessLoc, oNS) Then
            'Add assessor to collection
            oAs.Add oNewAs
        End If
    Next n
End Sub

Private Sub PopulateLocations(oWSLookup As Object)
    Dim oLocation As Location
    Dim n As Integer

    'Populate the locations collection
    PostcodeCount = oWSLookup.Range("PostcodeCount")
    For n = 1 To PostcodeCount
        'Create location
        Set oLocation = New Location
        oLocation.Name = oWSLookup.Range("Postcodes")(n)
        oLocation.Office = oWSLookup.Range("Postcodes")(n).offset(0, 1)
        Locations.Add oLocation, oLocation.Name
    Next n
    GPCount = oWSLookup.Range("SurgeryCount")
    For n = 1 To GPCount
        'Create location
        Set oLocation = New Location
        oLocation.Name = oWSLookup.Range("GPs")(n)
        oLocation.Office = oWSLookup.Range("GPs")(n).offset(0, 1)
        Locations.Add oLocation, oLocation.Name
    Next n
End Sub

Private Sub ClearCollections()
    Dim oLoopAss As Assessor

    'Bin all open folders and other stuff
    For Each oLoopAss In oAs
        oLoopAss.CloseAll
    Next oLoopAss

    'Clear the oAs collection
    If oAs.Count > 0 Then
        For n = 1 To oAs.Count
            oAs.Remove 1
        Next n
    End If

    'Clear the Locations collection
    If Locations.Count > 0 Then
        For n = 1 To Locations.Count
            Locations.Remove 1
        Next n
    End If
End Sub

Public Sub SetSchedule1GPs()
    Dim n As Integer
    Dim oLoc As Location

    For n = 1 To GPCount
        Set oLoc = Locations.Item(n + PostcodeCount)
        Schedule1.cbGP.AddItem oLoc.Name
    Next n
End Sub

Public Function SetFlags(FilterType As String, FilterName As String)
    On Error GoTo FlagFail

    Dim oLoopAss As Assessor
    Dim Office As String
    Dim Qual As Boolean
    Dim OT As Boolean
    Dim LocPass As Boolean

    Select Case FilterType
        Case "Location"
            'Get the office from the filtername
            Office = Locations.Item(FilterName).Office
            LocPass = False

            'Loop through the assessors and set flags. Location is done first so set both true and false
            For Each oLoopAss In oAs
                If oLoopAss.Location = Office Then
                    oLoopAss.FilterFlag = True
                    LocPass = True
                Else
                    oLoopAss.FilterFlag = False
                End If
            Next oLoopAss

            If LocPass = False Then GoTo FlagFail

        Case "Skill"
            'Get flag settings based on filtername
            Select Case FilterName
            Case "CM"
                Qual = True
                OT = False
            Case "AO"
                Qual = False
                OT = False
            Case "OT"
                Qual = True
                OT = True
            Case "AO OT"
                Qual = False
                OT = True
            End Select

            'Loop through assessors, trim off those who have the wrong specialty
            For Each oLoopAss In oAs
                If oLoopAss.FilterFlag Then
                    If Not (oLoopAss.OT = OT And oLoopAss.Qualified = Qual) Then
                        oLoopAss.FilterFlag = False
                    End If
                End If
            Next oLoopAss
    End Select

    SetFlags = True
Exit Function

FlagFail:
    SetFlags = False
End Function

Public Function GetSlot(Rank As Integer) As Slot
    If Rank > Slots.Count Then
        Set GetSlot = Nothing
    Else
        Set GetSlot = Slots(Rank)
    End If
End Function

Public Sub RetrieveCalItems()
    'We need to get the calendar items for all the assessors still left after filtering
    Dim oLoopAss As Assessor
    Dim oAptmt As AppointmentItem
    Dim tmpSlot As Slot
    Dim n As Long
    Dim oNS As NameSpace

    Set oNS = Application.GetNamespace("MAPI")

    'Empty the slots collection
    For n = 1 To Slots.Count
        Slots.Remove 1
    Next n

    For Each oLoopAss In oAs
        If oLoopAss.FilterFlag Then
            If oLoopAss.RefreshCalItems(Now + 1, Now + 29, oNS) Then
                oLoopAss.CalcUtilisation

                'Fill the slots from their calitems
                For Each oAptmt In oLoopAss.CalItems
                    Set tmpSlot = New Slot
                    tmpSlot.Initialise oAptmt, oLoopAss, oLoopAss.StoreID
                    Slots.Add tmpSlot
                Next oAptmt
            End If
        End If
    Next oLoopAss

    SortSlots
End Sub

Public Sub SortSlots()
    'Sort the slots collection based on score ascending
    Set Slots = SortCollection(Slots)
End Sub

Public Sub RecalcSlotScores()
    Dim oSlot As Slot

    For Each oSlot In Slots
        oSlot.CalcScore
    Next oSlot

    SortSlots
End Sub

Public Function UseSlot(ChosenSlot As Slot) As Boolean
    '
    ' Allocates the case to this slot
    '

    'The slot itself
    Dim SubjectLine As String

    SubjectLine = InputBox("Please paste the name and swift ID from AIS.", "Slot Subject Line Required")

    If SubjectLine = "" Then
        MsgBox "Action cancelled - slot not booked"
        UseSlot = False
        Exit Function
    End If

    ChosenSlot.oSlot.Subject = SubjectLine
    ChosenSlot.oSlot.Categories = "Slot: Initial Assessment"

    MsgBox "The selected slot will now open. Please make sure you set the label to red."
    ChosenSlot.oSlot.Display

    Dim oMail As MailItem
    'An autoemail for me
    Set oMail = Application.CreateItem(olMailItem)
    oMail.To = "ACTScheduling@kent.gov.uk"
    oMail.Subject = "Slot allocated: " & Format(ChosenSlot.SlotDate, "dd-mmm hh:mm ampm") & ", to " & SubjectLine
    oMail.Body = "Slot allocated by user " & Environ("Username") & "." & vbCrLf & vbCrLf & "1. Please check your calendar to confirm the appointment is red."
    oMail.Body = oMail.Body & vbCrLf & "2. The BICA will follow shortly in a separate email from the CAO."
    oMail.Send

    'An autoemail for the assessor
    Set oMail = Application.CreateItem(olMailItem)
    oMail.To = ChosenSlot.AssessorName
    oMail.CC = "ACTScheduling@kent.gov.uk"
    oMail.Subject = "Slot allocated: " & Format(ChosenSlot.SlotDate, "dd-mmm hh:mm ampm") & ", to " & SubjectLine
    oMail.Body = "Slot allocated by user " & Environ("Username") & "." & vbCrLf & vbCrLf & "1. Please check your calendar to confirm the appointment is red."
    oMail.Body = oMail.Body & vbCrLf & "2. The BICA will follow shortly in a separate email from the CAO."
    oMail.Recipients.ResolveAll
    oMail.Display

    Schedule3.hide
End Function
