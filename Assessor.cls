VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public AssessorRecipient As Recipient
Public Qualified As Boolean
Public OT As Boolean
Public Location As String
Public EmailName As String
Public FilterFlag As Boolean 'True means it is a suitable assessor
Private Calendar As MAPIFolder
Public CalItems As New Collection
Private Utilisation As Single
Private CalStoreID As Variant


Public Function SetAssessor(ByVal Name As String, ByVal Qual As Boolean, ByVal OccTherapist As Boolean, ByVal Loc As String, ByRef oNS As NameSpace)
    Set AssessorRecipient = oNS.CreateRecipient(Name)
    AssessorRecipient.Resolve

    Qualified = Qual
    OT = OccTherapist
    Location = Loc
    FilterFlag = True

    If AssessorRecipient.Resolved Then
        EmailName = Name
        SetAssessor = True
    Else
        MsgBox "The user " & Name & " cannot be found in the global address book - please check and confirm the name in Outlook."
        SetAssessor = False
    End If
End Function

Public Function RefreshCalItems(LowDate As Date, HighDate As Date, ByRef oNS As NameSpace) As Boolean
    Dim strQry As String
    Dim oSubItems As Items
    Dim oItem As AppointmentItem
    Dim oNew As AppointmentItem
    Dim sCats As String
    Dim n As Long

    strQry = "[AllDayEvent] = False " & _
                "AND [Start] > '" & Format(LowDate, "ddddd h:nn AMPM") & "' " & _
                "AND [End] <= '" & Format(HighDate, "ddddd h:nn AMPM") & "' " & _
                "AND [Label] <> 3"

    On Error GoTo OpFailed

    Set Calendar = oNS.GetSharedDefaultFolder(AssessorRecipient, olFolderCalendar)

    DoEvents
    Sleep 500

    Set oSubItems = Calendar.Items.Restrict(strQry)

    'Empty the calitems collection
    If Not CalItems Is Nothing Then
        For n = 1 To CalItems.Count
            CalItems.Remove 1
        Next n
    End If

    '
    'Repopulate with the right categories
    '
    For Each oItem In oSubItems
        sCats = oItem.Categories
        If sCats <> "" Then
            If InStr(",", sCats) > 0 Then
                sCats = Left(sCats, InStr(",", sCats) - 1)
            End If
            'See if we have the right category here
            If InStr("Empty", sCats) > 0 Then
                'Add the item to the CalItems collection
                CalItems.Add oItem
            End If
        End If
    Next oItem

    '
    'Add in the green labels
    '
    strQry = "[AllDayEvent] = False " & _
                "AND [Start] > '" & Format(LowDate, "ddddd h:nn AMPM") & "' " & _
                "AND [End] <= '" & Format(HighDate, "ddddd h:nn AMPM") & "' " & _
                "AND [Label] = 3"

    Set oSubItems = Calendar.Items.Restrict(strQry)

    For Each oItem In oSubItems
        CalItems.Add oItem
    Next oItem

    RefreshCalItems = True
Exit Function

OpFailed:
    MsgBox "The user " & AssessorRecipient.Name & " does not have correct permissions on their calendar - please inform them. You may be able to check their slots manually."
    RefreshCalItems = False
End Function

Public Sub CalcUtilisation()
    'Get a count of all their slots
    Dim CalSlots As Items
    Dim oItem As AppointmentItem
    Dim strQry As String
    Dim SlotCount As Long
    Dim sCats As String

    strQry = "[AllDayEvent] = False " & _
                "AND [Start] > '" & Format(Now + 1, "ddddd h:nn AMPM") & "' " & _
                "AND [End] <= '" & Format(Now + 29, "ddddd h:nn AMPM") & "' " & _
                "AND ([Label] = 3 or [Label] = 1 or [Label] = 5)"
    Set CalSlots = Calendar.Items.Restrict(strQry)

    SlotCount = 0
    For Each oItem In CalSlots
        SlotCount = SlotCount + 1
    Next oItem

    'Now do all the categories
    strQry = "[AllDayEvent] = False " & _
                "AND [Start] > '" & Format(Now + 1, "ddddd h:nn AMPM") & "' " & _
                "AND [End] <= '" & Format(Now + 29, "ddddd h:nn AMPM") & "' " & _
                "AND NOT ([Label] = 3 OR [Label] = 1 OR [Label] = 5)"
    Set CalSlots = Calendar.Items.Restrict(strQry)

    For Each oItem In CalSlots
        sCats = oItem.Categories
        If sCats <> "" Then
            If InStr(",", sCats) > 0 Then
                sCats = Left(sCats, InStr(",", sCats) - 1)
            End If
            'Add the item to the CalItems collection
            If InStr("Empty", sCats) > 0 Then
                SlotCount = SlotCount + 1
            ElseIf InStr("Assessment", sCats) > 0 Then
                SlotCount = SlotCount + 1
            ElseIf InStr("Review", sCats) > 0 Then
                SlotCount = SlotCount + 1
            End If
        End If
    Next oItem

    'Set utilisation to be one minus green over all
    If SlotCount = 0 Then
        Utilisation = 0
    Else
        Utilisation = Abs(1 - (CSng(CalItems.Count) / CSng(SlotCount)))
    End If
End Sub

Public Sub CloseAll()
    'Close all folders etc...
    Dim oLoopItem As AppointmentItem

    Set AssessorRecipient = Nothing
    Set Calendar = Nothing

    For Each oLoopItem In CalItems
        Set oLoopItem = Nothing
    Next oLoopItem
    Set CalItems = Nothing
End Sub

Public Function GetUtilisation() As Single
    GetUtilisation = Utilisation
End Function

Public Property Get StoreID() As Variant
    StoreID = CalStoreID
End Property
