Option Explicit

' Zero-minute meeting reminders for the primary Outlook calendar.
' Paste this entire file into the built-in ThisOutlookSession module.

Private WithEvents ZeroReminderCalendarItems As Outlook.Items
Private ZeroReminderIsApplying As Boolean

Private ZeroReminderChangedCount As Long
Private ZeroReminderSkippedCount As Long
Private ZeroReminderErrorCount As Long

Private Enum ZeroReminderEligibilityResult
    ZeroReminderEligibilityError = -1
    ZeroReminderEligibilitySkip = 0
    ZeroReminderEligibilityApply = 1
End Enum

Private Sub Application_Startup()
    ZeroReminderInitialize True
End Sub

Private Sub Application_MAPILogonComplete()
    ' If Startup ran before the default MAPI store was ready, retry once the
    ' profile has finished logging on. Avoid a second repair when already bound.
    If ZeroReminderCalendarItems Is Nothing Then
        ZeroReminderInitialize True
    End If
End Sub

Private Sub Application_Quit()
    Set ZeroReminderCalendarItems = Nothing
End Sub

Private Sub ZeroReminderCalendarItems_ItemAdd(ByVal Item As Object)
    ZeroReminderProcessObject Item, True
End Sub

Private Sub ZeroReminderCalendarItems_ItemChange(ByVal Item As Object)
    ZeroReminderProcessObject Item, True
End Sub

' Optional diagnostic entry point. Normal operation does not require this macro.
' Run it from Alt+F8 after installation if an immediate repair/check is wanted.
Public Sub ZeroReminder_RepairNow()
    ZeroReminderChangedCount = 0
    ZeroReminderSkippedCount = 0
    ZeroReminderErrorCount = 0

    ZeroReminderInitialize False

    Debug.Print "ZeroReminder repair finished at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "  Changed: " & CStr(ZeroReminderChangedCount)
    Debug.Print "  Skipped: " & CStr(ZeroReminderSkippedCount)
    Debug.Print "  Errors:  " & CStr(ZeroReminderErrorCount)
End Sub

Private Sub ZeroReminderInitialize(ByVal IsAutomaticStartup As Boolean)
    On Error GoTo InitializeFailed

    Set ZeroReminderCalendarItems = _
        Application.Session.GetDefaultFolder(olFolderCalendar).Items

    ZeroReminderRepairCalendar
    Exit Sub

InitializeFailed:
    ZeroReminderErrorCount = ZeroReminderErrorCount + 1
    ZeroReminderLogError "Initialize", Err.Number, Err.Description

    ' Keep Outlook startup quiet. A manually invoked repair leaves details
    ' in the Immediate window (Alt+F11, then Ctrl+G).
    If Not IsAutomaticStartup Then
        Debug.Print "ZeroReminder could not attach to the primary calendar."
    End If
    Err.Clear
End Sub

Private Sub ZeroReminderRepairCalendar()
    Dim FutureItems As Outlook.Items
    Dim RecurringItems As Outlook.Items
    Dim FutureFilter As String

    On Error GoTo FilterFailed

    If ZeroReminderCalendarItems Is Nothing Then Exit Sub

    ' Restrict the large non-recurring population to future items. Recurring
    ' masters need a separate pass because their master Start can be years in
    ' the past while the pattern still has future occurrences.
    FutureFilter = "[IsRecurring] = False AND [Start] > '" & _
        Format$(Now, "ddddd h:nn AMPM") & "'"

    Set FutureItems = ZeroReminderCalendarItems.Restrict(FutureFilter)
    Set RecurringItems = _
        ZeroReminderCalendarItems.Restrict("[IsRecurring] = True")

    ZeroReminderProcessCollection FutureItems, False
    ZeroReminderProcessCollection RecurringItems, True
    GoTo RepairFinished

FullScan:
    On Error GoTo RepairFailed
    ZeroReminderProcessCollection ZeroReminderCalendarItems, True
    GoTo RepairFinished

RepairFinished:
    Set RecurringItems = Nothing
    Set FutureItems = Nothing
    Exit Sub

FilterFailed:
    ' Regional settings and store providers can differ. If Restrict is not
    ' available, retain correctness with the original full-folder repair.
    ZeroReminderErrorCount = ZeroReminderErrorCount + 1
    ZeroReminderLogError "RepairCalendar filter; using full scan", _
        Err.Number, Err.Description
    Err.Clear
    Resume FullScan

RepairFailed:
    ZeroReminderErrorCount = ZeroReminderErrorCount + 1
    ZeroReminderLogError "RepairCalendar fallback", Err.Number, Err.Description
    Err.Clear
    Resume RepairFinished
End Sub

Private Sub ZeroReminderProcessCollection(ByVal CalendarItems As Outlook.Items, _
                                          ByVal IncludeSeriesExceptions As Boolean)
    Dim CalendarItem As Object

    On Error GoTo CollectionFailed

    If CalendarItems Is Nothing Then Exit Sub

    For Each CalendarItem In CalendarItems
        ZeroReminderProcessObject CalendarItem, IncludeSeriesExceptions
    Next CalendarItem

CollectionFinished:
    Set CalendarItem = Nothing
    Exit Sub

CollectionFailed:
    ZeroReminderErrorCount = ZeroReminderErrorCount + 1
    ZeroReminderLogError "ProcessCollection", Err.Number, Err.Description
    Err.Clear
    Resume CollectionFinished
End Sub

Private Sub ZeroReminderProcessObject(ByVal Item As Object, _
                                      ByVal IncludeSeriesExceptions As Boolean)
    Dim Appointment As Outlook.AppointmentItem

    On Error GoTo ProcessFailed

    If ZeroReminderIsApplying Then Exit Sub
    If Item Is Nothing Then Exit Sub
    If Not TypeOf Item Is Outlook.AppointmentItem Then Exit Sub

    Set Appointment = Item
    ZeroReminderApplyToAppointment Appointment, IncludeSeriesExceptions

ProcessFinished:
    Set Appointment = Nothing
    Exit Sub

ProcessFailed:
    ZeroReminderErrorCount = ZeroReminderErrorCount + 1
    ZeroReminderLogError "ProcessObject", Err.Number, Err.Description
    Err.Clear
    Resume ProcessFinished
End Sub

Private Sub ZeroReminderApplyToAppointment( _
        ByVal Appointment As Outlook.AppointmentItem, _
        ByVal IncludeSeriesExceptions As Boolean)

    Dim NeedsSave As Boolean
    Dim Eligibility As ZeroReminderEligibilityResult

    On Error GoTo ApplyFailed

    Eligibility = ZeroReminderGetEligibility(Appointment)
    Select Case Eligibility
        Case ZeroReminderEligibilitySkip
            ZeroReminderSkippedCount = ZeroReminderSkippedCount + 1
            Exit Sub
        Case ZeroReminderEligibilityError
            ' The eligibility function already recorded the error.
            Exit Sub
        Case ZeroReminderEligibilityApply
            ' Continue with reminder normalization.
    End Select

    NeedsSave = Not Appointment.ReminderSet

    If Appointment.ReminderSet Then
        If Appointment.ReminderMinutesBeforeStart <> 0 Then NeedsSave = True
    End If

    If Not Appointment.ReminderOverrideDefault Then NeedsSave = True
    If Not Appointment.ReminderPlaySound Then NeedsSave = True

    If NeedsSave Then
        ZeroReminderIsApplying = True

        Appointment.ReminderSet = True
        Appointment.ReminderOverrideDefault = True
        Appointment.ReminderPlaySound = True
        Appointment.ReminderMinutesBeforeStart = 0
        Appointment.Save

        ZeroReminderIsApplying = False
        ZeroReminderChangedCount = ZeroReminderChangedCount + 1
    Else
        ZeroReminderSkippedCount = ZeroReminderSkippedCount + 1
    End If

    ' A recurring master controls normal occurrences. Existing exceptions can
    ' have their own reminder values, so repair them separately without
    ' recursively reopening the series.
    If IncludeSeriesExceptions Then
        If Appointment.IsRecurring Then
            If Appointment.RecurrenceState = olApptMaster Then
                ZeroReminderApplyToSeriesExceptions Appointment
            End If
        End If
    End If

ApplyFinished:
    ZeroReminderIsApplying = False
    Exit Sub

ApplyFailed:
    ZeroReminderIsApplying = False
    ZeroReminderErrorCount = ZeroReminderErrorCount + 1
    ZeroReminderLogError "ApplyToAppointment", Err.Number, Err.Description
    Err.Clear
    Resume ApplyFinished
End Sub

Private Function ZeroReminderGetEligibility( _
        ByVal Appointment As Outlook.AppointmentItem) _
        As ZeroReminderEligibilityResult

    Dim MeetingStatus As Outlook.OlMeetingStatus

    On Error GoTo EligibilityFailed

    ZeroReminderGetEligibility = ZeroReminderEligibilitySkip

    If Appointment.AllDayEvent Then Exit Function

    MeetingStatus = Appointment.MeetingStatus
    Select Case MeetingStatus
        Case olMeeting, olMeetingReceived
            ' Include both organizer-owned and received meeting items.
        Case olMeetingCanceled, olMeetingReceivedAndCanceled, olNonMeeting
            Exit Function
        Case Else
            Exit Function
    End Select

    If Appointment.ResponseStatus = olResponseDeclined Then Exit Function

    If Appointment.IsRecurring And _
       Appointment.RecurrenceState = olApptMaster Then
        ZeroReminderGetEligibility = _
            ZeroReminderGetSeriesEligibility(Appointment)
    Else
        ' Ongoing and completed meetings are deliberately treated as past.
        If Appointment.Start > Now Then
            ZeroReminderGetEligibility = ZeroReminderEligibilityApply
        End If
    End If

    Exit Function

EligibilityFailed:
    ZeroReminderErrorCount = ZeroReminderErrorCount + 1
    ZeroReminderLogError "GetEligibility", Err.Number, Err.Description
    Err.Clear
    ZeroReminderGetEligibility = ZeroReminderEligibilityError
End Function

Private Function ZeroReminderGetSeriesEligibility( _
        ByVal SeriesMaster As Outlook.AppointmentItem) _
        As ZeroReminderEligibilityResult

    Dim Pattern As Outlook.RecurrencePattern

    On Error GoTo SeriesCheckFailed

    ZeroReminderGetSeriesEligibility = ZeroReminderEligibilitySkip

    Set Pattern = SeriesMaster.GetRecurrencePattern

    If Pattern.NoEndDate Then
        ZeroReminderGetSeriesEligibility = ZeroReminderEligibilityApply
    Else
        ' PatternEndDate is date-only. A series ending today is retained because
        ' it may still contain an occurrence or a moved exception later today.
        If Pattern.PatternEndDate >= Date Then
            ZeroReminderGetSeriesEligibility = ZeroReminderEligibilityApply
        End If
    End If

SeriesCheckFinished:
    Set Pattern = Nothing
    Exit Function

SeriesCheckFailed:
    ZeroReminderErrorCount = ZeroReminderErrorCount + 1
    ZeroReminderLogError "GetSeriesEligibility", Err.Number, Err.Description
    Err.Clear
    ZeroReminderGetSeriesEligibility = ZeroReminderEligibilityError
    Resume SeriesCheckFinished
End Function

Private Sub ZeroReminderApplyToSeriesExceptions( _
        ByVal SeriesMaster As Outlook.AppointmentItem)

    Dim Pattern As Outlook.RecurrencePattern
    Dim Exceptions As Outlook.Exceptions
    Dim SeriesException As Outlook.Exception
    Dim ExceptionAppointment As Outlook.AppointmentItem
    Dim ExceptionIndex As Long

    On Error GoTo ExceptionsFailed

    Set Pattern = SeriesMaster.GetRecurrencePattern
    Set Exceptions = Pattern.Exceptions

    For ExceptionIndex = 1 To Exceptions.Count
        Set SeriesException = Nothing
        Set ExceptionAppointment = Nothing

        On Error Resume Next
        Set SeriesException = Exceptions.Item(ExceptionIndex)

        If Err.Number <> 0 Then
            ZeroReminderErrorCount = ZeroReminderErrorCount + 1
            ZeroReminderLogError "Read series exception " & CStr(ExceptionIndex), _
                Err.Number, Err.Description
            Err.Clear
        ElseIf Not SeriesException Is Nothing Then
            If Not SeriesException.Deleted Then
                Set ExceptionAppointment = SeriesException.AppointmentItem
                If Err.Number = 0 And Not ExceptionAppointment Is Nothing Then
                    ZeroReminderApplyToAppointment ExceptionAppointment, False
                ElseIf Err.Number <> 0 Then
                    ZeroReminderErrorCount = ZeroReminderErrorCount + 1
                    ZeroReminderLogError _
                        "Open series exception " & CStr(ExceptionIndex), _
                        Err.Number, Err.Description
                    Err.Clear
                End If
            End If
        End If
        On Error GoTo ExceptionsFailed
    Next ExceptionIndex

ExceptionsFinished:
    Set ExceptionAppointment = Nothing
    Set SeriesException = Nothing
    Set Exceptions = Nothing
    Set Pattern = Nothing
    Exit Sub

ExceptionsFailed:
    ZeroReminderErrorCount = ZeroReminderErrorCount + 1
    ZeroReminderLogError "ApplyToSeriesExceptions", Err.Number, Err.Description
    Err.Clear
    Resume ExceptionsFinished
End Sub

Private Sub ZeroReminderLogError(ByVal Operation As String, _
                                 ByVal ErrorNumber As Long, _
                                 ByVal ErrorDescription As String)
    Debug.Print "ZeroReminder [" & Operation & "] error " & _
        CStr(ErrorNumber) & ": " & ErrorDescription
End Sub
