' listEvents.vbs - Lists calendar events within a specified date range
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    ' Get command line arguments
    Dim startDateStr, endDateStr, calendarName
    Dim startDate, endDate
    
    ' Get and validate arguments
    startDateStr = GetArgument("startDate")
    endDateStr = GetArgument("endDate")
    calendarName = GetArgument("calendar")
    
    ' Require start date
    RequireArgument "startDate"
    
    ' Parse dates
    startDate = ParseDate(startDateStr)
    
    ' If end date is not provided, use start date (single day)
    If endDateStr = "" Then
        endDate = startDate
    Else
        endDate = ParseDate(endDateStr)
    End If
    
    ' Ensure end date is not before start date
    If endDate < startDate Then
        OutputError "End date cannot be before start date"
        WScript.Quit 1
    End If
    
    ' Get calendar events
    Dim events
    Set events = GetCalendarEvents(startDate, endDate, calendarName)
    
    ' Output events as JSON
    OutputSuccess AppointmentsToJSON(events)
End Sub

' Gets calendar events within the specified date range
Function GetCalendarEvents(startDate, endDate, calendarName)
    On Error Resume Next
    
    ' Create Outlook objects
    Dim outlookApp, calendar, filter, events
    
    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    
    ' Get calendar folder
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If
    
    ' Instead of using Outlook's Restrict method which is unreliable across locales,
    ' we'll get all events and filter them manually

    ' Get all calendar items
    Set events = calendar.Items

    ' Sort by start date for consistent order
    events.Sort "[Start]"

    ' Create a new collection for filtered events
    Dim filteredEvents, item, itemStart, itemEnd
    Set filteredEvents = CreateObject("Scripting.Dictionary")

    ' The start of the day we're filtering from (at 00:00:00)
    Dim dayStart, dayEnd
    dayStart = DateSerial(Year(startDate), Month(startDate), Day(startDate))

    ' End of the last day in range (at 23:59:59)
    dayEnd = DateSerial(Year(endDate), Month(endDate), Day(endDate)) + 1 - (1/86400)

    ' Loop through all events and keep those in our date range
    For Each item In events
        ' Cast dates to ensure proper comparison
        itemStart = CDate(item.Start)
        itemEnd = CDate(item.End)

        ' Check if event overlaps with our date range
        ' An event overlaps if:
        ' 1. It starts within our range, OR
        ' 2. It ends within our range, OR
        ' 3. It starts before our range AND ends after our range (spans the range)
        If (itemStart >= dayStart And itemStart <= dayEnd) Or _
           (itemEnd >= dayStart And itemEnd <= dayEnd) Or _
           (itemStart < dayStart And itemEnd > dayEnd) Then

            ' Use EntryID as key to avoid duplicates
            If Not filteredEvents.Exists(item.EntryID) Then
                filteredEvents.Add item.EntryID, item
            End If
        End If
    Next

    ' Create a collection for the filtered events
    Dim resultCollection, counter
    Set resultCollection = CreateObject("System.Collections.ArrayList")

    ' Add all filtered events to the collection
    For Each item In filteredEvents.Items
        resultCollection.Add(item)
    Next

    If Err.Number <> 0 Then
        OutputError "Failed to get calendar events: " & Err.Description
        WScript.Quit 1
    End If

    ' Return filtered events
    Set GetCalendarEvents = resultCollection

    ' Clean up
    Set calendar = Nothing
    Set outlookApp = Nothing
    Set filteredEvents = Nothing
End Function

' Run the main function
Main
