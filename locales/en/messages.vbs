' messages.vbs - English localization strings for Outlook Calendar MCP
Option Explicit

' Create Dictionary object to store localized strings
Dim g_messages
Set g_messages = CreateObject("Scripting.Dictionary")

' Initialize messages dictionary with all localized strings
Sub InitializeMessages()
    ' -------------------------------------------------------------------
    ' Error messages
    ' -------------------------------------------------------------------
    g_messages.Add "ERROR_PREFIX", "ERROR:"
    g_messages.Add "ERROR_OUTLOOK_APP_FAILED", "Failed to create Outlook Application: "
    g_messages.Add "ERROR_MAPI_NAMESPACE_FAILED", "Failed to get MAPI namespace: "
    g_messages.Add "ERROR_DEFAULT_CALENDAR_FAILED", "Failed to get default calendar: "
    g_messages.Add "ERROR_CALENDAR_NOT_FOUND", "Calendar not found: "
    g_messages.Add "ERROR_PARSE_DATE", "Failed to parse date: "
    g_messages.Add "ERROR_INVALID_DATE_FORMAT", "Invalid date format. Expected MM/DD/YYYY: "
    g_messages.Add "ERROR_INVALID_TIME_FORMAT", "Invalid time format: "
    g_messages.Add "ERROR_MISSING_REQUIRED_ARG", "Missing required argument: "
    g_messages.Add "ERROR_CREATE_EVENT", "Failed to create calendar event: "
    g_messages.Add "ERROR_UPDATE_EVENT", "Failed to update calendar event: "
    g_messages.Add "ERROR_DELETE_EVENT", "Failed to delete calendar event: "
    g_messages.Add "ERROR_EVENT_NOT_FOUND", "Event not found with ID: "
    g_messages.Add "ERROR_END_TIME_BEFORE_START", "End time cannot be before or equal to start time"
    g_messages.Add "ERROR_SCRIPT_EXECUTION", "Script execution failed: "
    g_messages.Add "ERROR_SCRIPT_OUTPUT_PARSE", "Failed to parse script output: "
    g_messages.Add "ERROR_UNEXPECTED_OUTPUT", "Unexpected script output: "
    
    ' -------------------------------------------------------------------
    ' Success messages
    ' -------------------------------------------------------------------
    g_messages.Add "SUCCESS_PREFIX", "SUCCESS:"
    g_messages.Add "SUCCESS_EVENT_CREATED", "Event created successfully"
    g_messages.Add "SUCCESS_EVENT_UPDATED", "Event updated successfully"
    g_messages.Add "SUCCESS_EVENT_DELETED", "Event deleted successfully"
    
    ' -------------------------------------------------------------------
    ' Date and time formats
    ' -------------------------------------------------------------------
    g_messages.Add "DATE_FORMAT_MMDDYYYY", "MM/DD/YYYY"
    g_messages.Add "DATE_FORMAT_DDMMYYYY", "DD/MM/YYYY"
    g_messages.Add "DATE_FORMAT_YYYYMMDD", "YYYY-MM-DD"
    g_messages.Add "TIME_FORMAT_HHMM", "HH:MM AM/PM"
    
    ' -------------------------------------------------------------------
    ' Calendar & Event Status Labels
    ' -------------------------------------------------------------------
    g_messages.Add "STATUS_BUSY", "Busy"
    g_messages.Add "STATUS_TENTATIVE", "Tentative"
    g_messages.Add "STATUS_FREE", "Free"
    g_messages.Add "STATUS_OUT_OF_OFFICE", "Out of Office"
    g_messages.Add "STATUS_UNKNOWN", "Unknown"
    
    ' -------------------------------------------------------------------
    ' Response Status Labels
    ' -------------------------------------------------------------------
    g_messages.Add "RESPONSE_ACCEPTED", "Accepted"
    g_messages.Add "RESPONSE_DECLINED", "Declined"
    g_messages.Add "RESPONSE_TENTATIVE", "Tentative"
    g_messages.Add "RESPONSE_NOT_RESPONDED", "Not Responded"
    g_messages.Add "RESPONSE_UNKNOWN", "Unknown"
    
    ' -------------------------------------------------------------------
    ' UI Labels
    ' -------------------------------------------------------------------
    g_messages.Add "LABEL_SUBJECT", "Subject"
    g_messages.Add "LABEL_START_DATE", "Start Date"
    g_messages.Add "LABEL_START_TIME", "Start Time"
    g_messages.Add "LABEL_END_DATE", "End Date"
    g_messages.Add "LABEL_END_TIME", "End Time"
    g_messages.Add "LABEL_LOCATION", "Location"
    g_messages.Add "LABEL_BODY", "Body"
    g_messages.Add "LABEL_ORGANIZER", "Organizer"
    g_messages.Add "LABEL_ATTENDEES", "Attendees"
    g_messages.Add "LABEL_CALENDAR", "Calendar"
    g_messages.Add "LABEL_STATUS", "Status"
    g_messages.Add "LABEL_ID", "ID"
    g_messages.Add "LABEL_IS_RECURRING", "Is Recurring"
    g_messages.Add "LABEL_IS_MEETING", "Is Meeting"
    
    ' -------------------------------------------------------------------
    ' Month names
    ' -------------------------------------------------------------------
    g_messages.Add "MONTH_1", "January"
    g_messages.Add "MONTH_2", "February"
    g_messages.Add "MONTH_3", "March"
    g_messages.Add "MONTH_4", "April"
    g_messages.Add "MONTH_5", "May"
    g_messages.Add "MONTH_6", "June"
    g_messages.Add "MONTH_7", "July"
    g_messages.Add "MONTH_8", "August"
    g_messages.Add "MONTH_9", "September"
    g_messages.Add "MONTH_10", "October"
    g_messages.Add "MONTH_11", "November"
    g_messages.Add "MONTH_12", "December"
    
    ' -------------------------------------------------------------------
    ' Day names
    ' -------------------------------------------------------------------
    g_messages.Add "DAY_1", "Sunday"
    g_messages.Add "DAY_2", "Monday"
    g_messages.Add "DAY_3", "Tuesday"
    g_messages.Add "DAY_4", "Wednesday"
    g_messages.Add "DAY_5", "Thursday"
    g_messages.Add "DAY_6", "Friday"
    g_messages.Add "DAY_7", "Saturday"
    
    ' -------------------------------------------------------------------
    ' Time parts
    ' -------------------------------------------------------------------
    g_messages.Add "TIME_AM", "AM"
    g_messages.Add "TIME_PM", "PM"
End Sub

' Function to retrieve localized string for a given key
' If key is not found, returns the key itself
Function L(key)
    ' Initialize messages on first call
    If g_messages.Count = 0 Then
        InitializeMessages
    End If
    
    ' Return localized string if exists, otherwise return the key itself
    If g_messages.Exists(key) Then
        L = g_messages(key)
    Else
        ' Log missing key for translator's reference (could write to a file in production)
        Dim message
        message = "WARNING: Missing localization key: " & key
        ' Uncomment to log to a file
        ' LogMissingKey key
        
        ' Return the key itself as fallback
        L = key
    End If
End Function

' Helper function to log missing keys (can be implemented to write to a file)
Sub LogMissingKey(key)
    ' This could be implemented to write to a log file in production
    ' For now, we'll just keep as a stub
End Sub

