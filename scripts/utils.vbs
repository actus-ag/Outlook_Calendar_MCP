' utils.vbs - Shared utility functions for Outlook Calendar operations
Option Explicit

' Constants
Const olFolderCalendar = 9
Const olAppointmentItem = 1
Const olMeeting = 1
Const olBusy = 2
Const olTentative = 1
Const olFree = 0
Const olOutOfOffice = 3
Const olResponseAccepted = 3
Const olResponseDeclined = 4
Const olResponseTentative = 2
Const olResponseNotResponded = 5

' Error handling constants
Const ERROR_PREFIX = "ERROR:"  ' Keep these values for backward compatibility
Const SUCCESS_PREFIX = "SUCCESS:"  ' and for parsing in scriptRunner.js

' Check if localization functions are available
Function IsLocalizationActive()
    On Error Resume Next
    Dim result
    result = (GetRef("L") Is Nothing) = False
    IsLocalizationActive = (Err.Number = 0) And result
    Err.Clear
    On Error GoTo 0
End Function

' ===== Outlook Application Management =====

' Creates and returns an Outlook Application object
Function CreateOutlookApplication()
    On Error Resume Next
    Dim outlookApp, errorMsg
    Set outlookApp = CreateObject("Outlook.Application")
    
    If Err.Number <> 0 Then
        If IsLocalizationActive() Then
            errorMsg = L("ERROR_OUTLOOK_APP_FAILED") & Err.Description
        Else
            errorMsg = "Failed to create Outlook Application: " & Err.Description
        End If
        
        WScript.Echo ERROR_PREFIX & errorMsg
        WScript.Quit 1
    End If
    
    Set CreateOutlookApplication = outlookApp
End Function

' Gets the default calendar folder from Outlook
Function GetDefaultCalendar(outlookApp)
    On Error Resume Next
    Dim namespace, calendar, errorMsg
    
    Set namespace = outlookApp.GetNamespace("MAPI")
    If Err.Number <> 0 Then
        If IsLocalizationActive() Then
            errorMsg = L("ERROR_MAPI_NAMESPACE_FAILED") & Err.Description
        Else
            errorMsg = "Failed to get MAPI namespace: " & Err.Description
        End If
        
        WScript.Echo ERROR_PREFIX & errorMsg
        WScript.Quit 1
    End If
    
    Set calendar = namespace.GetDefaultFolder(olFolderCalendar)
    If Err.Number <> 0 Then
        If IsLocalizationActive() Then
            errorMsg = L("ERROR_DEFAULT_CALENDAR_FAILED") & Err.Description
        Else
            errorMsg = "Failed to get default calendar: " & Err.Description
        End If
        
        WScript.Echo ERROR_PREFIX & errorMsg
        WScript.Quit 1
    End If
    
    Set GetDefaultCalendar = calendar
End Function

' Gets a specific calendar folder by name
Function GetCalendarByName(outlookApp, calendarName)
    On Error Resume Next
    Dim namespace, folders, folder, i, errorMsg
    
    Set namespace = outlookApp.GetNamespace("MAPI")
    If Err.Number <> 0 Then
        If IsLocalizationActive() Then
            errorMsg = L("ERROR_MAPI_NAMESPACE_FAILED") & Err.Description
        Else
            errorMsg = "Failed to get MAPI namespace: " & Err.Description
        End If
        
        WScript.Echo ERROR_PREFIX & errorMsg
        WScript.Quit 1
    End If
    
    ' Get default calendar if no name specified
    If calendarName = "" Then
        Set GetCalendarByName = GetDefaultCalendar(outlookApp)
        Exit Function
    End If
    
    ' Try to find the specified calendar
    Set folders = namespace.Folders
    For i = 1 To folders.Count
        Set folder = folders.Item(i)
        If folder.Name = calendarName Then
            Set GetCalendarByName = folder.GetDefaultFolder(olFolderCalendar)
            Exit Function
        End If
    Next
    
    ' Calendar not found
    If IsLocalizationActive() Then
        errorMsg = L("ERROR_CALENDAR_NOT_FOUND") & calendarName
    Else
        errorMsg = "Calendar not found: " & calendarName
    End If
    
    WScript.Echo ERROR_PREFIX & errorMsg
    WScript.Quit 1
End Function

' ===== Date Handling =====

' Converts a date string in MM/DD/YYYY format to a Date object
Function ParseDate(dateStr)
    On Error Resume Next
    Dim errorMsg
    
    ' If localization is active, use the locale-aware function
    If IsLocalizationActive() Then
        ' Use default locale (g_currentLocale) from localization system
        ParseDate = ParseDateByLocale(dateStr, g_currentLocale)
        Exit Function
    End If
    
    ' Default non-localized implementation
    If IsDate(dateStr) Then
        ParseDate = CDate(dateStr)
    Else
        ' Try to parse MM/DD/YYYY format
        Dim parts, month, day, year
        parts = Split(dateStr, "/")
        
        If UBound(parts) = 2 Then
            month = parts(0)
            day = parts(1)
            year = parts(2)
            
            If IsNumeric(month) And IsNumeric(day) And IsNumeric(year) Then
                ParseDate = DateSerial(year, month, day)
            Else
                WScript.Echo ERROR_PREFIX & "Invalid date format. Expected MM/DD/YYYY: " & dateStr
                WScript.Quit 1
            End If
        Else
            WScript.Echo ERROR_PREFIX & "Invalid date format. Expected MM/DD/YYYY: " & dateStr
            WScript.Quit 1
        End If
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo ERROR_PREFIX & "Failed to parse date: " & dateStr & " - " & Err.Description
        WScript.Quit 1
    End If
End Function

' Formats a Date object to MM/DD/YYYY format or locale-specific format
Function FormatDate(dateObj)
    ' If localization is active, use the locale-aware function
    If IsLocalizationActive() Then
        FormatDate = FormatDateByLocale(dateObj, g_currentLocale)
    Else
        ' Default format (MM/DD/YYYY)
        FormatDate = Month(dateObj) & "/" & Day(dateObj) & "/" & Year(dateObj)
    End If
End Function

' Formats a Date object to MM/DD/YYYY HH:MM AM/PM format or locale-specific format
Function FormatDateTime(dateTimeObj)
    ' If localization is active, use the locale-aware function
    If IsLocalizationActive() Then
        FormatDateTime = FormatDateTimeByLocale(dateTimeObj, g_currentLocale)
    Else
        FormatDateTime = FormatDate(dateTimeObj) & " " & FormatTime(dateTimeObj)
    End If
End Function

' Formats a time to HH:MM AM/PM format
Function FormatTime(dateTimeObj)
    Dim hours, minutes, ampm
    
    hours = Hour(dateTimeObj)
    minutes = Minute(dateTimeObj)
    
    If hours >= 12 Then
        ampm = "PM"
        If hours > 12 Then hours = hours - 12
    Else
        ampm = "AM"
        If hours = 0 Then hours = 12
    End If
    
    FormatTime = Right("0" & hours, 2) & ":" & Right("0" & minutes, 2) & " " & ampm
End Function

' ===== JSON Handling =====

' Escapes a string for JSON
Function EscapeJSON(str)
    Dim result, i, charCode, hexStr

    ' First do the standard replacements
    result = Replace(str, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")

    ' Now handle non-ASCII characters (code points > 127)
    Dim finalResult, char
    finalResult = ""

    For i = 1 To Len(result)
        char = Mid(result, i, 1)
        charCode = AscW(char)

        ' If it's a standard ASCII character (0-127), keep it as is
        If charCode >= 0 And charCode <= 127 Then
            finalResult = finalResult & char
        Else
            ' For non-ASCII characters, use Unicode escape sequence
            ' Convert to hex and ensure it's 4 digits with leading zeros
            hexStr = Hex(charCode)
            ' Pad with leading zeros if needed to make it 4 digits
            hexStr = String(4 - Len(hexStr), "0") & hexStr
            finalResult = finalResult & "\u" & hexStr
        End If
    Next

    EscapeJSON = finalResult
End Function

' Converts a VBScript boolean to a JSON boolean (always in English)
Function BoolToJSON(boolValue)
    If boolValue Then
        BoolToJSON = "true"
    Else
        BoolToJSON = "false"
    End If
End Function

' Converts a VBScript array to a JSON array
Function ArrayToJSON(arr)
    Dim i, result

    result = "["
    For i = LBound(arr) To UBound(arr)
        If i > LBound(arr) Then result = result & ","

        If IsNull(arr(i)) Then
            result = result & "null"
        ElseIf IsArray(arr(i)) Then
            result = result & ArrayToJSON(arr(i))
        ElseIf IsObject(arr(i)) Then
            result = result & "null" ' Objects not supported in this simple implementation
        ElseIf VarType(arr(i)) = vbString Then
            result = result & """" & EscapeJSON(arr(i)) & """"
        ElseIf VarType(arr(i)) = vbBoolean Then
            result = result & BoolToJSON(arr(i))
        Else
            result = result & arr(i)
        End If
    Next

    result = result & "]"
    ArrayToJSON = result
End Function

' ===== Outlook Item Conversion =====

' Converts an Outlook appointment item to a JSON string
Function AppointmentToJSON(appointment)
    Dim json, recipients, recipient, i, attendees, attendeeStatus
    
    ' Start building the JSON object
    json = "{"
    
    ' Include EntryID for event identification
    json = json & """id"":""" & EscapeJSON(appointment.EntryID) & ""","
    
    ' Basic properties
    json = json & """subject"":""" & EscapeJSON(appointment.Subject) & ""","
    json = json & """start"":""" & FormatDateTime(appointment.Start) & ""","
    json = json & """end"":""" & FormatDateTime(appointment.End) & ""","
    json = json & """location"":""" & EscapeJSON(appointment.Location) & ""","
    json = json & """body"":""" & EscapeJSON(appointment.Body) & ""","
    json = json & """organizer"":""" & EscapeJSON(appointment.Organizer) & ""","
    json = json & """isRecurring"":" & BoolToJSON(appointment.IsRecurring) & ","

    ' Meeting status
    json = json & """isMeeting"":" & BoolToJSON(appointment.MeetingStatus = olMeeting) & ","
    
    ' Busy status
    Select Case appointment.BusyStatus
        Case olBusy
            If IsLocalizationActive() Then
                json = json & """busyStatus"":""" & L("STATUS_BUSY") & ""","
            Else
                json = json & """busyStatus"":""Busy"","
            End If
        Case olTentative
            If IsLocalizationActive() Then
                json = json & """busyStatus"":""" & L("STATUS_TENTATIVE") & ""","
            Else
                json = json & """busyStatus"":""Tentative"","
            End If
        Case olFree
            If IsLocalizationActive() Then
                json = json & """busyStatus"":""" & L("STATUS_FREE") & ""","
            Else
                json = json & """busyStatus"":""Free"","
            End If
        Case olOutOfOffice
            If IsLocalizationActive() Then
                json = json & """busyStatus"":""" & L("STATUS_OUT_OF_OFFICE") & ""","
            Else
                json = json & """busyStatus"":""Out of Office"","
            End If
        Case Else
            If IsLocalizationActive() Then
                json = json & """busyStatus"":""" & L("STATUS_UNKNOWN") & ""","
            Else
                json = json & """busyStatus"":""Unknown"","
            End If
    End Select
    
    ' Attendees (if it's a meeting)
    If appointment.MeetingStatus = olMeeting Then
        Set recipients = appointment.Recipients
        attendees = ""
        
        For i = 1 To recipients.Count
            Set recipient = recipients.Item(i)
            
            If i > 1 Then attendees = attendees & ","
            
            attendees = attendees & "{"
            attendees = attendees & """name"":""" & EscapeJSON(recipient.Name) & ""","
            attendees = attendees & """email"":""" & EscapeJSON(recipient.Address) & ""","
            
            ' Response status
            Select Case recipient.MeetingResponseStatus
                Case olResponseAccepted
                    If IsLocalizationActive() Then
                        attendeeStatus = L("RESPONSE_ACCEPTED")
                    Else
                        attendeeStatus = "Accepted"
                    End If
                Case olResponseDeclined
                    If IsLocalizationActive() Then
                        attendeeStatus = L("RESPONSE_DECLINED")
                    Else
                        attendeeStatus = "Declined"
                    End If
                Case olResponseTentative
                    If IsLocalizationActive() Then
                        attendeeStatus = L("RESPONSE_TENTATIVE")
                    Else
                        attendeeStatus = "Tentative"
                    End If
                Case olResponseNotResponded
                    If IsLocalizationActive() Then
                        attendeeStatus = L("RESPONSE_NOT_RESPONDED")
                    Else
                        attendeeStatus = "Not Responded"
                    End If
                Case Else
                    If IsLocalizationActive() Then
                        attendeeStatus = L("RESPONSE_UNKNOWN")
                    Else
                        attendeeStatus = "Unknown"
                    End If
            End Select
            
            attendees = attendees & """responseStatus"":""" & attendeeStatus & """"
            attendees = attendees & "}"
        Next
        
        json = json & """attendees"":[" & attendees & "]"
    Else
        json = json & """attendees"":[]"
    End If
    
    ' Close the JSON object
    json = json & "}"
    
    AppointmentToJSON = json
End Function

' Converts a collection of Outlook appointment items to a JSON array
Function AppointmentsToJSON(appointments)
    Dim i, json
    
    json = "["
    
    For i = 1 To appointments.Count
        If i > 1 Then json = json & ","
        json = json & AppointmentToJSON(appointments.Item(i))
    Next
    
    json = json & "]"
    
    AppointmentsToJSON = json
End Function

' ===== Command Line Argument Handling =====

' Gets a command line argument by name
Function GetArgument(name)
    Dim args, i, arg, parts
    
    Set args = WScript.Arguments
    
    For i = 0 To args.Count - 1
        arg = args(i)
        
        If Left(arg, 1) = "/" Or Left(arg, 1) = "-" Then
            parts = Split(Mid(arg, 2), ":", 2)
            
            If UBound(parts) >= 0 Then
                If LCase(parts(0)) = LCase(name) Then
                    If UBound(parts) = 1 Then
                        GetArgument = parts(1)
                    Else
                        GetArgument = "true"
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
    
    GetArgument = ""
End Function

' Checks if a required argument is present
Sub RequireArgument(name)
    Dim value, errorMsg
    
    value = GetArgument(name)
    
    If value = "" Then
        If IsLocalizationActive() Then
            errorMsg = L("ERROR_MISSING_REQUIRED_ARG") & name
        Else
            errorMsg = "Missing required argument: " & name
        End If
        
        WScript.Echo ERROR_PREFIX & errorMsg
        WScript.Quit 1
    End If
End Sub

' ===== Output Formatting =====

' Outputs a success message with JSON data
Sub OutputSuccess(jsonData)
    ' Always use SUCCESS_PREFIX for compatibility with scriptRunner.js
    WScript.Echo SUCCESS_PREFIX & jsonData
End Sub

' Outputs an error message
Sub OutputError(message)
    ' Always use ERROR_PREFIX for compatibility with scriptRunner.js
    WScript.Echo ERROR_PREFIX & message
End Sub
