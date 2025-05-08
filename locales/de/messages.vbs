' messages.vbs - German localization strings for Outlook Calendar MCP
Option Explicit

' Create Dictionary object to store localized strings
Dim g_messages
Set g_messages = CreateObject("Scripting.Dictionary")

' Initialize messages dictionary with all localized strings
Sub InitializeMessages()
    ' -------------------------------------------------------------------
    ' Error messages
    ' -------------------------------------------------------------------
    g_messages.Add "ERROR_PREFIX", "ERROR:"  ' Keep original for parsing
    g_messages.Add "ERROR_OUTLOOK_APP_FAILED", "Fehler beim Erstellen der Outlook-Anwendung: "
    g_messages.Add "ERROR_MAPI_NAMESPACE_FAILED", "Fehler beim Abrufen des MAPI-Namespace: "
    g_messages.Add "ERROR_DEFAULT_CALENDAR_FAILED", "Fehler beim Abrufen des Standardkalenders: "
    g_messages.Add "ERROR_CALENDAR_NOT_FOUND", "Kalender nicht gefunden: "
    g_messages.Add "ERROR_PARSE_DATE", "Fehler beim Analysieren des Datums: "
    g_messages.Add "ERROR_INVALID_DATE_FORMAT", "Ung?ltiges Datumsformat. Erwartet wird DD.MM.YYYY: "
    g_messages.Add "ERROR_INVALID_TIME_FORMAT", "Ung?ltiges Zeitformat: "
    g_messages.Add "ERROR_MISSING_REQUIRED_ARG", "Fehlendes erforderliches Argument: "
    g_messages.Add "ERROR_CREATE_EVENT", "Fehler beim Erstellen des Kalenderereignisses: "
    g_messages.Add "ERROR_UPDATE_EVENT", "Fehler beim Aktualisieren des Kalenderereignisses: "
    g_messages.Add "ERROR_DELETE_EVENT", "Fehler beim L?schen des Kalenderereignisses: "
    g_messages.Add "ERROR_EVENT_NOT_FOUND", "Ereignis mit ID nicht gefunden: "
    g_messages.Add "ERROR_END_TIME_BEFORE_START", "Die Endzeit kann nicht vor oder gleich der Startzeit sein"
    g_messages.Add "ERROR_SCRIPT_EXECUTION", "Fehler bei der Skriptausf?hrung: "
    g_messages.Add "ERROR_SCRIPT_OUTPUT_PARSE", "Fehler beim Analysieren der Skriptausgabe: "
    g_messages.Add "ERROR_UNEXPECTED_OUTPUT", "Unerwartete Skriptausgabe: "
    
    ' -------------------------------------------------------------------
    ' Success messages
    ' -------------------------------------------------------------------
    g_messages.Add "SUCCESS_PREFIX", "SUCCESS:"  ' Keep original for parsing
    g_messages.Add "SUCCESS_EVENT_CREATED", "Ereignis erfolgreich erstellt"
    g_messages.Add "SUCCESS_EVENT_UPDATED", "Ereignis erfolgreich aktualisiert"
    g_messages.Add "SUCCESS_EVENT_DELETED", "Ereignis erfolgreich gel?scht"
    
    ' -------------------------------------------------------------------
    ' Date and time formats
    ' -------------------------------------------------------------------
    g_messages.Add "DATE_FORMAT_MMDDYYYY", "MM/DD/YYYY"
    g_messages.Add "DATE_FORMAT_DDMMYYYY", "DD.MM.YYYY"  ' German format uses dots
    g_messages.Add "DATE_FORMAT_YYYYMMDD", "YYYY-MM-DD"
    g_messages.Add "TIME_FORMAT_HHMM", "HH:MM"  ' Germany typically uses 24-hour format
    
    ' -------------------------------------------------------------------
    ' Calendar & Event Status Labels
    ' -------------------------------------------------------------------
    g_messages.Add "STATUS_BUSY", "Besch?ftigt"
    g_messages.Add "STATUS_TENTATIVE", "Vorl?ufig"
    g_messages.Add "STATUS_FREE", "Frei"
    g_messages.Add "STATUS_OUT_OF_OFFICE", "Abwesend"
    g_messages.Add "STATUS_UNKNOWN", "Unbekannt"
    
    ' -------------------------------------------------------------------
    ' Response Status Labels
    ' -------------------------------------------------------------------
    g_messages.Add "RESPONSE_ACCEPTED", "Angenommen"
    g_messages.Add "RESPONSE_DECLINED", "Abgelehnt"
    g_messages.Add "RESPONSE_TENTATIVE", "Vorl?ufig"
    g_messages.Add "RESPONSE_NOT_RESPONDED", "Nicht beantwortet"
    g_messages.Add "RESPONSE_UNKNOWN", "Unbekannt"
    
    ' -------------------------------------------------------------------
    ' UI Labels
    ' -------------------------------------------------------------------
    g_messages.Add "LABEL_SUBJECT", "Betreff"
    g_messages.Add "LABEL_START_DATE", "Startdatum"
    g_messages.Add "LABEL_START_TIME", "Startzeit"
    g_messages.Add "LABEL_END_DATE", "Enddatum"
    g_messages.Add "LABEL_END_TIME", "Endzeit"
    g_messages.Add "LABEL_LOCATION", "Ort"
    g_messages.Add "LABEL_BODY", "Inhalt"
    g_messages.Add "LABEL_ORGANIZER", "Organisator"
    g_messages.Add "LABEL_ATTENDEES", "Teilnehmer"
    g_messages.Add "LABEL_CALENDAR", "Kalender"
    g_messages.Add "LABEL_STATUS", "Status"
    g_messages.Add "LABEL_ID", "ID"
    g_messages.Add "LABEL_IS_RECURRING", "Ist wiederkehrend"
    g_messages.Add "LABEL_IS_MEETING", "Ist Besprechung"
    
    ' -------------------------------------------------------------------
    ' Month names
    ' -------------------------------------------------------------------
    g_messages.Add "MONTH_1", "Januar"
    g_messages.Add "MONTH_2", "Februar"
    g_messages.Add "MONTH_3", "M?rz"
    g_messages.Add "MONTH_4", "April"
    g_messages.Add "MONTH_5", "Mai"
    g_messages.Add "MONTH_6", "Juni"
    g_messages.Add "MONTH_7", "Juli"
    g_messages.Add "MONTH_8", "August"
    g_messages.Add "MONTH_9", "September"
    g_messages.Add "MONTH_10", "Oktober"
    g_messages.Add "MONTH_11", "November"
    g_messages.Add "MONTH_12", "Dezember"
    
    ' -------------------------------------------------------------------
    ' Day names
    ' -------------------------------------------------------------------
    g_messages.Add "DAY_1", "Sonntag"
    g_messages.Add "DAY_2", "Montag"
    g_messages.Add "DAY_3", "Dienstag"
    g_messages.Add "DAY_4", "Mittwoch"
    g_messages.Add "DAY_5", "Donnerstag"
    g_messages.Add "DAY_6", "Freitag"
    g_messages.Add "DAY_7", "Samstag"
    
    ' -------------------------------------------------------------------
    ' Time parts
    ' -------------------------------------------------------------------
    g_messages.Add "TIME_AM", "AM"  ' Keep original for code compatibility
    g_messages.Add "TIME_PM", "PM"  ' Keep original for code compatibility
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
        message = "WARNUNG: Fehlender Lokalisierungsschl?ssel: " & key
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
