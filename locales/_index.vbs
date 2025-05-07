' _index.vbs - Main localization management file for Outlook Calendar MCP
Option Explicit

' Global variables for localization
Dim g_currentLocale     ' Current locale code (e.g., "en", "es")
Dim g_defaultLocale     ' Default locale code ("en")
Dim g_supportedLocales  ' Array of supported locale codes
Dim g_dateFormatByLocale ' Dictionary mapping locale to preferred date format
Dim g_fso               ' File System Object
Dim g_localesPath       ' Path to locales directory

' Initialize localization system
Sub InitializeLocalization()
    ' Set default locale
    g_defaultLocale = "en"
    
    ' Define supported locales
    g_supportedLocales = Array("en", "es")
    
    ' Create dictionary for date formats by locale
    Set g_dateFormatByLocale = CreateObject("Scripting.Dictionary")
    g_dateFormatByLocale.Add "en", "MM/DD/YYYY" ' English (US) format
    g_dateFormatByLocale.Add "es", "DD/MM/YYYY" ' Spanish format
    
    ' Create File System Object
    Set g_fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get path to the locales directory
    g_localesPath = g_fso.GetParentFolderName(WScript.ScriptFullName)
    
    ' Detect and set current locale
    g_currentLocale = DetectLocale()
    
    ' Load the appropriate language file
    LoadLanguageFile g_currentLocale
End Sub

' Detects the locale to use based on command line arguments or system settings
Function DetectLocale()
    Dim locale, shell, regLocale
    
    ' First priority: Check for locale parameter from command line
    locale = GetArgument("locale")
    
    ' Second priority: Try to get system locale from registry
    If locale = "" Then
        On Error Resume Next
        Set shell = CreateObject("WScript.Shell")
        regLocale = shell.RegRead("HKEY_CURRENT_USER\Control Panel\International\LocaleName")
        If Err.Number = 0 And regLocale <> "" Then
            ' Extract the language code (first two characters) from the locale name
            locale = Left(regLocale, 2)
        End If
        On Error GoTo 0
    End If
    
    ' Default to English if no locale detected or if detected locale is not supported
    If locale = "" Or Not IsLocaleSupported(locale) Then
        locale = g_defaultLocale
    End If
    
    DetectLocale = locale
End Function

' Checks if a locale is supported
Function IsLocaleSupported(locale)
    Dim i
    
    For i = LBound(g_supportedLocales) To UBound(g_supportedLocales)
        If LCase(locale) = LCase(g_supportedLocales(i)) Then
            IsLocaleSupported = True
            Exit Function
        End If
    Next
    
    IsLocaleSupported = False
End Function

' Loads the appropriate language file
Sub LoadLanguageFile(locale)
    Dim languageFilePath
    
    ' Ensure the locale is supported, otherwise use default
    If Not IsLocaleSupported(locale) Then
        locale = g_defaultLocale
    End If
    
    ' Construct path to language file
    languageFilePath = g_fso.BuildPath(g_localesPath, locale & "\messages.vbs")
    
    ' Check if the file exists
    If Not g_fso.FileExists(languageFilePath) Then
        ' Fall back to English if file doesn't exist
        languageFilePath = g_fso.BuildPath(g_localesPath, g_defaultLocale & "\messages.vbs")
    End If
    
    ' Include the language file
    ExecuteGlobal g_fso.OpenTextFile(languageFilePath, 1).ReadAll
End Sub

' Gets the preferred date format for the current locale
Function GetPreferredDateFormat()
    If g_dateFormatByLocale.Exists(g_currentLocale) Then
        GetPreferredDateFormat = g_dateFormatByLocale(g_currentLocale)
    Else
        GetPreferredDateFormat = g_dateFormatByLocale(g_defaultLocale)
    End If
End Function

' Parses a date string according to the locale-specific format
Function ParseDateByLocale(dateStr, locale)
    On Error Resume Next
    Dim parts, day, month, year, result
    
    ' Default to current locale if not specified
    If locale = "" Then locale = g_currentLocale
    
    ' Handle the case where date is already a Date object
    If IsDate(dateStr) Then
        ParseDateByLocale = CDate(dateStr)
        Exit Function
    End If
    
    ' Split the date string
    parts = Split(dateStr, "/")
    
    ' Make sure we have three parts
    If UBound(parts) <> 2 Then
        Err.Raise 1000, "ParseDateByLocale", L("ERROR_INVALID_DATE_FORMAT") & dateStr
        Exit Function
    End If
    
    ' Parse based on locale format preference
    If g_dateFormatByLocale(locale) = "MM/DD/YYYY" Then
        ' MM/DD/YYYY format (English US)
        month = parts(0)
        day = parts(1)
        year = parts(2)
    ElseIf g_dateFormatByLocale(locale) = "DD/MM/YYYY" Then
        ' DD/MM/YYYY format (Spanish, etc.)
        day = parts(0)
        month = parts(1)
        year = parts(2)
    Else
        ' Default to MM/DD/YYYY
        month = parts(0)
        day = parts(1)
        year = parts(2)
    End If
    
    ' Create date if all parts are numeric
    If IsNumeric(month) And IsNumeric(day) And IsNumeric(year) Then
        result = DateSerial(year, month, day)
        
        ' Check for errors
        If Err.Number <> 0 Then
            Err.Raise 1001, "ParseDateByLocale", L("ERROR_PARSE_DATE") & dateStr & " - " & Err.Description
            Exit Function
        End If
        
        ParseDateByLocale = result
    Else
        Err.Raise 1002, "ParseDateByLocale", L("ERROR_INVALID_DATE_FORMAT") & dateStr
    End If
End Function

' Formats a date according to the locale-specific format
Function FormatDateByLocale(dateObj, locale)
    ' Default to current locale if not specified
    If locale = "" Then locale = g_currentLocale
    
    ' Format based on locale preference
    If g_dateFormatByLocale(locale) = "MM/DD/YYYY" Then
        FormatDateByLocale = Month(dateObj) & "/" & Day(dateObj) & "/" & Year(dateObj)
    ElseIf g_dateFormatByLocale(locale) = "DD/MM/YYYY" Then
        FormatDateByLocale = Day(dateObj) & "/" & Month(dateObj) & "/" & Year(dateObj)
    Else
        ' Default MM/DD/YYYY
        FormatDateByLocale = Month(dateObj) & "/" & Day(dateObj) & "/" & Year(dateObj)
    End If
End Function

' Formats a date and time according to the locale-specific format
Function FormatDateTimeByLocale(dateTimeObj, locale)
    FormatDateTimeByLocale = FormatDateByLocale(dateTimeObj, locale) & " " & FormatTime(dateTimeObj)
End Function

' Gets a command line argument by name (from utils.vbs)
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

' Gets localized month name for a month number
Function GetMonthName(monthNumber)
    If monthNumber >= 1 And monthNumber <= 12 Then
        GetMonthName = L("MONTH_" & monthNumber)
    Else
        GetMonthName = ""
    End If
End Function

' Gets localized day name for a day of week number (1=Sunday, 7=Saturday)
Function GetDayName(dayNumber)
    If dayNumber >= 1 And dayNumber <= 7 Then
        GetDayName = L("DAY_" & dayNumber)
    Else
        GetDayName = ""
    End If
End Function

' Initialize the localization system
InitializeLocalization

