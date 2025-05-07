# Multilanguage Support for Outlook Calendar MCP

This document provides information on how to use, extend, and work with the localization system in the Outlook Calendar MCP application.

## Table of Contents

- [Overview](#overview)
- [Architecture](#architecture)
- [Using the Localization System](#using-the-localization-system)
- [Date Formats by Locale](#date-formats-by-locale)
- [Adding a New Language](#adding-a-new-language)
- [Localization Functions](#localization-functions)
- [Examples](#examples)
- [Fallback Mechanism](#fallback-mechanism)

## Overview

The Outlook Calendar MCP application supports multiple languages, allowing users to interact with the calendar in their preferred language. The system handles:

- Localized user-facing text and error messages
- Date and time formats appropriate for different regions
- Locale detection based on system settings or command-line parameters
- Fallback to English when a translation is not available

Currently supported languages:
- English (en) - Default
- Spanish (es)

## Architecture

The localization system consists of three main components:

1. **Language Files** (`locales/[lang-code]/messages.vbs`): 
   Contains key-value pairs for all strings in a specific language.

2. **Localization Index** (`locales/_index.vbs`): 
   Manages locale detection, loading appropriate language files, and provides utility functions.

3. **Integration Layer**: 
   Modifications to `scriptRunner.js` and VBS scripts to use the localization functions.

### Directory Structure

```
locales/
  ├── _index.vbs            # Main localization management file
  ├── en/
  │   └── messages.vbs      # English strings
  ├── es/
  │   └── messages.vbs      # Spanish strings
  └── README.md             # This documentation
```

### Flow

1. The user specifies a locale or the system detects it automatically
2. `_index.vbs` loads the appropriate language file
3. Scripts call the `L()` function to get localized strings
4. Date handling functions use locale-specific formats

## Using the Localization System

### Specifying a Locale

You can specify a locale in several ways:

1. **Command Line Parameter**:
   ```
   outlook-calendar-mcp list_events --startDate="05/07/2025" --locale="es"
   ```

2. **JavaScript API**:
   ```javascript
   // From JavaScript
   const events = await listEvents("05/07/2025", null, null, "es");
   ```

3. **Environment Variable**:
   ```bash
   # In bash/sh
   export LOCALE=es
   outlook-calendar-mcp list_events --startDate="05/07/2025"
   ```
   
   ```powershell
   # In PowerShell
   $env:LOCALE="es"
   outlook-calendar-mcp list_events --startDate="05/07/2025"
   ```
   
   ```fish
   # In fish shell
   set -x LOCALE es
   outlook-calendar-mcp list_events --startDate="05/07/2025"
   ```

If no locale is specified, the system will:
1. Try to detect the system locale from the registry
2. Default to English ("en") if detection fails or the detected locale is not supported

## Date Formats by Locale

Different regions use different date formats. The localization system supports the following formats:

| Locale | Format | Example |
|--------|--------|---------|
| en     | MM/DD/YYYY | 05/07/2025 |
| es     | DD/MM/YYYY | 07/05/2025 |

### Important Notes on Date Formatting

- When using a specific locale, provide dates in the format appropriate for that locale
- The system will handle parsing and converting to the internal format
- Dates are always displayed in the format appropriate for the current locale

## Adding a New Language

To add support for a new language:

1. Create a new directory under `locales/` with the language code (e.g., `locales/fr/` for French)

2. Create a `messages.vbs` file in that directory with the same structure as the English version:

   ```vbs
   ' messages.vbs - French localization strings
   Option Explicit

   ' Create Dictionary object to store localized strings
   Dim g_messages
   Set g_messages = CreateObject("Scripting.Dictionary")

   ' Initialize messages dictionary with all localized strings
   Sub InitializeMessages()
       ' Error messages
       g_messages.Add "ERROR_PREFIX", "ERROR:"  ' Keep this value
       g_messages.Add "ERROR_OUTLOOK_APP_FAILED", "Échec de la création de l'application Outlook: "
       ' ... Add all other translations
   End Sub

   ' Function to retrieve localized string for a given key
   Function L(key)
       ' ... Keep this function as is
   End Function
   ```

3. Update the supported locales array in `_index.vbs`:

   ```vbs
   ' Define supported locales
   g_supportedLocales = Array("en", "es", "fr")  ' Add your new locale
   ```

4. Add the date format for your locale:

   ```vbs
   ' Create dictionary for date formats by locale
   Set g_dateFormatByLocale = CreateObject("Scripting.Dictionary")
   g_dateFormatByLocale.Add "en", "MM/DD/YYYY"
   g_dateFormatByLocale.Add "es", "DD/MM/YYYY"
   g_dateFormatByLocale.Add "fr", "DD/MM/YYYY"  ' Add your locale's format
   ```

## Localization Functions

The localization system provides several key functions:

### String Localization

- `L(key)` - Returns the localized string for the given key
  ```vbs
  ' Example:
  Dim message
  message = L("ERROR_CALENDAR_NOT_FOUND") & calendarName
  ```

### Date Handling

- `ParseDateByLocale(dateStr, locale)` - Parses a date string according to locale-specific format
  ```vbs
  ' Example:
  Dim dateObj
  dateObj = ParseDateByLocale("07/05/2025", "es")  ' Parses as May 7, 2025 in Spanish format
  ```

- `FormatDateByLocale(dateObj, locale)` - Formats a date according to locale-specific format
  ```vbs
  ' Example:
  Dim dateStr
  dateStr = FormatDateByLocale(Now, "es")  ' Formats current date in Spanish format
  ```

### Locale Detection

- `DetectLocale()` - Detects the current locale from command line or system settings
  ```vbs
  ' Example:
  Dim currentLocale
  currentLocale = DetectLocale()
  ```

- `IsLocaleSupported(locale)` - Checks if a locale is supported
  ```vbs
  ' Example:
  If IsLocaleSupported("fr") Then
      ' Do something
  End If
  ```

## Examples

### Example 1: Basic String Localization

```vbs
' Check if a calendar exists
If Not CalendarExists(calendarName) Then
    OutputError L("ERROR_CALENDAR_NOT_FOUND") & calendarName
    WScript.Quit 1
End If
```

### Example 2: Date Handling with Localization

```vbs
' Parse a date using the current locale
Function ParseEventDate(dateStr)
    On Error Resume Next
    
    ' Use the current locale (g_currentLocale)
    Dim dateObj
    dateObj = ParseDateByLocale(dateStr, g_currentLocale)
    
    If Err.Number <> 0 Then
        OutputError L("ERROR_PARSE_DATE") & dateStr
        WScript.Quit 1
    End If
    
    ParseEventDate = dateObj
End Function
```

### Example 3: Creating a Meeting with Localized Status

```vbs
' Format the meeting status in the user's language
Function GetMeetingStatusDescription(statusCode)
    Select Case statusCode
        Case olResponseAccepted
            GetMeetingStatusDescription = L("RESPONSE_ACCEPTED")
        Case olResponseDeclined
            GetMeetingStatusDescription = L("RESPONSE_DECLINED")
        Case olResponseTentative
            GetMeetingStatusDescription = L("RESPONSE_TENTATIVE")
        Case Else
            GetMeetingStatusDescription = L("RESPONSE_UNKNOWN")
    End Select
End Function
```

## Fallback Mechanism

The localization system includes a robust fallback mechanism:

1. If a requested key is not found in the current locale's dictionary:
   - The system will look for it in the English dictionary
   - If it's not found there either, the key itself is returned

2. If a requested locale is not supported:
   - The system falls back to English
   - A warning is logged (depending on settings)

3. If a locale file is missing or cannot be loaded:
   - The system falls back to English
   - No error is shown to maintain functionality

This ensures the application will always function, even if translations are incomplete or a locale is not fully supported.

### Reporting Missing Translations

To help improve translations, missing keys are logged in development mode. You can enable this by:

1. Setting `LogMissingKey` to write to a file
2. Uncommenting the log line in the `L()` function

The logged information can then be used to identify and add missing translations.

