' messages.vbs - Spanish localization strings for Outlook Calendar MCP
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
    g_messages.Add "ERROR_OUTLOOK_APP_FAILED", "Error al crear la aplicación de Outlook: "
    g_messages.Add "ERROR_MAPI_NAMESPACE_FAILED", "Error al obtener el espacio de nombres MAPI: "
    g_messages.Add "ERROR_DEFAULT_CALENDAR_FAILED", "Error al obtener el calendario predeterminado: "
    g_messages.Add "ERROR_CALENDAR_NOT_FOUND", "Calendario no encontrado: "
    g_messages.Add "ERROR_PARSE_DATE", "Error al analizar la fecha: "
    g_messages.Add "ERROR_INVALID_DATE_FORMAT", "Formato de fecha no válido. Se esperaba DD/MM/AAAA: "
    g_messages.Add "ERROR_INVALID_TIME_FORMAT", "Formato de hora no válido: "
    g_messages.Add "ERROR_MISSING_REQUIRED_ARG", "Falta el argumento requerido: "
    g_messages.Add "ERROR_CREATE_EVENT", "Error al crear el evento del calendario: "
    g_messages.Add "ERROR_UPDATE_EVENT", "Error al actualizar el evento del calendario: "
    g_messages.Add "ERROR_DELETE_EVENT", "Error al eliminar el evento del calendario: "
    g_messages.Add "ERROR_EVENT_NOT_FOUND", "Evento no encontrado con ID: "
    g_messages.Add "ERROR_END_TIME_BEFORE_START", "La hora de finalización no puede ser anterior o igual a la hora de inicio"
    g_messages.Add "ERROR_SCRIPT_EXECUTION", "Error en la ejecución del script: "
    g_messages.Add "ERROR_SCRIPT_OUTPUT_PARSE", "Error al analizar la salida del script: "
    g_messages.Add "ERROR_UNEXPECTED_OUTPUT", "Salida inesperada del script: "
    
    ' -------------------------------------------------------------------
    ' Success messages
    ' -------------------------------------------------------------------
    g_messages.Add "SUCCESS_PREFIX", "SUCCESS:"  ' Keep original for parsing
    g_messages.Add "SUCCESS_EVENT_CREATED", "Evento creado exitosamente"
    g_messages.Add "SUCCESS_EVENT_UPDATED", "Evento actualizado exitosamente"
    g_messages.Add "SUCCESS_EVENT_DELETED", "Evento eliminado exitosamente"
    
    ' -------------------------------------------------------------------
    ' Date and time formats
    ' -------------------------------------------------------------------
    g_messages.Add "DATE_FORMAT_MMDDYYYY", "MM/DD/AAAA"
    g_messages.Add "DATE_FORMAT_DDMMYYYY", "DD/MM/AAAA"
    g_messages.Add "DATE_FORMAT_YYYYMMDD", "AAAA-MM-DD"
    g_messages.Add "TIME_FORMAT_HHMM", "HH:MM AM/PM"
    
    ' -------------------------------------------------------------------
    ' Calendar & Event Status Labels
    ' -------------------------------------------------------------------
    g_messages.Add "STATUS_BUSY", "Ocupado"
    g_messages.Add "STATUS_TENTATIVE", "Provisional"
    g_messages.Add "STATUS_FREE", "Libre"
    g_messages.Add "STATUS_OUT_OF_OFFICE", "Fuera de oficina"
    g_messages.Add "STATUS_UNKNOWN", "Desconocido"
    
    ' -------------------------------------------------------------------
    ' Response Status Labels
    ' -------------------------------------------------------------------
    g_messages.Add "RESPONSE_ACCEPTED", "Aceptado"
    g_messages.Add "RESPONSE_DECLINED", "Rechazado"
    g_messages.Add "RESPONSE_TENTATIVE", "Provisional"
    g_messages.Add "RESPONSE_NOT_RESPONDED", "Sin respuesta"
    g_messages.Add "RESPONSE_UNKNOWN", "Desconocido"
    
    ' -------------------------------------------------------------------
    ' UI Labels
    ' -------------------------------------------------------------------
    g_messages.Add "LABEL_SUBJECT", "Asunto"
    g_messages.Add "LABEL_START_DATE", "Fecha de inicio"
    g_messages.Add "LABEL_START_TIME", "Hora de inicio"
    g_messages.Add "LABEL_END_DATE", "Fecha de finalización"
    g_messages.Add "LABEL_END_TIME", "Hora de finalización"
    g_messages.Add "LABEL_LOCATION", "Ubicación"
    g_messages.Add "LABEL_BODY", "Cuerpo"
    g_messages.Add "LABEL_ORGANIZER", "Organizador"
    g_messages.Add "LABEL_ATTENDEES", "Asistentes"
    g_messages.Add "LABEL_CALENDAR", "Calendario"
    g_messages.Add "LABEL_STATUS", "Estado"
    g_messages.Add "LABEL_ID", "ID"
    g_messages.Add "LABEL_IS_RECURRING", "Es recurrente"
    g_messages.Add "LABEL_IS_MEETING", "Es reunión"
    
    ' -------------------------------------------------------------------
    ' Month names
    ' -------------------------------------------------------------------
    g_messages.Add "MONTH_1", "Enero"
    g_messages.Add "MONTH_2", "Febrero"
    g_messages.Add "MONTH_3", "Marzo"
    g_messages.Add "MONTH_4", "Abril"
    g_messages.Add "MONTH_5", "Mayo"
    g_messages.Add "MONTH_6", "Junio"
    g_messages.Add "MONTH_7", "Julio"
    g_messages.Add "MONTH_8", "Agosto"
    g_messages.Add "MONTH_9", "Septiembre"
    g_messages.Add "MONTH_10", "Octubre"
    g_messages.Add "MONTH_11", "Noviembre"
    g_messages.Add "MONTH_12", "Diciembre"
    
    ' -------------------------------------------------------------------
    ' Day names
    ' -------------------------------------------------------------------
    g_messages.Add "DAY_1", "Domingo"
    g_messages.Add "DAY_2", "Lunes"
    g_messages.Add "DAY_3", "Martes"
    g_messages.Add "DAY_4", "Miércoles"
    g_messages.Add "DAY_5", "Jueves"
    g_messages.Add "DAY_6", "Viernes"
    g_messages.Add "DAY_7", "Sábado"
    
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
        message = "ADVERTENCIA: Clave de localización no encontrada: " & key
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

