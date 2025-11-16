Attribute VB_Name = "Tools_DateTime"
Option Explicit

#If VBA7 Then
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long
#Else
Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long
#End If
Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type


Public Function ANTIConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo utc_ErrorHandling

    #If Mac Then
    'ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
    #Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    With utc_TimeZoneInfo
        .utc_Bias = -.utc_Bias
        .utc_DaylightBias = -.utc_DaylightBias
    End With
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate

    ANTIConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
    #End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
    utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
    utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
    utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
    utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
    utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
    utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
    utc_DateToSystemTime.utc_wMilliseconds = 0
End Function
Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function
Function DateToIso(Dt As Date, Optional Timezone As Long) As String
    DateToIso = Format(Dt, "yyyy-mm-ssTHH:mm:ss") & IIf(Timezone < 0, "", "+") & Format(Timezone, "00") & ":00"
End Function
Function TimeFromTFormat(isoDateTime As String) As Date
    '"2024-10-14T10:43:32.798988"
    '"2024-10-17T10:27:28.361473-07:00"
    '"2024-10-17T10:27:28.361473+03:00"
    '"2025-02-13T15:37:00.000Z"
    Dim p() As String, ss() As String
    p = Split(isoDateTime, "T")
    ss = Split(p(0), "-")
    TimeFromTFormat = DateSerial(ss(0), ss(1), ss(2))
    ss = Split(p(1), ":")
    If InStr(p(1), "+") Then
        ss(0) = CDbl(ss(0)) - CDbl(Split(ss(2), "+")(1))
        ss(1) = CDbl(ss(1)) - CDbl(ss(3))
        ss(2) = Split(ss(2), "+")(0)
    ElseIf InStr(p(1), "-") Then
        ss(0) = CDbl(ss(0)) + CDbl(Split(ss(2), "-")(1))
        ss(1) = CDbl(ss(1)) + CDbl(ss(3))
        ss(2) = Split(ss(2), "-")(0)
    ElseIf InStr(p(1), "Z") Then
        If InStr(ss(2), ".") Then
            ss(0) = CDbl(ss(0)) - CDbl(Split(ss(2), ".")(0))
            ss(1) = CDbl(ss(1)) - Val(Split(ss(2), ".")(1))
        End If
        ss(2) = Split(ss(2), "Z")(0)
    End If
    TimeFromTFormat = TimeFromTFormat + TimeSerial(ss(0), ss(1), ss(2))
End Function
Function ExtractTimezoneDiff(isoDateTime As String) As Single
    '"2024-10-17T10:27:28.361473-07:00"
    '"2024-10-17T10:27:28.361473+03:00"
    ExtractTimezoneDiff = Mid$(isoDateTime, Len(isoDateTime) - 5, 1) & "1" * _
                (CDbl(Mid$(isoDateTime, Len(isoDateTime) - 4, 2)) + CDbl(Mid$(isoDateTime, Len(isoDateTime) - 1) / 60))
End Function


