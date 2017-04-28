Attribute VB_Name = "Other_UTCHelper"
' This module was borrowed from: http://www.vbusers.com/code/codeget.asp?PostID=1&ThreadID=632

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Public Function ConvertLocalToGMT(dtLocalDate As Date) As Date
    Dim lSecsDiff As Long
    
    'Get the GMT time difference
    lSecsDiff = GetLocalToGMTDifference()

    'Return the time in GMT
    ConvertLocalToGMT = DateAdd("s", -lSecsDiff, dtLocalDate)
End Function

Public Function GetLocalToGMTDifference() As Long
    Const TIME_ZONE_ID_INVALID& = &HFFFFFFFF
    Const TIME_ZONE_ID_STANDARD& = 1
    Const TIME_ZONE_ID_UNKNOWN& = 0
    Const TIME_ZONE_ID_DAYLIGHT& = 2
    
    Dim tTimeZoneInf As TIME_ZONE_INFORMATION
    Dim lRet As Long
    Dim lDiff As Long
    
    'Get time zone information
    lRet = GetTimeZoneInformation(tTimeZoneInf)
    
    'Convert differenceto seconds
    lDiff = -tTimeZoneInf.Bias * 60
    GetLocalToGMTDifference = lDiff
    
    'Checks if we are in daylight saving time.
    If lRet = TIME_ZONE_ID_DAYLIGHT& Then
        'If in daylight savings, apply the bias
        If tTimeZoneInf.DaylightDate.wMonth <> 0 Then
            'If tTimeZoneInf.DaylightDate.wMonth = 0 then the daylight,
            'saving time change doesn't occur
            GetLocalToGMTDifference = lDiff - tTimeZoneInf.DaylightBias * 60
        End If
    End If
End Function
