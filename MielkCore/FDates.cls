VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FDates"
'----------------------------------------------------------------------------------------------------------
Private Const DEFAULT_DATE_TIME_FORMAT As String = "YYYY-MM-DD HH:MM:SS"
Private Const DEFAULT_DATE_ONLY_FORMAT As String = "YYYY-MM-DD"
Private Const DEFAULT_TIME_ONLY_FORMAT As String = "HH:MM:SS"
'----------------------------------------------------------------------------------------------------------
Private Const ISO_TIME_REGEX_PATTERN As String = "((\d{4})-([01]\d)-([0-3]\d)T([0-2]\d):([0-5]\d):([0-5]\d)\.(\d+([+-][0-2]\d:[0-5]" & _
                                                 "\d|Z)))|((\d{4})-([01]\d)-([0-3]\d)T([0-2]\d):([0-5]\d):([0-5]\d([+-][0-2]\d:[0-5]\d|Z)))" & _
                                                 "|((\d{4})-([01]\d)-([0-3]\d)T([0-2]\d):([0-5]\d)(([+-][0-2]\d:[0-5]\d|Z)))"
Private Const TIMEPART_DISALLOWED_REGEX_PATTERN As String = "^\s*(\d{4})\s*(?:-|\.|\/|\\)\s*(\d{2})\s*(?:-|\.|\/|\\)\s*(\d{2})\s*$"
Private Const TIMEPART_ALLOWED_REGEX_PATTERN As String = "^\s*(\d{4})\s*(?:-|\.|\/|\\|_)\s*(\d{2})\s*(?:-|\.|\/|\\|_)\s*(\d{2})\s*" & _
                                                         "(?:|(\d{2})\s*(?:-|\.|\/|\\|_|:)\s*(\d{2})\s*(?:|(?:-|\.|\/|\\|_|:)\s*(\d{2})\s*))$"
Private Const TIMEPART_DISALLOWED_REVERSED_REGEX_PATTERN As String = "^\s*(\d{2})\s*(?:-|\.|\/|\\)\s*(\d{2})\s*(?:-|\.|\/|\\)\s*(\d{4})\s*$"
Private Const TIMEPART_ALLOWED_REVERSED_REGEX_PATTERN As String = "^\s*(\d{2})\s*(?:-|\.|\/|\\|_)\s*(\d{2})" & _
                                    "\s*(?:-|\.|\/|\\|_)\s*(\d{4})\s*(|\d{2})\s*(?:-|\.|\/|\\|_|:)\s*(\d{2})" & _
                                    "\s*(?:|(?:-|\.|\/|\\|_|:)\s*(\d{2})\s*)$"
Private Const MMDDYY_REGEX_PATTERN As String = "^\s*(\d{2})[\/\.](\d{2})[\/\.](\d{2})\s*$"
'----------------------------------------------------------------------------------------------------------


Public Function isDateString(ByVal value As Variant, Optional allowTimePart As Boolean = True, Optional ByRef d As Date) As Boolean
    Dim patterns As VBA.Collection
    Dim pattern As Variant
    Dim offset As Long
    Dim arrDateParts As Variant
    '------------------------------------------------------------------------------------------------------

    Set patterns = New VBA.Collection
    With patterns
        Call .Add(TIMEPART_DISALLOWED_REGEX_PATTERN)
        Call .Add(TIMEPART_DISALLOWED_REVERSED_REGEX_PATTERN)
        
        If allowTimePart Then
            Call .Add(ISO_TIME_REGEX_PATTERN)
            Call .Add(TIMEPART_ALLOWED_REGEX_PATTERN)
            Call .Add(TIMEPART_ALLOWED_REVERSED_REGEX_PATTERN)
        End If
    End With
    
    For Each pattern In patterns
        If f.regex.checkIfMatch(VBA.CStr(value), VBA.CStr(pattern)) Then
            arrDateParts = f.regex.getMatchArray(VBA.CStr(value), VBA.CStr(pattern))
            If arrDateParts(LBound(arrDateParts)) = value Then offset = 1
            d = convertDatePartsArrayToDate(arrDateParts, offset)
            isDateString = True
            Exit For
        End If
    Next pattern
    
    
    '[Other data formats. Check if they matches VBA built-in date formats.
    If Not isDateString Then
        If VBA.IsDate(value) Then
            d = VBA.CDate(value)
            isDateString = True
        End If
    End If
    

End Function


Public Function isTimeString(ByVal value As String) As Boolean
    Const REGEX_PATTERN As String = "^([01][0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])?$"
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------

    If objRegex Is Nothing Then
        Set objRegex = f.regex.Create(REGEX_PATTERN)
    End If
    isTimeString = objRegex.test(VBA.trim$(value))

    'Check if there are any blank characters
    If Not isTimeString Then
        isTimeString = objRegex.test(VBA.trim$(value))
    End If

End Function

Public Function getDateFromString(ByVal value As String) As Variant
    Dim d As Date
    '------------------------------------------------------------------------------------------------------
    If isDateString(value, True, d) Then
        getDateFromString = d
    Else
        getDateFromString = Empty
    End If
End Function

Private Function convertDatePartsArrayToDate(arr As Variant, Optional offset As Long = 1) As Date
    Dim yearFirst As Boolean
    '------------------------------------------------------------------------------------------------------
    Dim year As Long
    Dim month As Long
    Dim day As Long
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long
    '------------------------------------------------------------------------------------------------------
    
    yearFirst = VBA.CLng(arr(LBound(arr) + offset)) > 1000
    
    If yearFirst Then
        year = VBA.CLng(arr(LBound(arr) + offset))
        day = VBA.CLng(arr(LBound(arr) + 2 + offset))
    Else
        year = VBA.CLng(arr(LBound(arr) + offset))
        day = VBA.CLng(arr(LBound(arr) + offset))
    End If
    
    month = VBA.CLng(arr(LBound(arr) + 1 + offset))
    If f.arrays.getSize(arr) >= 6 Then
        hours = VBA.CLng(arr(LBound(arr) + 3 + offset))
        minutes = VBA.CLng(arr(LBound(arr) + 4 + offset))
        seconds = VBA.CLng(arr(LBound(arr) + 5 + offset))
    End If
    
    convertDatePartsArrayToDate = VBA.DateSerial(year, month, day) + VBA.TimeSerial(hours, minutes, seconds)
    
End Function

Private Function getDateFromFullMonthNameDateString(value As String) As Variant
    Const REGEX_PATTERN As String = "^\s*\d{1,2}\s*\(January|February|March|April|May|June|July|August|September|October|November|December)\s*\d{4}\s*$"
    '------------------------------------------------------------------------------------------------------
    Dim dateParts As Variant
    '------------------------------------------------------------------------------------------------------
    
    If f.regex.checkIfMatch(value, REGEX_PATTERN) Then
        dateParts = f.regex.getMatchArray(value, REGEX_PATTERN)
        getDateFromFullMonthNameDateString = VBA.DateSerial( _
                                dateParts(LBound(dateParts)), _
                                dateParts(LBound(dateParts) + 1), _
                                dateParts(LBound(dateParts) + 2))
    End If
    
End Function

Public Function fromString(ByVal value As String) As Date
    Const ILLEGAL_STRING_ERR_DESC As String = "Date/Time string was expected"
    '------------------------------------------------------------------------------------------------------
    Dim d As Date
    '------------------------------------------------------------------------------------------------------

    If isDDMMYYYYDate(value) Then
        fromString = fromDDMMYYYYString(value)
    ElseIf f.regex.checkIfMatch(value, MMDDYY_REGEX_PATTERN) Then
        fromString = fromDDMMYYString(value)
    ElseIf VBA.IsDate(value) Then
        fromString = VBA.CDate(value)
    ElseIf isDateString(value) Then
        On Error Resume Next
        fromString = VBA.CDate(value)
        If fromString > 0 Then
        Else
            fromString = VBA.CDate(VBA.replace(value, ".", "-"))
        End If
        On Error GoTo 0
    ElseIf isTimeString(value) Then
        fromString = VBA.CDate(value)
    Else
        On Error Resume Next
        d = fromIso8601(value)
        If d > 0 Then
            fromString = d
        Else
            Call Exceptions.IllegalValue.raiseWithCustomMessage(ILLEGAL_STRING_ERR_DESC)
        End If
        On Error GoTo 0
    End If

End Function

Private Function isDDMMYYYYDate(value As String) As Boolean
    Const REGEX_PATTERN As String = "^\d{2}[-|.|\/]\d{2}[-|.|\/]\d{4}$"
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------

    If objRegex Is Nothing Then
        Set objRegex = f.regex.Create(REGEX_PATTERN)
    End If
    isDDMMYYYYDate = objRegex.test(value)

End Function

Public Function fromYYYYMMDDString(ByVal value As String) As Date
    Dim days As Long
    Dim months As Long
    Dim years As Long
    '------------------------------------------------------------------------------------------------------
    If VBA.Len(value) = 8 Then
        years = VBA.left$(value, 4)
        months = VBA.mid$(value, 5, 2)
        days = VBA.right$(value, 2)
        fromYYYYMMDDString = VBA.DateSerial(years, months, days)
    End If
End Function

Private Function fromDDMMYYYYString(ByVal value As String) As Date
    Dim days As Long
    Dim months As Long
    Dim years As Long
    '-----------------------------------------------------------------------------------------------------
    value = f.strings.onlyDigits(value, False)
    If VBA.Len(value) = 8 Then
        days = VBA.left$(value, 2)
        months = VBA.mid$(value, 3, 2)
        years = VBA.right$(value, 4)
        fromDDMMYYYYString = VBA.DateSerial(years, months, days)
    End If
End Function

Private Function fromDDMMYYString(ByVal value As String) As Date
    Dim days As Long
    Dim months As Long
    Dim years As Long
    '-----------------------------------------------------------------------------------------------------
    value = f.strings.onlyDigits(value, False)
    If VBA.Len(value) = 6 Then
        days = VBA.left$(value, 2)
        months = VBA.mid$(value, 3, 2)
        years = 2000 + VBA.right$(value, 2)
        fromDDMMYYString = VBA.DateSerial(years, months, days)
    End If
End Function

Public Function fromMMDDYYYYString(text As String) As Date
    Const REGEX_PATTERN As String = "(\d{1,2})\/(\d{1,2})\/(\d{2,4}) "
    '------------------------------------------------------------------------------------------------------
    Dim parts As Variant
    '------------------------------------------------------------------------------------------------------
    
    parts = f.regex.getMatchArray(text, REGEX_PATTERN)
    If f.arrays.isDefined(parts) Then
        fromMMDDYYYYString = VBA.DateSerial(VBA.IIf(parts(3) < 100, parts(3) + 2000, parts(3)), parts(1), parts(2))
    End If
    
End Function

Public Function isYYYYMMDDString(value As String) As BookmarkEnum
    Const REGEX_PATTERN As String = "^\d{4}(?:-|\.|\/)\d{2}(?:-|\.|\/)\d{2}$"
    '------------------------------------------------------------------------------------------------------
    If VBA.Len(value) = 10 Then
        isYYYYMMDDString = f.regex.checkIfMatch(value, REGEX_PATTERN)
    End If
End Function


'Public Function isIso8601DateTime(value As String) As Boolean
'    Const REGEX_PATTERN As String = "^([\+-]?\d{4}(?!\d{2}\b))((-?)((0[1-9]|1[0-2])(\3([12]\d|0[1-9]|3[01]))?" & _
'            "|W([0-4]\d|5[0-2])(-?[1-7])?|(00[1-9]|0[1-9]\d|[12]\d{2}|3([0-5]\d|6[1-6])))([T\s]((([01]\d|2[0-3])" & _
'            "((:?)[0-5]\d)?|24\:?00)([\.,]\d+(?!:))?)?(\17[0-5]\d([\.,]\d+)?)?([zZ]|([\+-])([01]\d|2[0-3]):?([0-5]\d)?)?)?)?$"
'    Static objRegex As Object
'    '------------------------------------------------------------------------------------------------------
'    If objRegex Is Nothing Then
'        Set objRegex = F.Regex.create(REGEX_PATTERN)
'    End If
'    isIso8601DateTime = objRegex.test(value)
'End Function






'***** WITHOUT UNIT TESTS *****
Public Function getTimestampString() As String
    Dim hundrethsSecond As Double
    Dim currentTimestamp As Double
    '------------------------------------------------------------------------------------------------------

    currentTimestamp = VBA.Timer
    hundrethsSecond = VBA.CInt((currentTimestamp - Int(currentTimestamp)) * 1000)
    getTimestampString = VBA.format(VBA.Now, "YYYYMMDDHHMMSS" & f.strings.addLeadingZeros(hundrethsSecond, 3))
End Function

Public Function toMMSSFormat(value As Double) As String
    Dim minutes As Integer
    Dim seconds As Integer
    '------------------------------------------------------------------------------------------------------
    minutes = value \ 60
    seconds = value Mod 60
    toMMSSFormat = f.strings.addLeadingZeros(minutes) & ":" & f.strings.addLeadingZeros(seconds)
End Function

Public Function getTimeOnly(value As Date) As Date
    Dim dateOnly As Long
    '------------------------------------------------------------------------------------------------------
    dateOnly = VBA.Int(value)
    getTimeOnly = value - dateOnly
End Function

Public Function fromIso8601(dateString As String) As Date
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    Dim hour As Integer
    Dim minute As Integer
    Dim second As Integer
    '------------------------------------------------------------------------------------------------------

    year = VBA.left$(dateString, 4) * 1
    month = VBA.mid$(dateString, 6, 2) * 1
    day = VBA.mid$(dateString, 9, 2) * 1
    hour = VBA.mid$(dateString, 12, 2) * 1
    minute = VBA.mid$(dateString, 15, 2) * 1
    second = VBA.mid$(dateString, 18, 2) * 1

    fromIso8601 = VBA.DateSerial(year, month, day) + VBA.TimeSerial(hour, minute, second)

End Function

Public Function toUnix(dt As Variant) As Long
    If Not f.Variables.isUndefined(dt) Then
        toUnix = VBA.DateDiff("s", "1/1/1970", dt)
    End If
End Function

Public Function fromUnix(timestamp As Long) As Date
    fromUnix = VBA.DateAdd("s", timestamp, #1/1/1970#)
End Function

Public Function getMonthNumberFromName(name As String) As Integer
    Static regexes As Variant
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    If VBA.IsEmpty(regexes) Then
        ReDim regexes(1 To 12)
        Set regexes(1) = f.regex.Create("^[ ]*(Jan(uary|)|sty(cze.|))[ ]*$")
        Set regexes(2) = f.regex.Create("^[ ]*(Feb(ruary|)|lut(y|))[ ]*$")
        Set regexes(3) = f.regex.Create("^[ ]*Mar(ch|zec|)[ ]*$")
        Set regexes(4) = f.regex.Create("^[ ]*(Apr(il|)|kwi(ecie.|))[ ]*$")
        Set regexes(5) = f.regex.Create("^[ ]*Ma(y|j)[ ]*$")
        Set regexes(6) = f.regex.Create("^[ ]*(Jun(e|)|cze(rwiec|))[ ]*$")
        Set regexes(7) = f.regex.Create("^[ ]*(Jul(y|)|lip(iec|))[ ]*$")
        Set regexes(8) = f.regex.Create("^[ ]*(Aug(ust|)|sie(rpie.|))[ ]*$")
        Set regexes(9) = f.regex.Create("^[ ]*(Sep(tember|)|wrz(esie.|))[ ]*$")
        Set regexes(10) = f.regex.Create("^[ ]*(Oct(ober|)|pa.(dziernik|))[ ]*$")
        Set regexes(11) = f.regex.Create("^[ ]*(Nov(ember|)|lis(topad|))[ ]*$")
        Set regexes(12) = f.regex.Create("^[ ]*(Dec(ember|)|gru(dzie.|))[ ]*$")
    End If

    For i = LBound(regexes) To UBound(regexes)
        If regexes(i).test(name) Then
            getMonthNumberFromName = i
            Exit For
        End If
    Next i

End Function

Public Function getMonthName(d As Date, _
                                Optional shortName As Boolean = True, _
                                Optional nameType As MonthNameTypeEnum = MonthName_ByRegionalSettings) As String
    Dim index As Long
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------
    
    If nameType = MonthName_ByRegionalSettings Then
        If shortName Then
            getMonthName = VBA.format(d, "mmm")
        Else
            getMonthName = VBA.format(d, "mmmm")
        End If
    Else
        If shortName Then
            Select Case nameType
                Case MonthName_English:     arr = VBA.Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                                                            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
            End Select
        Else
            Select Case nameType
                Case MonthName_English:     arr = VBA.Array("January", "February", "March", "April", "May", "June", _
                                                            "July", "August", "September", "October", "November", "December")
            End Select
        End If
    End If
    
    index = VBA.month(d)
    If f.arrays.isDefined(arr) Then
        getMonthName = arr(index + LBound(arr) - 1)
    End If
                                
End Function

Public Function getNextMonthStart(d As Date) As Date
    Dim tempDate As Date
    '------------------------------------------------------------------------------------------------------
    tempDate = VBA.DateAdd("m", 1, d)
    getNextMonthStart = VBA.DateSerial(VBA.year(tempDate), VBA.month(tempDate), 1)
End Function

Public Function getMonthFirstDay(d As Date) As Date
    Dim year As Long
    Dim month As Long
    '------------------------------------------------------------------------------------------------------
    year = VBA.year(d)
    month = VBA.month(d)
    getMonthFirstDay = VBA.DateSerial(year, month, 1)
End Function

Public Function getMonthLastDay(d As Date) As Date
    getMonthLastDay = VBA.DateAdd("m", 1, VBA.DateSerial(VBA.year(d), VBA.month(d), 1)) - 1
End Function


'
'Public Function countWorkingDays(startDate As Date, endDate As Date) As Integer
'    Dim startWeekDay As Integer
'    Dim endWeekDay As Integer
'    Dim daysDifference As Integer
'    Dim weeks As Integer
'    Dim remainingPart As Integer
'    '------------------------------------------------------------------------------------------------------
'
'    startWeekDay = VBA.Weekday(startDate, vbMonday)
'    endWeekDay = VBA.Weekday(endDate, vbMonday)
'
'    daysDifference = endDate - startDate + 1
'    weeks = VBA.Int(daysDifference / 7)
'    remainingPart = daysDifference Mod 7
'
'    If endWeekDay >= startWeekDay Then
'        countWorkingDays = remainingPart - Application.WorksheetFunction.Max(endWeekDay - 5, 0)
'        If countWorkingDays < 0 Then countWorkingDays = 0
'    ElseIf remainingPart Then
'        countWorkingDays = remainingPart - Application.WorksheetFunction.Min(8 - startWeekDay, 2)
'    End If
'
'    countWorkingDays = countWorkingDays + weeks * 5
'
'End Function

Public Function convertIso8601ToDate(dateString As String) As Date
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    Dim hour As Integer
    Dim minute As Integer
    Dim second As Integer
    '------------------------------------------------------------------------------------------------------

    year = VBA.left$(dateString, 4) * 1
    month = VBA.mid$(dateString, 6, 2) * 1
    day = VBA.mid$(dateString, 9, 2) * 1
    hour = VBA.mid$(dateString, 12, 2) * 1
    minute = VBA.mid$(dateString, 15, 2) * 1
    second = VBA.mid$(dateString, 18, 2) * 1

    convertIso8601ToDate = DateSerial(year, month, day) + TimeSerial(hour, minute, second)

End Function



Public Function getPreviousWorkingDay(baseDate As Date) As Date
    getPreviousWorkingDay = VBA.DateAdd("d", baseDate, -1 * f.Numerics.getMax(1, VBA.Weekday(baseDate, vbTuesday) - 4))
End Function

Public Function toMonthYearString(value As Variant) As String
    Dim arr As Variant
    Dim month As Long
    Dim year As Long
    '----------------------------------------------------------------------------------------------------------
    
    If f.Variables.isUndefined(value) Then
        ' skip
    ElseIf Not f.Variables.isPrimitive(value) Then
        ' skip
    ElseIf f.Dates.isDateString(VBA.CStr(value)) Then
        month = VBA.month(value)
        year = VBA.year(value)
    Else
        arr = f.regex.getMatchArray(VBA.CStr(value), MONTH_YEAR_REGEX_PATTERN)
        If f.arrays.isDefined(arr) Then
            month = arr(LBound(arr))
            year = arr(LBound(arr) + 1)
        End If
    End If
    
    If Not f.Math.isInRange(month, 1, 12) Then
    ElseIf Not f.Math.isInRange(year, 2000, 2100) Then
    Else
        toMonthYearString = f.strings.addLeadingZeros(month, 2) & "/" & year
    End If
    
End Function

Public Function getDateFromMonthYearString(value As Variant) As Date
    Dim arr As Variant
    Dim month As Long
    Dim year As Long
    '----------------------------------------------------------------------------------------------------------
    
    If Not f.Variables.isUndefined(value) Then
        arr = f.regex.getMatchArray(VBA.CStr(value), MONTH_YEAR_REGEX_PATTERN)
        If f.arrays.isDefined(arr) Then
            month = arr(LBound(arr))
            year = arr(LBound(arr) + 1)
        End If
        
        If Not f.Math.isInRange(month, 1, 12) Then
        ElseIf year <= 0 Then
        Else
            getDateFromMonthYearString = f.strings.addLeadingZeros(month, 2) & "/" & year
        End If
    End If
    
End Function




Public Function toString(Optional ByVal value As Date, Optional ByVal dateVisibility As DateVisibilityEnum, _
                                                       Optional ByVal format As String) As String
    Dim DEFAULT_FORMATS As Variant:             DEFAULT_FORMATS = VBA.Array(DEFAULT_DATE_ONLY_FORMAT, _
                                                                            DEFAULT_DATE_TIME_FORMAT, _
                                                                            DEFAULT_TIME_ONLY_FORMAT)
    '------------------------------------------------------------------------------------------------------
    
    '[Validate input parameters]
    If Not f.Math.isInRange(dateVisibility, LBound(DEFAULT_FORMATS), UBound(DEFAULT_FORMATS)) Then
        dateVisibility = DateVisibility_DateOnly
    End If
    If value = 0 Then value = VBA.Now

    If VBA.Len(format) = 0 Then format = DEFAULT_FORMATS(dateVisibility)
    toString = VBA.format(value, format)
    
End Function

