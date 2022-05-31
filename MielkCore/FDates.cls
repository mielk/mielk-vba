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



'Public Function isDateString(ByVal value As String, Optional allowTimePart As Boolean = True) As Boolean
'    Const TIMEPART_DISALLOWED_REGEX_PATTERN As String = "^(((((1[26]|2[048])00)|[12]\d([2468][048]|[13579][26]|0[48]))-" & _
'            "((((0[13578]|1[02])-(0[1-9]|[12]\d|3[01]))|((0[469]|11)-(0[1-9]|[12]\d|30)))|" & _
'            "(02-(0[1-9]|[12]\d))))|((([12]\d([02468][1235679]|[13579][01345789]))|((1[1345789]|2[1235679])00))" & _
'            "-((((0[13578]|1[02])-(0[1-9]|[12]\d|3[01]))|((0[469]|11)-(0[1-9]|[12]\d|30)))|(02-(0[1-9]|1\d|2[0-8])))))$"
'    Const TIMEPART_ALLOWED_REGEX_PATTERN As String = "^(((((1[26]|2[048])00)|[12]\d([2468][048]|[13579][26]|0[48]))-" & _
'            "((((0[13578]|1[02])-(0[1-9]|[12]\d|3[01]))|((0[469]|11)-(0[1-9]|[12]\d|30)))|(02-(0[1-9]|[12]\d))))|" & _
'            "((([12]\d([02468][1235679]|[13579][01345789]))|((1[1345789]|2[1235679])00))-((((0[13578]|1[02])-(0[1-9]|[12]\d|3[01]))|" & _
'            "((0[469]|11)-(0[1-9]|[12]\d|30)))|(02-(0[1-9]|1\d|2[0-8])))))( ([01][0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])?)?$"
'    Const TIMEPART_DISALLOWED_REVERSED_REGEX_PATTERN As String = "^(?:(?:31(\/|-|\.)(?:0?[13578]|1[02]))\1|(?:(?:29|30)(\/|-|\.)" & _
'            "(?:0?[1,3-9]|1[0-2])\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:29(\/|-|\.)0?2\3(?:(?:(?:1[6-9]|[2-9]\d)?" & _
'            "(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\d|2[0-8])(\/|-|\.)" & _
'            "(?:(?:0?[1-9])|(?:1[0-2]))\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$"
'    Const TIMEPART_ALLOWED_REVERSED_REGEX_PATTERN As String = "^(?:(?:31(\/|-|\.)(?:0?[13578]|1[02]))\1|(?:(?:29|30)(\/|-|\.)" & _
'            "(?:0?[1,3-9]|1[0-2])\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:29(\/|-|\.)0?2\3(?:(?:(?:1[6-9]|[2-9]\d)?" & _
'            "(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\d|2[0-8])(\/|-|\.)" & _
'            "(?:(?:0?[1-9])|(?:1[0-2]))\4(?:(?:1[6-9]|[2-9]\d)?\d{2})( ([01][0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])?)?$"
'
'    Static objRegexForTimepartDisallowed As Object
'    Static objRegexForTimepartAllowed As Object
'    Static objRevRegexForTimepartDisallowed As Object
'    Static objRevRegexForTimepartAllowed As Object
'    '------------------------------------------------------------------------------------------------------
'    Dim objRegex As Object
'    Dim objRevRegex As Object
'    '------------------------------------------------------------------------------------------------------
'
'    If allowTimePart Then
'        If objRegexForTimepartAllowed Is Nothing Then Set objRegexForTimepartAllowed = F.regex.create(TIMEPART_ALLOWED_REGEX_PATTERN)
'        If objRevRegexForTimepartAllowed Is Nothing Then Set objRevRegexForTimepartAllowed = F.regex.create(TIMEPART_ALLOWED_REVERSED_REGEX_PATTERN)
'        Set objRegex = objRegexForTimepartAllowed
'        Set objRevRegex = objRevRegexForTimepartAllowed
'    Else
'        If objRegexForTimepartDisallowed Is Nothing Then Set objRegexForTimepartDisallowed = F.regex.create(TIMEPART_DISALLOWED_REGEX_PATTERN)
'        If objRevRegexForTimepartDisallowed Is Nothing Then Set objRevRegexForTimepartDisallowed = F.regex.create(TIMEPART_DISALLOWED_REVERSED_REGEX_PATTERN)
'        Set objRegex = objRegexForTimepartDisallowed
'        Set objRevRegex = objRevRegexForTimepartDisallowed
'    End If
'
'    If objRegex.Test(value) Then
'        isDateString = True
'    ElseIf objRegex.Test(VBA.Replace(F.Strings.trimFull(F.Strings.stringify(value)), "/", "-")) Then
'        isDateString = True
'    ElseIf objRegex.Test(VBA.Replace(F.Strings.trimFull(F.Strings.stringify(value)), ".", "-")) Then
'        isDateString = True
'    ElseIf objRevRegex.Test(value) Then
'        isDateString = True
'    ElseIf objRevRegex.Test(F.Strings.trimFull(value)) Then
'        isDateString = True
'    End If
'
'End Function


Public Function isDateString(ByVal value As String, Optional allowTimePart As Boolean = True, Optional ByRef d As Date) As Boolean
    Const TIMEPART_DISALLOWED_REGEX_PATTERN As String = "^\s*\d{4}\s*(-|\.|\/|\\)\s*\d{2}\s*(-|\.|\/|\\)\s*\d{2}\s*$"
    Const TIMEPART_ALLOWED_REGEX_PATTERN As String = "^\s*\d{4}\s*(-|\.|\/|\\|_)\s*\d{2}\s*(-|\.|\/|\\|_)\s*\d{2}\s*(|\d{2}\s*(-|\.|\/|\\|_|:)\s*\d{2}\s*(|(-|\.|\/|\\|_|:)\s*\d{2}\s*))$"
    Const TIMEPART_DISALLOWED_REVERSED_REGEX_PATTERN As String = "^\s*\d{2}\s*(-|\.|\/|\\)\s*\d{2}\s*(-|\.|\/|\\)\s*\d{4}\s*$"
    Const TIMEPART_ALLOWED_REVERSED_REGEX_PATTERN As String = "^\s*\d{2}\s*(-|\.|\/|\\|_)\s*\d{2}\s*(-|\.|\/|\\|_)\s*\d{4}\s*(|\d{2}\s*(-|\.|\/|\\|_|:)\s*\d{2}\s*(|(-|\.|\/|\\|_|:)\s*\d{2}\s*))$"
    '------------------------------------------------------------------------------------------------------
    Static objRegexForTimepartDisallowed As Object
    Static objRegexForTimepartAllowed As Object
    Static objRevRegexForTimepartDisallowed As Object
    Static objRevRegexForTimepartAllowed As Object
    '------------------------------------------------------------------------------------------------------
    Dim objRegex As Object
    Dim objRevRegex As Object
    Dim yearFirst As Long
    Dim hasTime As Boolean
    Dim rejected As Boolean
    Dim strDate As String
    '------------------------------------------------------------------------------------------------------
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    Dim hour As Integer
    Dim minute As Integer
    Dim second As Integer
    '------------------------------------------------------------------------------------------------------

    If allowTimePart Then
        If objRegexForTimepartAllowed Is Nothing Then Set objRegexForTimepartAllowed = F.regex.Create(TIMEPART_ALLOWED_REGEX_PATTERN)
        If objRevRegexForTimepartAllowed Is Nothing Then Set objRevRegexForTimepartAllowed = F.regex.Create(TIMEPART_ALLOWED_REVERSED_REGEX_PATTERN)
        Set objRegex = objRegexForTimepartAllowed
        Set objRevRegex = objRevRegexForTimepartAllowed
    Else
        If objRegexForTimepartDisallowed Is Nothing Then Set objRegexForTimepartDisallowed = F.regex.Create(TIMEPART_DISALLOWED_REGEX_PATTERN)
        If objRevRegexForTimepartDisallowed Is Nothing Then Set objRevRegexForTimepartDisallowed = F.regex.Create(TIMEPART_DISALLOWED_REVERSED_REGEX_PATTERN)
        Set objRegex = objRegexForTimepartDisallowed
        Set objRevRegex = objRevRegexForTimepartDisallowed
    End If

    If objRegex.test(value) Then
        yearFirst = True
        'isDateString = True
    ElseIf objRevRegex.test(value) Then
        yearFirst = False
    Else
        rejected = True
    End If

    If Not rejected Then
        strDate = F.strings.onlyDigits(value, False)
        If yearFirst Then
            year = VBA.left$(strDate, 4)
            month = VBA.mid$(strDate, 5, 2)
            day = VBA.mid$(strDate, 7, 2)
        Else
            day = VBA.left$(strDate, 2)
            month = VBA.mid$(strDate, 3, 2)
            year = VBA.mid$(strDate, 5, 4)
        End If

        If VBA.Len(strDate) > 8 Then
            hour = VBA.mid$(strDate, 9, 2)
            minute = VBA.mid$(strDate, 11, 2)
            If VBA.Len(strDate) > 12 Then
                second = VBA.mid$(strDate, 13, 2)
            End If
        End If

        d = VBA.DateSerial(year, month, day) + VBA.TimeSerial(hour, minute, second)
        If year <> VBA.year(d) Then
            isDateString = False
        ElseIf month <> VBA.month(d) Then
            isDateString = False
        ElseIf day <> VBA.day(d) Then
            isDateString = False
        ElseIf hour <> VBA.hour(d) Then
            isDateString = False
        ElseIf minute <> VBA.minute(d) Then
            isDateString = False
        ElseIf second <> VBA.second(d) Then
            isDateString = False
        Else
            isDateString = True
        End If

    End If

End Function


Public Function isTimeString(ByVal value As String) As Boolean
    Const REGEX_PATTERN As String = "^([01][0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])?$"
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------

    If objRegex Is Nothing Then
        Set objRegex = F.regex.Create(REGEX_PATTERN)
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

Public Function fromString(ByVal value As String) As Date
    Const ILLEGAL_STRING_ERR_DESC As String = "Date/Time string was expected"
    '------------------------------------------------------------------------------------------------------

    If isDDMMYYYDate(value) Then
        fromString = fromDDMMYYYYString(value)
    ElseIf VBA.isDate(value) Then
        fromString = VBA.CDate(value)
    ElseIf isDateString(value) Then
        On Error Resume Next
        fromString = VBA.CDate(value)
        If fromString > 0 Then
        Else
            fromString = VBA.CDate(VBA.Replace(value, ".", "-"))
        End If
        On Error GoTo 0
    ElseIf isTimeString(value) Then
        fromString = VBA.CDate(value)
    Else
        Call Exceptions.IllegalValue.raiseWithCustomMessage(ILLEGAL_STRING_ERR_DESC)
    End If

End Function

Private Function isDDMMYYYDate(value As String) As Boolean
    Const REGEX_PATTERN As String = "^\d\d[-|.|\/]\d\d[-|.|\/]\d\d\d\d$"
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------

    If objRegex Is Nothing Then
        Set objRegex = F.regex.Create(REGEX_PATTERN)
    End If
    isDDMMYYYDate = objRegex.test(value)

End Function

Public Function fromYYYYMMDDString(value As String) As Date
    Dim days As Long
    Dim months As Long
    Dim years As Long
    '------------------------------------------------------------------------------------------------------
    years = VBA.left$(value, 4)
    months = VBA.mid$(value, 5, 2)
    days = VBA.right$(value, 2)
    fromYYYYMMDDString = VBA.DateSerial(years, months, days)
End Function

Private Function fromDDMMYYYYString(ByVal value As String) As Date
    Dim days As Long
    Dim months As Long
    Dim years As Long
    '------------------------------------------------------------------------------------------------------
    value = F.strings.onlyDigits(value)
    days = VBA.left$(value, 2)
    months = VBA.mid$(value, 3, 2)
    years = VBA.right$(value, 4)
    fromDDMMYYYYString = VBA.DateSerial(years, months, days)
End Function

Public Function fromMMDDYYYYString(text As String) As Date
    Const REGEX_PATTERN As String = "(\d{1,2})\/(\d{1,2})\/(\d{2,4}) "
    '------------------------------------------------------------------------------------------------------
    Dim parts As Variant
    '------------------------------------------------------------------------------------------------------
    
    parts = F.regex.getMatchArray(text, REGEX_PATTERN)
    If F.arrays.isDefined(parts) Then
        fromMMDDYYYYString = VBA.DateSerial(VBA.IIf(parts(3) < 100, parts(3) + 2000, parts(3)), parts(1), parts(2))
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
    getTimestampString = VBA.format(VBA.Now, "YYYYMMDDHHMMSS" & F.strings.addLeadingZeros(hundrethsSecond, 3))
End Function

Public Function toMMSSFormat(value As Double) As String
    Dim minutes As Integer
    Dim seconds As Integer
    '------------------------------------------------------------------------------------------------------
    minutes = value \ 60
    seconds = value Mod 60
    toMMSSFormat = F.strings.addLeadingZeros(minutes) & ":" & F.strings.addLeadingZeros(seconds)
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
    If Not F.Variables.isUndefined(dt) Then
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
        Set regexes(1) = F.regex.Create("^[ ]*(Jan(uary|)|sty(cze.|))[ ]*$")
        Set regexes(2) = F.regex.Create("^[ ]*(Feb(ruary|)|lut(y|))[ ]*$")
        Set regexes(3) = F.regex.Create("^[ ]*Mar(ch|zec|)[ ]*$")
        Set regexes(4) = F.regex.Create("^[ ]*(Apr(il|)|kwi(ecie.|))[ ]*$")
        Set regexes(5) = F.regex.Create("^[ ]*Ma(y|j)[ ]*$")
        Set regexes(6) = F.regex.Create("^[ ]*(Jun(e|)|cze(rwiec|))[ ]*$")
        Set regexes(7) = F.regex.Create("^[ ]*(Jul(y|)|lip(iec|))[ ]*$")
        Set regexes(8) = F.regex.Create("^[ ]*(Aug(ust|)|sie(rpie.|))[ ]*$")
        Set regexes(9) = F.regex.Create("^[ ]*(Sep(tember|)|wrz(esie.|))[ ]*$")
        Set regexes(10) = F.regex.Create("^[ ]*(Oct(ober|)|pa.(dziernik|))[ ]*$")
        Set regexes(11) = F.regex.Create("^[ ]*(Nov(ember|)|lis(topad|))[ ]*$")
        Set regexes(12) = F.regex.Create("^[ ]*(Dec(ember|)|gru(dzie.|))[ ]*$")
    End If

    For i = LBound(regexes) To UBound(regexes)
        If regexes(i).test(name) Then
            getMonthNumberFromName = i
            Exit For
        End If
    Next i

End Function

Public Function getNextMonthStart(d As Date) As Date
    Dim tempDate As Date
    '------------------------------------------------------------------------------------------------------
    tempDate = VBA.DateAdd("m", 1, d)
    getNextMonthStart = VBA.DateSerial(VBA.year(tempDate), VBA.month(tempDate), 1)
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
    getPreviousWorkingDay = VBA.DateAdd("d", baseDate, -1 * F.Numerics.getMax(1, VBA.Weekday(baseDate, vbTuesday) - 4))
End Function

