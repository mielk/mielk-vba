VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FFunctions"
'--------------------------------------------------------------------------------------------------------------

Public Function getDaysRange(Optional startDate As Variant, Optional endDate As Variant, _
                                                                Optional maxRange As Long = 0) As EDatesRange
    Dim daysRange As EDatesRange
    Dim validation As ValidationObject
    '----------------------------------------------------------------------------------------------------------
    
    Set daysRange = MApp.Create.DatesRange(False, maxRange)
    With daysRange
        If F.Dates.isDateString(VBA.CStr(startDate)) Then Call .setStartDate(F.Dates.fromString(VBA.CStr(startDate)))
        If F.Dates.isDateString(VBA.CStr(endDate)) Then Call .setEndDate(F.Dates.fromString(VBA.CStr(endDate)))
        Set validation = .edit
    End With
    
    If validation.getStatus Then Set getDaysRange = daysRange
    
End Function


Public Function getMonthsRange(Optional ByVal startDate As Variant, Optional ByVal endDate As Variant, _
                                                                Optional maxRange As Long = 0) As EDatesRange
    Dim monthsRange As EDatesRange
    Dim validation As ValidationObject
    '----------------------------------------------------------------------------------------------------------
    
    'Try to convert given dates
    If Not VBA.IsDate(startDate) Then startDate = F.Dates.getDateFromMonthYearString(startDate)
    If Not VBA.IsDate(endDate) Then endDate = F.Dates.getDateFromMonthYearString(endDate)
    
    Set monthsRange = MApp.Create.DatesRange(True, maxRange)
    With monthsRange
        Call .setStartDate(VBA.CDate(startDate))
        Call .setEndDate(VBA.CDate(endDate))
        Set validation = .edit
    End With
    
    If validation.getStatus Then Set getMonthsRange = monthsRange
    
End Function


