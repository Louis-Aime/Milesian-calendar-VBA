Attribute VB_Name = "MilesianMoon"
'MilesianMoon: Find a moon near a date in Julian day.
'Copyright Miletus SARL 2017. www.calendriermilesien.org
'For use as an MS Excel VBA module. Tested under MS Excel 2016.
'No warranty.
'If transmitted, even with changes, present header shall be maintained in full.
'This package uses no other element.
'Note: the moon computed here is a mean moon, with no Terrestrial Time correction.
'The error on mean moon is less than one day for three milliniums before and after year 2000.
'Functions:
'   LastMoonPhase : last time the mean moon reached a given phase
'   NewtMoonPhase : next time where the mean moon reaches a given phase
'Parameters
'   FromDate: a date and time UTC expressed in Excel
'   MoonPhase: 0 or omitted: new moon; 1: first quarter; 2: full moon; 3: last quarter; <0 or >3: error.
'Version V1.0 M2017-06-15 -
Const MeanSynodicMoon As Double = 29.53058883 'Mean duration of a synodic month in decimal days = 29d 12h 44mn 2s 7/8s
Const MeanNewMoon2000 As Double = 36531.59773 'G2000-01-06T14-20-44 TT, conventional date of a mean new moon
Const DayOffsetMacOS As Long = 1462 'Offset for Date1904
Function LastMoonPhase(ByVal FromDate As Date, Optional ByVal MoonPhase As Integer) As Date
'Date of last mean new moon, if present date is FromDate (in UTC).
'MoonPhase: 0 or omitted, new moon; 1: first quarter; 2: full moon; 3: last quarter; else: error.
Dim Phase As Double
If IsMissing(MoonPhase) Then MoonPhase = 0
If ActiveWorkbook.Date1904 And FromDate >= 1 Then FromDate = FromDate - DayOffsetMacOS
If MoonPhase < 0 Or MoonPhase > 3 Then Error 1
Phase = FromDate - MeanNewMoon2000 - (MoonPhase / 4) * MeanSynodicMoon
If ActiveWorkbook.Date1904 Then Phase = Phase + DayOffsetMacOS
 While Phase < 0
    Phase = Phase + MeanSynodicMoon
 Wend
 While Phase >= MeanSynodicMoon
    Phase = Phase - MeanSynodicMoon
 Wend
LastMoonPhase = FromDate - Phase
End Function
Function NextMoonPhase(ByVal FromDate As Double, Optional ByVal MoonPhase As Integer) As Date
'Fractional days until next new moon, if present date is FromDate (in UTC).
'MoonPhase is as in LastMoonPhase
Dim Phase As Double
If IsMissing(MoonPhase) Then MoonPhase = 0
If ActiveWorkbook.Date1904 And FromDate >= 1 Then FromDate = FromDate - DayOffsetMacOS
If MoonPhase < 0 Or MoonPhase > 3 Then Error 1
Phase = MeanNewMoon2000 - FromDate + (MoonPhase / 4) * MeanSynodicMoon
If ActiveWorkbook.Date1904 Then Phase = Phase - DayOffsetMacOS
 While Phase < 0
    Phase = Phase + MeanSynodicMoon
 Wend
 While Phase >= MeanSynodicMoon
    Phase = Phase - MeanSynodicMoon
 Wend
NextMoonPhase = FromDate + Phase
End Function
