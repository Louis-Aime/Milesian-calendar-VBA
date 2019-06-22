Attribute VB_Name = "MilesianCalendar"
'MilesianCalendar: Enter and display dates in Microsoft Excel following Milesian calendar conventions
'Copyright Miletus SARL 2016-2019. www.calendriermilesien.org
'For use as an MS Excel VBA module.
'Developped under Excel 2016
'Tested under MS Excel 2007 (Windows) and 2016 (Windows and MacOS)
'No warranty.
'May be used for personal or professional purposes.
'If transmitted or integrated, even with changes, present header shall be maintained in full.
'Functions are aimed at extending Date & Time functions to the Milesian calendar.
'Whenever possible, similar parameter lists and syntax as for standard Date-Time functions are used.
'Public function include:
'   MILESIAN_DATE(Year, Month, Day): the Date object correponding to the Milesian date of parameters.
'   MILESIAN_YEAR, MILESIAN_MONTH, MILESIAN_DAY, MILESIAN_TIME: same as standard YEAR, MONTH, DAY, TIME with Milesian components.
'   MILESIAN_DISPLAY: display a date in Milesian. Set second parameter (Wtime) to False to discard time part.
'   DURATION: Duration between two dates, in decimal days.
'   DATE_SHIFT: Date shifted by a duration.
'   MILESIAN_MONTH_SHIFT, MILESIAN_MONTH_END: same as standard, with Milesian months
'   MILESIAN_IS_LONG_YEAR(Year): Boolean, true if the Milesian Year is long i.e. just before a Gregorian leap year.
'   MILESIAN_YEAR_BASE(Year): the day just before 1 1m Year, used for Doomsday and Milesian epact.
'       MILESIAN_EPACT(Year): the Milesian epact of the year, rounded to nearest half-integer value
'       MILESIAN_DOOMSDAY(Year): the Doomsday (clavedi), key day of week for this Milesian and Gregorian year
'   JULIAN_EPOCH_COUNT: Decimal Julian Day from a date expression
'   JULIAN_EPOCH_DATE: Date from a Julian Day
'Two special moon functions
'   MOON_PHASE_LAST : last time the mean moon reached a given phase
'   MOON_PHASE_NEXT : next time where the mean moon reaches a given phase
'Parameters for these two
'   FromDate: a date and time UTC expressed as a date expression
'   MoonPhase: 0 or omitted: new moon; 1: first quarter; 2: full moon; 3: last quarter; <0 or >3: error.
'Special functions for Excel:
'   DAYOFWEEK_Ext: Works like standard DAYOFWEEK but is extended:
'       Second paramater is 0 by default, and yields: 0:Sunday to 6: Saturday instead of 1 to 7.
'       Other value of the second parameter give the same result as DAYOFWEEK.
'   EASTER_SUNDAY(Year): Date of Easter in the Gregorian calendar, for any year > 1582.
'Version V2 M2018-02-10
    ' A special function, MICROSOFT_DATE_TIME_FIX, is used for converting Date objects from textual date-time expression before 30 Dec 1899.
'Version V3 M2019-06-13
    'Solar intercalation is Gregorian (no 3200 years cycle)
    'Excel Date object is considered with its special properties for values before 30/12/1899.
    'MICROSOFT_DATE_TIME_FIX suppressed
    'Change management of Date1904: Object Date results shall always be in MacOS count, and any result < 0 raises an error.
'Version V4 M2019-06-28
    'Change control of Date parameters in such a way that they are always properly converted, even with Date1904.
    'Utilities functions are not visible anymore.
    'In all cases, including Date1904, translation to Date object is made effectively - MS dates 1900-01-01 to 1900-02-29 are considered false.
    'Add MILESIAN_EPACT and MILESIAN_DOOMSDAY
    'Change YEAR_BASE to the day before 1 1m at 7:30
    'Move moon routines into this module.

Const MStoPresentEra As Long = 693969 'Offset between 1/1m/0 and Microsoft origin (1899-12-30T00:00:00 is 0)
Const MStoJulianMinus05 As Long = 2415018 'Offset between julian day epoch and Microsoft origin, minus 0.5
Const HighYear = 10000   'Highest year: Excel goes up to 31/12/9999 Gregorian, i.e. 10000 Milesian
Const LowYear = 100     'As long as MS does not handle date before Gregorian 1/1/100. Else we'd take -9999
Public Const DayOffsetMacOS As Long = 1462 'Used also in DateParse
'Moon parameters
Const MeanSynodicMoon As Double = 29.53058883 'Mean duration of a synodic month in decimal days = 29d 12h 44mn 2s 7/8s
Const MeanNewMoon2000 As Double = 36531.59773 'G2000-01-06T14-20-44 TT, conventional date of a mean new moon

'#Part 1: internal procedures
Sub Milesian_IntegDiv(ByVal Dividend As Variant, ByVal Divisor As Long, Cycle As Long, Phase As Variant)
'Quotient and modulo in the same operation. Divisor shall by positive.
'Cycle (i.e. Quotient) is same sign as Dividend. 0 <= Phase (i.e. modulo) < Divisor.
Cycle = 0
Phase = Dividend
If Divisor > 0 Then
  While Phase < 0
    Phase = Phase + Divisor
    Cycle = Cycle - 1
    Wend
  While Phase >= Divisor
    Phase = Phase - Divisor
    Cycle = Cycle + 1
    Wend
Else
  Err.Raise 1
End If
End Sub

Sub Milesian_IntegDivCeiling(ByVal Dividend As Variant, ByVal Divisor As Long, ByVal ceiling As Integer, Cycle As Long, Phase As Variant)
'Quotient and modulo in the same operation. By exception, remainder may be = divisor if quotient = ceiling
'Cycle (i.e. Quotient) is same sign as Dividend. 0 <= Phase (i.e. modulo) <= Divisor.
Cycle = 0
Phase = Dividend
If Divisor > 0 And Dividend >= 0 And Dividend <= ceiling * Divisor + 1 Then
  ceiling = ceiling - 1 'Decrease ceiling by 1 in order to simplify test in the next loop
  While (Phase >= Divisor) And Cycle < ceiling
    Phase = Phase - Divisor
    Cycle = Cycle + 1
    Wend
Else
  Err.Raise 1
End If
End Sub

Private Function PosDiv(ByVal A, D)  'The Integer division with positive remainder
PosDiv = 0
  If D <= 0 Then
    Err.Raise 1
  Else
    While (A < 0)
        A = A + D
        PosDiv = PosDiv - 1
    Wend
    While (A >= D)
        A = A - D
        PosDiv = PosDiv + 1
    Wend
  End If
End Function

Private Function PosMod(ByVal A As Long, D As Integer) As Integer  'The always positive modulo, even if A<0
    If D <= 0 Then
        Err.Raise 1
    Else
        While (A < 0)
            A = A + D
        Wend
        While (A >= D)
            A = A - D
        Wend
    PosMod = A
    End If
End Function

'#Part 1.1 : MS Excel specific internal functions

'Two issues under MS Excel

'1. Date1904 sheets: in principle, a date is stored as the number of days since 1/1/1904
' but any call to a VBA Date object converts it into a number starting à 30/12/1899,
' except for 1/1/1904 itself (+/- 1 day).

'2. Before Day 0 (1899-12-30), the time of the day part of a time stamp is represented backwards:
'Let D + T be the classical representation of a timestamp TS
'where D = Int(TS) and 0<=T<1 (Int(X) is the highest integer lower than or equal to X)
'Then if TS < 0, then the Microsoft timestamp MTS is
' MTS = D - T = Int(TS) - (TS - Int(TS)) = 2*Int(TS)-TS
' In the other way:if MTS < 0
' TS = - 2*Int(-MTS)-MTS
'Note that -1 < MTS < 0 is not possible

Sub Any_to_Date(TheNum As Variant, MyDate As Date)
' Internal procedure (not visible function) that converts any expression TheNum into the VBA Date object MyDate
MyDate = TheNum     'Force conversion to Date, which is enough for most situations
If ActiveWorkbook.Date1904 Then
    If Not IsDate(TheNum) Then MyDate = TheNum + DayOffsetMacOS
    If IsDate(TheNum) And TheNum > -1 And TheNum < 1 Then MyDate = TheNum + DayOffsetMacOS 'Special ill-handled date
End If
End Sub

Sub Any_to_Uniform(ThePossibleDate As Variant, TheNum As Double)
' ThePossibleDate: anything expressing a date in the context of the sheet.
' TheNum is the Date number as a "uniform" numeric object, i.e. time part is always the decimal positive part of the number representing a date
Dim TheDate As Date
' 1. Convert into the right date object
TheDate = ThePossibleDate    'Force conversion into Date object
If ActiveWorkbook.Date1904 Then
    If Not IsDate(ThePossibleDate) Then TheDate = ThePossibleDate + DayOffsetMacOS
    If IsDate(ThePossibleDate) And ThePossibleDate > -1 And ThePossibleDate < 1 Then TheDate = ThePossibleDate + DayOffsetMacOS 'Special 1/1/1904
End If
' 2. Convert into a uniform double expression
If TheDate >= 0 Then
    TheNum = TheDate
Else
    If TheDate > -1 Then Err.Raise 1, , "Invalid counter" 'No value in ]-1, 0[
    TheNum = -2 * Int(-TheDate) - TheDate
End If
End Sub

Private Function Standard_to_MS_Date(SDate As Double) As Date
'Back to a date expression after a computation on a uniform date counter
Dim Mac As Double
'2019:  Setting MacOS offset is necessary only if calling function is used in a cell's formula or alone.
'       If calling function is called within a VBA routine, add immediately DayOffsetMacOS to result.
Mac = 0
If ActiveWorkbook.Date1904 Then Mac = -DayOffsetMacOS
If SDate >= 0 Then
    Standard_to_MS_Date = SDate + Mac
Else
    Standard_to_MS_Date = 2 * Int(SDate) - SDate + Mac
End If
End Function

'#Part 2: Compute date from milesian parameters
'Date computed here is integer by construction.
'Result is converted to Date1904 if necessary

Function MILESIAN_DATE(Year, Month, DayInMonth) As Date
Attribute MILESIAN_DATE.VB_Description = "Date, from Year (100-9999), Month (1-12), Day (1-31) of Milesian calendar."
'The Date type element from a Milesian date given as year (common era, relative), month, daynumber in month.
Dim A As Integer 'Intermediate computations as non-long integer
Dim B As Long   'Bimester number, for intermediate computations
Dim M1 As Long  'Month rank
Dim D As Long   'Days expressed in long integer
Dim Y As Integer 'Internal Year
'Check that Milesian date is OK
If Year <> Int(Year) Or Month <> Int(Month) Or DayInMonth <> Int(DayInMonth) Then Err.Raise 1, , "Invalid date"
If Year >= LowYear And Year <= HighYear And Month > 0 And Month < 13 And DayInMonth > 0 And DayInMonth < 32 Then 'Basic filter
  M1 = Month - 1 'Count month rank, i.e. 0..11
  Milesian_IntegDiv M1, 2, B, M1 'B = full bimesters, M1 = 1 if a full month added, else 0
  If DayInMonth < 31 Or (M1 = 1 And (B < 5 Or MILESIAN_IS_LONG_YEAR(Year))) Then
    D = Year         'Initiate internal year and force long-integer conversion
    A = PosDiv(D, 4) - PosDiv(D, 100) + PosDiv(D, 400) 'Sum non-long terms: leap days
    D = D * 365      'Begin computation of days in long-integer;
    D = D - MStoPresentEra - 1 + B * 61 + M1 * 30 + A + DayInMonth 'Computations in long-integer first
    If ActiveWorkbook.Date1904 Then
        D = D - DayOffsetMacOS
        End If
    MILESIAN_DATE = D
  Else
    Err.Raise 1, , "Invalid date"
  End If
Else
  Err.Raise 1, , "Invalid date"
End If
End Function

'#Part 3: Extract Milesian date elements from a Date-type number or string
'The time-of-day element of a negative Date objects is properly converted in positive

Sub Milesian_DateElement(AnyDate As Variant, Y As Integer, M As Integer, Q As Integer, T As Variant)
' From AnyDate, a token that holds a date in any representation (not necessarily a Date object),
' compute the milesian date element Q / M / Y (day in month, month, year)
' and also the decimal part T, set to a positive number, which is the UTC time in the day.
' Y is year in common era, relative value (may be 0 or negative)
' M is milesian month number, 1 to 12
' Q is number of day in month, 1 to 31
' This is an internal subroutine. Corresponding functions come after.
' Note2: On Excel 2016, Gregorian year of Date element is greater of equal to 100.
Dim Cycle As Long, Day As Long      'Cycle is used serveral times with a different meaning each time
Dim Dnum As Double 'The date in standard i.e. continuous representation, even for negative values.
Any_to_Uniform AnyDate, Dnum 'Convert to a uniform number
Day = Int(Dnum)  'Initiate Day as highest integer lower or equal to DNum, and avoid control on Date type
T = Dnum - Day   'Extract time part of the Date element. Always positive.
Day = Day + MStoPresentEra
Milesian_IntegDiv Day, 146097, Cycle, Day    'Day is day rank in 400 years period, Cycle is quadrisaeculum
Y = Y + Cycle * 400
Milesian_IntegDivCeiling Day, 36524, 4, Cycle, Day   'Day is day rank in century, Cycle is rank of century
Y = Y + Cycle * 100
Milesian_IntegDiv Day, 1461, Cycle, Day              'Day rank in quadriannum
Y = Y + Cycle * 4
Milesian_IntegDivCeiling Day, 365, 4, Cycle, Day     'Day rank in year
Y = Y + Cycle
Milesian_IntegDiv Day, 61, Cycle, Day             'Day rank in bimester
M = 2 * Cycle
Milesian_IntegDivCeiling Day, 30, 2, Cycle, Day  'Day: day rank within month, Cycle = month rank in bimester
M = M + Cycle + 1                       'M: month number, 1 to 12
Q = Day + 1                             'Q: day number within month, 1 to 31
End Sub

Function MILESIAN_YEAR(TheDate As Variant) As Integer  'The milesian year (common era) for a Date argument.
Attribute MILESIAN_YEAR.VB_Description = "Milesian year from a date expression."
Dim Y As Integer, M As Integer, Q As Integer, T As Variant
Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
MILESIAN_YEAR = Y
End Function

Function MILESIAN_MONTH(TheDate As Variant) As Integer  'The milesian month number (1-12) for a Date argument.
Attribute MILESIAN_MONTH.VB_Description = "Milesian month number (1 to 12), from a date expression."
Dim Y As Integer, M As Integer, Q As Integer, T As Variant
Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
MILESIAN_MONTH = M
End Function

Function MILESIAN_DAY(TheDate As Variant) As Integer  'The day number in the milesian month for a Date argument.
Attribute MILESIAN_DAY.VB_Description = "Day in Milesian month, from a date expression."
Dim Y As Integer, M As Integer, Q As Integer, T As Variant
Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
MILESIAN_DAY = Q
End Function

Function MILESIAN_TIME(TheDate As Variant) As Date
Attribute MILESIAN_TIME.VB_Description = "Time component in a date expression."
'Extract date from a date element, even negative.
Dim Y As Integer, M As Integer, Q As Integer, T As Date 'Force T to a date, i.e. a time since < 1
Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
MILESIAN_TIME = T
End Function

Function MILESIAN_DISPLAY(TheDate As Variant, Optional Wtime As Boolean = True) As String
Attribute MILESIAN_DISPLAY.VB_Description = "Date & time string representing a Milesian date, from a date expression. Set Wtime to 'False' to drop time component."
'Milesian date as a string, from a Date element
Dim Y As Integer, M As Integer, Q As Integer, T As Date 'Force T to a date, i.e. a time since < 1
Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
MILESIAN_DISPLAY = Q & " " & M & "m " & Y
If Wtime Then MILESIAN_DISPLAY = MILESIAN_DISPLAY & " " & T
End Function

'#Part 4: Computations on duration, distant dates, and milesian months

Function DURATION(Begin_date As Variant, End_date As Variant) As Double
Attribute DURATION.VB_Description = "Elapsed time from Begin to End, in decimal days, can be formatted in hh:mm:ss"
'Elapsed time, in decimal days, from one date to the other.
Dim Begin_num As Double, End_num As Double
Any_to_Uniform Begin_date, Begin_num
Any_to_Uniform End_date, End_num
DURATION = End_num - Begin_num
End Function

Function DATE_SHIFT(Origin As Variant, TimeShift As Double) As Date
Attribute DATE_SHIFT.VB_Description = "Elapsed time from Begin to End, in decimal days, can be formatted in hh:mm:ss"
' Add time in decimal days to a date expression, to obtain a date object.
' TimeShift may be negative
Dim Dnum As Double
Any_to_Uniform Origin, Dnum
Dnum = Dnum + TimeShift
DATE_SHIFT = Standard_to_MS_Date(Dnum)
End Function

Function MILESIAN_MONTH_SHIFT(TheDate As Variant, MonthShift As Long) As Date 'Same date several (milesian) months later of earlier
Attribute MILESIAN_MONTH_SHIFT.VB_Description = "Date of same day in MonthShift Milesian months from TheDate."
Dim Y As Integer, M As Integer, Q As Integer, D As Integer
Dim M1 As Long, Cycle As Long, Phase As Long
'Compute begin milesian date
Milesian_DateElement TheDate, Y, M, Q, D
'Compute month rank from 1m of year 0
M1 = Y                     ' Force computation of month in Long
M1 = (M1 * 12) + MonthShift + M - 1 'In this order, Long shall be before simple Integer
'Compute year and month rank
Milesian_IntegDiv M1, 12, Cycle, Phase
If Cycle < LowYear Or Cycle > HighYear Then Err.Raise 1 'Stop if computed year is too low
Y = Cycle
M = Phase + 1
'If Q was 31, set to end of month, else use same day number
If (Q = 31) And (((M Mod 2) = 1) Or ((M = 12) And Not MILESIAN_IS_LONG_YEAR(Y))) Then Q = 30
MILESIAN_MONTH_SHIFT = MILESIAN_DATE(Y, M, Q)
End Function

Function MILESIAN_MONTH_END(TheDate As Variant, MonthShift As Long) As Date 'End of month several (milesian) months later of earlier
Attribute MILESIAN_MONTH_END.VB_Description = "Date of last day of Milesian month in MonthShift months from TheDate."
Dim Y As Integer, M As Integer, Q As Integer, D As Integer
Dim M1 As Long, Cycle As Long, Phase As Long
'Compute begin milesian date
Milesian_DateElement TheDate, Y, M, Q, D
'Compute month rank from 1m of year 0
M1 = Y                     ' Force computation of month in Long
M1 = (M1 * 12) + MonthShift + M - 1 'In this order, Long shall be before simple Integer
'Compute year and month rank
Milesian_IntegDiv M1, 12, Cycle, Phase
If Cycle < LowYear Or Cycle > HighYear Then Err.Raise 1 'Stop if computed year is too low
Y = Cycle
M = Phase + 1
'If Q was 31, set to end of month, else use same day number
If (((M Mod 2) = 1) Or ((M = 12) And Not MILESIAN_IS_LONG_YEAR(Y))) Then
    Q = 30
Else: Q = 31
End If
MILESIAN_MONTH_END = MILESIAN_DATE(Y, M, Q)
End Function

'#Part 5: Year's characteristics

Function MILESIAN_IS_LONG_YEAR(ByVal Year) As Boolean
Attribute MILESIAN_IS_LONG_YEAR.VB_Description = "Is this year a 366-days Milesian year."
'Is year Year a 366 days year, i.e. a year just before a bissextile year following the Milesian rule.
If Year <> Int(Year) Or Year < LowYear Or Year > HighYear Then Err.Raise 1, , "Invalid year"
Year = Year + 1
MILESIAN_IS_LONG_YEAR = PosMod(Year, 4) = 0 And (PosMod(Year, 100) <> 0 Or PosMod(Year, 400) = 0)
End Function

Function MILESIAN_YEAR_BASE(ByVal Year) As Date 'The Year base of a year i.e. the date just before the 1 1m of the year, at 7:30
Attribute MILESIAN_YEAR_BASE.VB_Description = "The day before the Milesian new year's day at 7:30 (UTC), where the Milesian epact is computed"
Dim A As Integer, D As Long, YB As Date
If Year <> Int(Year) Or Year < LowYear Or Year > HighYear Then Err.Raise 1, , "Invalid year"
D = Year        'Force long-integer conversion
D = D * 365     'Begin computation of days in long-integer;
A = PosDiv(Year, 4) - PosDiv(Year, 100) + PosDiv(Year, 400)
D = D - MStoPresentEra + A - 1           'Computations in long-integer first
YB = D  'Force Date conversion
MILESIAN_YEAR_BASE = DATE_SHIFT(YB, 0.3125) 'This day at 7:30 (for moon computations)
'If ActiveWorkbook.Date1904 Then D = D - DayOffsetMacOS
End Function

Function MILESIAN_EPACT(ByVal Year) As Double 'The Milesian Epact computed from the mean moon, a duration
Attribute MILESIAN_EPACT.VB_Description = "The moon age one day before new Milesian year's day"
    Dim Dnum As Double, B1 As Date, B2 As Date
    If Year <> Int(Year) Or Year < LowYear Or Year > HighYear Then Err.Raise 1, , "Invalid year"
    B1 = MILESIAN_YEAR_BASE(Year)
    If ActiveWorkbook.Date1904 Then B1 = B1 + DayOffsetMacOS 'Cancel artificial final offset for intermediate computation
    B2 = MOON_PHASE_LAST(B1)
    If ActiveWorkbook.Date1904 Then B2 = B2 + DayOffsetMacOS 'Cancel artificial final offset for intermediate computation
    Dnum = DURATION(B2, B1)
    MILESIAN_EPACT = Int(2 * Dnum + 0.5) / 2
End Function

Function MILESIAN_DOOMSDAY(ByVal Year, Optional DispType As Integer) As Integer 'The Doomsday for Milesian and also Gregorian year
Attribute MILESIAN_DOOMSDAY.VB_Description = "Doomsday, or key day of week (clavedi) for this Milesian and Gregorian year, rounded to half-integer"
    Dim YBase As Date
    If Year <> Int(Year) Or Year < LowYear Or Year > HighYear Then Err.Raise 1, , "Invalid year"
    YBase = MILESIAN_YEAR_BASE(Year)
    If ActiveWorkbook.Date1904 Then YBase = YBase + DayOffsetMacOS 'Cancel artificial final offset for intermediate computation
    If IsMissing(DispType) Then DispType = 0
    MILESIAN_DOOMSDAY = DAYOFWEEK_Ext(YBase, DispType)
End Function

'#Part 6: Julian Epoch Day conversion functions

Function JULIAN_EPOCH_COUNT(AnyDate As Variant)
Attribute JULIAN_EPOCH_COUNT.VB_Description = "Decimal julian day from date expression."
    Dim IntDate As Long, TimePart   'Compute separately integer part and decimal (time) part
    Dim Dnum As Double
    Any_to_Uniform AnyDate, Dnum 'Convert to Dnum
    IntDate = Int(Dnum)  'Integer part is Date at 00:00
    TimePart = Dnum - IntDate     'Time in day
    TimePart = TimePart + 0.5       '...shifted to Julian Day convention
    IntDate = IntDate + MStoJulianMinus05 'Epoch shift
    JULIAN_EPOCH_COUNT = TimePart + IntDate
End Function

Function JULIAN_EPOCH_DATE(Julian_Count) As Date
Attribute JULIAN_EPOCH_DATE.VB_Description = "Date from decimal julian day."
    Dim IntDate, TimePart As Date
    IntDate = Int(Julian_Count)       'Integer part of Julian Day
    TimePart = Julian_Count - IntDate 'Decimal part, i.e. time after noon
    TimePart = TimePart + 0.5 'Add, not substract, a half day
    IntDate = IntDate - MStoJulianMinus05 - 1 'Compensate full day added from above
    JULIAN_EPOCH_DATE = Standard_to_MS_Date(TimePart + IntDate) 'Convert back to MS notation
End Function

'#Part 7: Moon computations

Function MOON_PHASE_LAST(TheDate As Variant, Optional MoonPhase As Integer) As Date
Attribute MOON_PHASE_LAST.VB_Description = "Date of last mean moon phase before TheDate (UTC). 0 or omitted: new moon, 1: 1st quarter, 2: full moon, 3: last quarter."
'Date of last mean moon phase before FromDate (in UTC).
'MoonPhase: 0 or omitted, new moon; 1: first quarter; 2: full moon; 3: last quarter; else: error.
Dim FromDate As Double
Dim Phase As Double
If IsMissing(MoonPhase) Then MoonPhase = 0
If MoonPhase < 0 Or MoonPhase > 3 Then Err.Raise 1
Any_to_Uniform TheDate, FromDate 'Use FromDate as a uniform representation of the asked date
Phase = FromDate - MeanNewMoon2000 - (MoonPhase / 4) * MeanSynodicMoon
 While Phase < 0
    Phase = Phase + MeanSynodicMoon
 Wend
 While Phase >= MeanSynodicMoon
    Phase = Phase - MeanSynodicMoon
 Wend
MOON_PHASE_LAST = Standard_to_MS_Date(FromDate - Phase)
End Function

Function MOON_PHASE_NEXT(TheDate As Variant, Optional MoonPhase As Integer) As Date
Attribute MOON_PHASE_NEXT.VB_Description = "Date of next mean moon phase since TheDate (UTC). 0 or omitted: new moon, 1: 1st quarter, 2: full moon, 3: last quarter."
'Date of next mean moon phase after FromDate (in UTC).
'MoonPhase: 0 or omitted, new moon; 1: first quarter; 2: full moon; 3: last quarter; else: error.
Dim FromDate As Double
Dim Phase As Double
If IsMissing(MoonPhase) Then MoonPhase = 0
If MoonPhase < 0 Or MoonPhase > 3 Then Err.Raise 1
Any_to_Uniform TheDate, FromDate 'Use FromDate as a representation of the asked date
Phase = MeanNewMoon2000 - FromDate + (MoonPhase / 4) * MeanSynodicMoon
 While Phase < 0
    Phase = Phase + MeanSynodicMoon
 Wend
 While Phase >= MeanSynodicMoon
    Phase = Phase - MeanSynodicMoon
 Wend
MOON_PHASE_NEXT = Standard_to_MS_Date(FromDate + Phase)
End Function

'#Part 8: Excel specific public functions

Function DAYOFWEEK_Ext(AnyDate As Variant, Optional DispType As Integer) As Integer
Attribute DAYOFWEEK_Ext.VB_Description = "Day of week of date expression. Type: 0 (default) 0-6, 0=Sunday; 1 (Excel default) 1-7, 1=Sunday; Others: same as Excel's"
    Dim IntDate As Long, Start As Integer, Phase As Integer, Dnum As Double
    
    '1. Compute Start and Phase from DispType
    If IsMissing(DispType) Then DispType = 0
    'DispType 0 is not used with standard DOW routines.
    'It uses Milesian way and John Conway's convention: Sunday = 0, Monday = 1, up to Saturday = 6
    Phase = 6   'The most common case: cycle starts with Sunday
    Select Case DispType
        Case 0          'The Milesian, John Conway, the most simple to memorize
            Start = 0
        Case 1          'The Spreadsheets' standard
            Start = 1
        Case 2
            Start = 1
            Phase = Phase - 1
        Case 3
            Start = 0
            Phase = Phase - 1
        Case 11 To 17
            Start = 1
            Phase = Phase - (DispType - 10)
        Case Else
            Err.Raise 1
        End Select
    
    '2. Extract Date element and compute
    Any_to_Uniform AnyDate, Dnum
    IntDate = Int(Dnum)  'Convert date-time to hold date component only
    DAYOFWEEK_Ext = PosMod(IntDate + Phase, 7) + Start
End Function

Function EASTER_SUNDAY(ByVal Year) As Date 'Easter Date computed after the Milesian method www.calendriermilesien.org
Attribute EASTER_SUNDAY.VB_Description = "Date of Easter Sunday (Gregorian computus) for the given Year"
    Dim S As Long, B As Long, N As Long, H As Integer, R As Integer 'Components of year, Golden number minus 1, Easter residue
    Dim Dnum As Double 'Possible date
    If Year <> Int(Year) Or Year < 1583 Or Year > HighYear Then Err.Raise 1
    Milesian_IntegDiv Year, 100, S, N   'Decompose Year in centuries (S) + years in century (N)
    Milesian_IntegDiv N, 4, B, N        'Decompose in groups of 4 years (B) + supplemental years (N)
    H = Year Mod 19                     'Gold number minus one. We can use Mod as arguments are always positive.
    R = (15 + 19 * H + S - S \ 4 - (8 * S + 13) \ 25) Mod 30 'First value of pascal residue, before correction.
    R = R - (H + 11 * R) \ 319          'Correction, if Residue is 28 or 29.
    Dnum = DateValue("21/03/" & Year) + 1 + R + (32 - S \ 4 + 2 * S + 2 * B - N - R) Mod 7
    ' From 21 March, add 1 ("Day after Good Saturday") + Residue + days to next Sunday.
    If ActiveWorkbook.Date1904 Then Dnum = Dnum - DayOffsetMacOS
    EASTER_SUNDAY = Dnum
End Function
