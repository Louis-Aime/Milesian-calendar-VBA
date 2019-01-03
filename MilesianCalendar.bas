Attribute VB_Name = "MilesianCalendar"
'MilesianCalendar: Enter and display dates in Microsoft Excel following Milesian calendar conventions
'Copyright Miletus SARL 2016-2019. www.calendriermilesien.org
'For use as an MS Excel VBA module.
'Developped under Excel 2016
'Tested under MS Excel 2007 (Windows) and 2016 (Windows and MacOS)
'No warranty.
'May be used for personal or professional purposes.
'If transmitted or integrated, even with changes, present header shall be maintained in full.
'Functions are aimed at extending Date & Time functions, and use similar parameters syntax in English
'IMPORTANT VERSION NOTE
' Since MS Engineer only make garbage when using dates,
' the maximum is done here to avoid letting Excel impose his silly "Date" type.
' A special function, MICROSOFT_DATE_TIME_FIX is used for text used for entering date and time before 1 March 1900.
'Version V3 M2019-01-14

Const MStoPresentEra As Long = 693969 'Offset between 1/1m/000 epoch and Microsoft origin (1899-12-30T00:00 is 0)
Const MStoJulianMinus05 As Long = 2415018 'Offset between julian day epoch and Microsoft origin, minus 0.5
Const HighYear = 9999   'Higher year that is handles (Excel goes up to 31/12/9999 Gregorian)
Const LowYear = 100     'As long as MS does not handle date before Gregorian 1/1/100. Else we'd take -9999
Const DayOffsetMacOS As Long = 1462

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

'#Part 2: a function used internally, but available to user

Function MILESIAN_IS_LONG_YEAR(ByVal Year) As Boolean
Attribute MILESIAN_IS_LONG_YEAR.VB_Description = "Return whether year is 366 days long under milesian rules."
'Is year Year a 366 days year, i.e. a year just before a bissextile year following the Milesian rule.
If Year <> Int(Year) Or Year < LowYear Or Year > HighYear Then Err.Raise 1
Year = Year + 1
MILESIAN_IS_LONG_YEAR = PosMod(Year, 4) = 0 And (PosMod(Year, 100) <> 0 Or PosMod(Year, 400) = 0)
End Function

'#Part 3: Compute date from milesian parameters
'Date computed here is integer by construction.

Function MILESIAN_DATE(Year, Month, DayInMonth) As Date
Attribute MILESIAN_DATE.VB_Description = "Excel date from Year, Month, Day in milesian calendar."
'The Date type element from a Milesian date given as year (common era, relative), month, daynumber in month.
Dim A As Integer 'Intermediate computations as non-long integer
Dim B As Long   'Bimester number, for intermediate computations
Dim M1 As Long  'Month rank
Dim D As Long   'Days expressed in long integer
Dim Y As Integer 'Internal Year
'Check that Milesian date is OK
If Year <> Int(Year) Or Month <> Int(Month) Or DayInMonth <> Int(DayInMonth) Then Err.Raise 1
If Year >= LowYear And Year <= HighYear And Month > 0 And Month < 13 And DayInMonth > 0 And DayInMonth < 32 Then 'Basic filter
  M1 = Month - 1 'Count month rank, i.e. 0..11
  Milesian_IntegDiv M1, 2, B, M1 'B = full bimesters, M1 = 1 if a full month added, else 0
  If DayInMonth < 31 Or (M1 = 1 And (B < 5 Or MILESIAN_IS_LONG_YEAR(Year))) Then
'    Y = Year    'Set Epoch to the year 0
    A = PosDiv(Year, 4) - PosDiv(Year, 100) + PosDiv(Year, 400) 'Sum non-long terms: leap days
    D = Year         'Force long-integer conversion
    D = D * 365      'Begin computation of days in long-integer;
    D = D - MStoPresentEra - 1 + B * 61 + M1 * 30 + A + DayInMonth 'Computations in long-integer first
    MILESIAN_DATE = D
  Else
    Err.Raise 1
  End If
Else
  Err.Raise 1
End If
End Function

Function MILESIAN_YEAR_BASE(ByVal Year) As Date 'The Year base or Doomsday of a year i.e. the date just before the 1 1m of the year
Attribute MILESIAN_YEAR_BASE.VB_Description = "Date of last day before milesian year (at 00:00), for doomsday and epact."
Dim A As Integer, D As Long   'Force long integer
If Year <> Int(Year) Or Year < LowYear Or Year > HighYear Then Err.Raise 1
'Year = Year + 800    'Set Epoch to the year -800
D = Year        'Force long-integer conversion
D = D * 365     'Begin computation of days in long-integer;
A = PosDiv(Year, 4) - PosDiv(Year, 100) + PosDiv(Year, 400)
D = D - MStoPresentEra + A - 1           'Computations in long-integer first
MILESIAN_YEAR_BASE = D
End Function

'#Part 4: Extract Milesian date elements from a Date-type number or string
'IMPORTANT NOTICE: If you insert dates before 30/12/1899 with a decimal part (with a time part)
'using the Microsoft "date expression", you should use MICROSOFT_DATE_TIME_FIX on the expression,
'in order to avoid a Microsoft software engineer's garbage.

Sub Milesian_DateElement(DNum As Date, Y As Integer, M As Integer, Q As Integer, T As Variant)
' From DNum, a Date type argument, compute the milesian date element Q / M / Y (day in month, month, year)
' and also the positive decimal part H, which is the UTC time in the day.
' Y is year in common era, relative value (may be 0 or negative)
' M is milesian month number, 1 to 12
' Q is number of day in month, 1 to 31
' This is an internal subroutine. Corresponding functions come after.
' Note2: On Excel 2016, Gregorian year of Date element is greater of equal to 100.
Dim Cycle As Long, Day As Long      'Cycle is used serveral times with a different meaning each time
Day = Int(DNum)  'Initiate Day as highest integer lower or equal to DNum, and avoid control on Date type
T = DNum - Day   'Extract time part of the Date element. ALWAYS POSITIVE, DUMB MICROSOFT ENGINEERS.
Day = Day + MStoPresentEra
Milesian_IntegDiv Day, 146097, Cycle, Day    'Day is day rank in 400 years period, Cycle is quadrisaeculum
Y = Cycle * 400
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

Function MILESIAN_YEAR(TheDate As Date) As Integer  'The milesian year (common era) for a Date argument.
Attribute MILESIAN_YEAR.VB_Description = "Milesian year from Excel date value."
Dim Y As Integer, M As Integer, Q As Integer, T As Variant
Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
MILESIAN_YEAR = Y
End Function

Function MILESIAN_MONTH(TheDate As Date) As Integer  'The milesian month number (1-12) for a Date argument.
Attribute MILESIAN_MONTH.VB_Description = "Milesian month number (1 to 12), from Excel date value."
Dim Y As Integer, M As Integer, Q As Integer, T As Variant
Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
MILESIAN_MONTH = M
End Function

Function MILESIAN_DAY(TheDate As Date) As Integer  'The day number in the milesian month for a Date argument.
Attribute MILESIAN_DAY.VB_Description = "Day in milesian month , from Excel date value."
Dim Y As Integer, M As Integer, Q As Integer, T As Variant
Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
MILESIAN_DAY = Q
End Function

Function MILESIAN_TIME(TheDate As Date) As Date
Attribute MILESIAN_TIME.VB_Description = "Extract Time component of a Date-Time value, even before 1900."
'Extract date from a date element, even negative.
Dim Y As Integer, M As Integer, Q As Integer, T As Date 'Force T to a date, i.e. a time since < 1
Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
MILESIAN_TIME = T
End Function

Function MILESIAN_DISPLAY(TheDate As Date, Optional Wtime As Boolean = True) As String
Attribute MILESIAN_DISPLAY.VB_Description = "Date & time string representing the Date value in the milesian calendar. Set Wtime to False (in your language) to drop time component."
'Milesian date as a string, from a Date element
Dim Y As Integer, M As Integer, Q As Integer, T As Date 'Force T to a date, i.e. a time since < 1
Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
MILESIAN_DISPLAY = Q & " " & M & "m " & Y
If Wtime Then MILESIAN_DISPLAY = MILESIAN_DISPLAY & " " & T
End Function

'#Part 5: Computations on milesian months

Function MILESIAN_MONTH_SHIFT(TheDate As Date, MonthShift As Long) As Date 'Same date several (milesian) months later of earlier
Attribute MILESIAN_MONTH_SHIFT.VB_Description = "Find same day in Shift milesian months from TheDate."
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

Function MILESIAN_MONTH_END(TheDate As Date, MonthShift As Long) As Date 'End of month several (milesian) months later of earlier
Attribute MILESIAN_MONTH_END.VB_Description = "Find last day of month in Shift months from TheDate."
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

'#Part 6: Julian Epoch Day conversion functions

Function JULIAN_EPOCH_COUNT(TheDate As Date)
Attribute JULIAN_EPOCH_COUNT.VB_Description = "Compute decimal julian day from Excel date and time."
    Dim IntDate As Long, TimePart   'Compute separately integer part and decimal (time) part
    IntDate = Int(TheDate)  'Integer part is Date at 00:00
    TimePart = TheDate - IntDate     'Time in day
    TimePart = TimePart + 0.5       '...shifted to Julian Day convention
    IntDate = IntDate + MStoJulianMinus05 'Epoch shift
    JULIAN_EPOCH_COUNT = TimePart + IntDate
End Function

Function JULIAN_EPOCH_DATE(Julian_Count) As Date
Attribute JULIAN_EPOCH_DATE.VB_Description = "Convert decimal julian day to Excel date-time count."
    Dim IntDate, TimePart As Date
    IntDate = Int(Julian_Count)       'Integer part of Julian Day
    TimePart = Julian_Count - IntDate 'Decimal part, i.e. time after noon
    TimePart = TimePart + 0.5 'Add, not substract, a half day
    IntDate = IntDate - MStoJulianMinus05 - 1 'Compensate full day added from above
    JULIAN_EPOCH_DATE = TimePart + IntDate
End Function

'#Part 7: Excel specific functions

Function DATE_Exceltype(TheDate As Date) As Double
Attribute DATE_Exceltype.VB_Description = "Convert general Date to Excel current series number - Usefull if Date1904 is set."
'Change to a MacOS Date count from a Date object, if Date1904 is set. Else just pass date.
'No control. You can generate a foolish date from a negative number, as you can do without this function.
If ActiveWorkbook.Date1904 Then
    DATE_Exceltype = TheDate - DayOffsetMacOS
Else
    DATE_Exceltype = TheDate
End If
End Function

Function DAYOFWEEK_Ext(TheDate As Date, Optional DispType As Integer) As Integer 'Milesian way: Sunday = 0, Monday = 1, up to Saturday = 6
Attribute DAYOFWEEK_Ext.VB_Description = "Day of week of Date object. Type- 0 (default) 0-6, 0=Sunday; 1 (Excel default) 1-7, 1=Sunday; Others: same as Excel's"
    Dim IntDate As Long, Start As Integer, Phase As Integer
    
    '1. Compute Start and Phase from DispType
    If IsMissing(DispType) Then DispType = 0    'This option value is not used with standard DOW routines
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
    IntDate = Int(TheDate)  'Convert date-time to hold date component only
    DAYOFWEEK_Ext = PosMod(IntDate + Phase, 7) + Start
    
End Function

Function EASTER_SUNDAY(Year As Integer) As Date 'Easter Date computed after the Milesian method www.calendriermilesien.org
Attribute EASTER_SUNDAY.VB_Description = "Date of Easter Sunday (Gregorian computus) for the given Year"
Dim S As Long, B As Long, N As Long, H As Integer, R As Integer 'Components of year, Golden number minus 1, Easter residue
If Year < 1583 Or Year > HighYear Then Err.Raise 1
Milesian_IntegDiv Year, 100, S, N   'Decompose Year in centuries (S) + years in century (N)
Milesian_IntegDiv N, 4, B, N        'Decompose in groups of 4 years (B) + supplemental years (N)
H = Year Mod 19                     'Gold number minus one. We can use Mod as arguments are always positive.
R = (15 + 19 * H + S - S \ 4 - (8 * S + 13) \ 25) Mod 30 'First value of pascal residue, before correction.
R = R - (H + 11 * R) \ 319          'Correction, if Residue is 28 or 29.
EASTER_SUNDAY = DateValue("21/03/" & Year) + 1 + R + (32 - S \ 4 + 2 * S + 2 * B - N - R) Mod 7
' From 21 March, add 1 ("Day after Good Saturday") + Residue + days to next Sunday.
End Function

