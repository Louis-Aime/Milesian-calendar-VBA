Attribute VB_Name = "DateParse"
'DateParse.
'Copyright Miletus SARL 2018. www.calendriermilesien.org
'For use as an MS Excel VBA module.
'Developped under Excel 2016
'Tested under MS Excel 2007 (Windows) and 2016 (Windows and MacOS)
'No warranty of conformance to business objectives
'Neither resale, neither transmission, neither compilation nor integration in other modules
'are permitted without authorisation
'DATE_PARSE, applied to any Date-Time expression as a string, returns the corresponding Excel Date series number.
'EASTER_SUNDAY (Excel only): for a given year > 1582, date of Easter sunday after Gregorian computus.
'(This function is available on Libre Office Calc).
'This module uses MilesianCalendar
'Do not apply to an Excel formatted date between 1 Jan and 29 Feb 1900. Hint: force input to Text format.
'Version of this module: M2018-02-10

Function DATE_PARSE(TheCell As String) As Date 'A String is converted into a Date
Attribute DATE_PARSE.VB_Description = "Extract date and time value from string. Standard or Milesian expression. Earliest is Greg. 1/1/100. Fix bug for Date-time expressions before 1900."
'TheCell holds a valid string for a date, either in milesian, either in a specific calendar,
'or in the standard calendar (gregorian).
'TheCell is supposed a numeric date string. It shall not work with months expressed by their names.
'Summary:
'   1. Prepare, convert all in uppercase, trim
'   2. Extract calendar indication when available.
'   3. Extract Time part, store in T
'   4. Split Date part into minimum 2 elements
'   5. Find Month element after its pattern ("M" month marker)
'   6. Find Year element, of provide present year
'   7. Find Day element of provide "1" by default, return
'   Note: authorised delimiters are "./-". Comma is dropped. m of Milesian month is recognised. Other lead to error.
'   If as first character: canvas is yyy/mm/dd or similar, even M-yyy-mm-dd is possible.
'   If not as first, find and extract "m", compute d, m and year elements, call MILESIAN_DATE.

'1. Prepare: Convert string to Uppercase and drop blanks in excess
Dim Elem, I    'Eleme: the array of elements of TheCell when splitted, I: a standard index
Dim T As Date, Y, M, D    'The date elements
Dim Yindex, Mindex, Dindex    'Index of year, month and day in Elem, -1 means "unknown"
Dim Delimiters      'The list of possible delimiters (outside " ")
Delimiters = Array("/", ".", "-")

TheCell = Trim(TheCell) 'Suppress all leading and trailing blank
TheCell = UCase(TheCell) 'All uppercase
'For Open Office: Drop initial "'"
If Left(TheCell, 1) = "'" Then TheCell = Right(TheCell, Len(TheCell) - 1)
TheCell = Replace(TheCell, ", ", " ")    'Drop comma followed by blank (which is authorised)
Elem = Split(TheCell)   'Split with " "
TheCell = ""            'Reconstruct TheCell
For I = LBound(Elem) To UBound(Elem)
    If Len(Elem(I)) > 0 Then TheCell = TheCell & Switch(Len(TheCell) = 0, "", Len(TheCell) > 0, " ") & Elem(I)
    Next I

'2. Did the user specify the calendar with the first character ?
'Implementation note: it is possible to add other tags here, e.g. "J" for julian calendar.

Dim K As String     'Calendar code, First letter of TheCell if it is a letter
Select Case Left(TheCell, 1)
    Case "M"
        K = "M"     'At least, year in first place, month without "m" in second place
        Yindex = 0
        Mindex = 1
        Dindex = -1 'Day may be not present, to be checked later
    Case "-", "0" To "9"   'Fist character is one of plain number, including a possible minus sign
        K = "D" 'Not specified, default, order is not known
        Yindex = -1
        Mindex = -1
        Dindex = -1
    'Other calendar could come here as other case tags
    Case Else
        K = "U" 'Specified but unknown
    End Select
If K <> "D" Then TheCell = Right(TheCell, Len(TheCell) - 1) 'Drop the first character from TheCell

'3. Extract Time element. This is necessary here since DateValue ignores the time part.
Elem = Split(TheCell)   'Split again, knowing that there is no empty element
'Check whether last element contains ":", if yes it is the "time" part, else no time part.
If InStr(Elem(UBound(Elem)), ":") > 0 Then    'Last elements holds a time string
    T = TimeValue(Elem(UBound(Elem)))
    'Drop Time value from TheCell and re-compute Elem
    TheCell = ""
    For I = LBound(Elem) To UBound(Elem) - 1
        If Len(Elem(I)) > 0 Then TheCell = TheCell & Switch(Len(TheCell) = 0, "", Len(TheCell) > 0, " ") & Elem(I)
        Next I
    Elem = Split(TheCell, " ", 4)
Else
    T = TimeValue("00:00:00")
End If
'Here T holds the time part, TheCell holds the date part, Elem holds a possible detailed decomposition.
    
'4. Extract Date elements, whichever the separator is
'4.1 To begin with, solve case where significant cell begins with "-" and consider this first part as year.
If Left(TheCell, 1) = "-" Then
    If K = "D" Then Err.Raise 1 'No "free" date string shall begin with a minus sign !!
    Yindex = 0          'Year found
    Y = Val(TheCell)    'Year value (negative) is first part of TheCell
    TheCell = Right(TheCell, Len(TheCell) - 1)  'Drop initial "-" which is not a separator
End If
'Extract remaining part
I = UBound(Elem)
Do While UBound(Elem) = LBound(Elem) And I <= UBound(Delimiters) 'only one element?  Try other delimiters
  Elem = Split(TheCell, Delimiters(I), 3)
  I = I + 1
  Loop
If UBound(Elem) = LBound(Elem) Then Err.Raise 1  'We did not find how to split
If Y < 0 Then Elem(Yindex) = Y  'Meanwhile we found the year, a negative value

'5. Check whether one element, and only one, is a milesian month notation:
'Search elements looking like "1M" to "12M", as long as there is no other indication
For I = LBound(Elem) To UBound(Elem)    'Examine each element
  If Right(Elem(I), 1) = "M" And ((Val(Elem(I)) > 0 And Val(Elem(I)) < 10 And Len(Elem(I)) = 2) Or ((Val(Elem(I)) >= 10 And Val(Elem(I)) <= 12 And Len(Elem(I)) = 3))) Then 'A Milesian month
    If K = "D" Then 'Calendar still not defined
        K = "M" 'Milesian calendar
        Mindex = I  'This is the month's index
        Elem(I) = Left(Elem(I), Len(Elem(I)) - 1) 'Set this element to a pure number
       Else
        Err.Raise 1     'Only one month indication authorised
       End If
     End If
  Next I

'6. Search for year element: three (numeric) character. whether first or last element, or non-existent
'Note: this part is not valid if we authorise month names, for other calendars
If Yindex = -1 Then 'Year still not found
  For I = LBound(Elem) To UBound(Elem)    'Examine each element
    If Len(Elem(I)) >= 3 And IsNumeric(Elem(I)) Then    'This can represent a year
        If Yindex = I Or Yindex = -1 Then   'Year field recognised
            Yindex = I
          Else                              'Only one year field authorised
            Err.Raise 1
          End If
      End If
    Next I
  If K = "M" And Yindex = -1 And UBound(Elem) = 2 Then Err.Raise 1 'No 2-char year authorised in Milesian notation
End If

'7. Find whether there is a Day indication, make last computations and return
Select Case K
    Case "D"
        'Here a piece of code to be placed when Date objects hold near year 0 date, and if MS does not compute them well.
        DATE_PARSE = DateValue(TheCell) + T    'Standard procedure for non-specified calendar
     Case "M"    'At this level, Mindex is known  (>-1). Find and check other elements.
        If Mindex > 1 Then Err.Raise 1  'Month may never be indicated as 3rd element
        M = Val(Elem(Mindex))
        If Yindex = -1 Then 'Year is not specified, provide with today's date
            Dim MyDate As Date
            MyDate = Date   'Today's date
            Y = MILESIAN_YEAR(MyDate)
        Else
            Y = Val(Elem(Yindex))
        End If
        'Find place of day or set default day
        I = LBound(Elem)
        Do While Dindex = -1 And I <= UBound(Elem)
            If I <> Yindex And I <> Mindex Then Dindex = I  'Found
            I = I + 1
        Loop
        If Dindex > -1 Then 'D found
            If IsNumeric(Elem(Dindex)) Then
                D = Val(Elem(Dindex))
            Else
                Err.Raise 1
            End If
        Else    'D was not specified
            D = 1
        End If
        DATE_PARSE = MILESIAN_DATE(Y, M, D) + T
    Case Else
        Err.Raise 1
    End Select

End Function
