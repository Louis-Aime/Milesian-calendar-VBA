# Milesian-calendar-VBA
Excel VBA functions for Milesian calendar computations

Copyright (c) Miletus, Louis-Aimé de Fouquières, 2017-2020

MIT licence applies

MAC OS users: 
1. This package only works on versions of Excel that handle VBA, i.e. from 2011 (or 2013 ?) on.
1. Date conversion starting from 1904 is handled, but calling Milesian functions from VBA packages
might not work with this option.

## Installation
1. Create a new Excel file, save as "Excel file with macros".
1. If you don't see the "developer" menu, get access to this menu through Excel Options.
1. In "developer" menu, choose "Visual Basic" (leftmost).
1. A Visual Basic sheet opens. In "File" menu of this sheet, choose "Import file".
1. Import the suitable .bas files of the release. You may import one or several.
1. After importing, you see new "modules", one for each .bas file. You can see the source code and comments.
1. In "File" menu, hit "Close and back to Excel".
1. In Excel file, you can call the functions of the modules you selected.

## Using the functions
* Hit "insert function" near the input bar.
* Choose "custom" - you can see the functions.
* If you choose one function, the parameter list appears (sorry, no help in this version).
* Functions are sensitive to "1904 Calendar" (by default on MacOS in old versions of Excel)

## Considerations on Date object and date expressions (strings representing a date) in Excel
* Excel Date object is a fractional number of days counted from 30 Dec 1899 00:00 Gregorian, 
with no time zone consideration.
* Dates that Excel displays 1/1/1900 to 29/2/1900 are one day in advance to real date.
However, Excel converts properly any *string* expressing dates from 1/1/1900 to 28/2/1900
into the right date object passed to VBA. 
* Excel does not display any date object from 1/1/100 to 31/12/1899, 
but VBA converts such date expressions (string values) into a negative Date object.
* Time part of a VBA Date object for a date prior to 31/12/1899 00:00:00 is a *negative* value.
the fractional part representing the time in the day is subtracted, instead of added.
* Lowest date handled is 1 Jan 100 00:00 (Gregorian), highest is 31 Dec 9999 23:59:59 (Gregorian).

The Milesian functions handle these issues. 

## MilesianCalendar
Compute a system date with Milesian date elements, or retrieve Milesian date elements from a system date.
Display any date in Milesian.
Retrieve key elements for a Milesian year.
Compute date shift, duration between dates, Milesian month shift, end of Milesian months.
Compute Day of week.
Compute next or last mean moon phase. Error is +/- 6 hours for +/- 3000 years from year 2000.

### Excel-similar functions of this module 
They work like the standard date-time functions of Excel. 
Under Microsoft (1900-calendar), negative results are handled. Excel does not display those date.
The minimum negative value is Gregorian 1 January of year 100. 
For old MacOS sheets, no date before 1 Jan 1904 can be handled.

* MILESIAN_YEAR, MILESIAN_MONTH, MILESIAN_DAY: the Milesian date elements of an Excel date-time stamp.
* MILESIAN_DATE (Year, Month, Day_in_month): the time stamp (at 00:00) of a Milesian date given by its elements.
* MILESIAN_TIME: the time part of a time stamp; works with dates prior to 30/12/1899. 
The time is always the positive fractional part of the time stamp, even if the stamp is negative.
* MILESIAN_DISPLAY (Date, Wtime) : a string that expresses a date in Milesian.
If optional Wtime is *true* or missing, time part is added to string.
* MILESIAN_MONTH_END : works like MONTH.END.
* MILESIAN_MONTH_SHIFT : works like MONTH.SHIFT.

### DATE_SHIFT
Date, the date after a duration
* StartDate, start date...
* Shift: a duration in decimal days to add to the date. Target date will be before start date if Shift is negative.

### DURATION
A number of decimal days between two dates. Negative if begin date is after end date.
* Begin date
* End date

### MILESIAN_IS_LONG_YEAR (Year)
Boolean, whether the year is long (366 days) or not. 
* Year, the year in question.

A long Milesian year is just before a leap year, e.g. 2015 is a long year because 2016 is a leap year. 
The Milesian calendar use the Gregorian rules for leap years.
Remember that by mistake, dates 1/1/1900 to 29/2/1900 are wrong under Microsoft Windows.

### MILESIAN_YEAR_BASE (Year) 
Date of the day before the 1 1m of year Y, i.e. the "doomsday", at 7:30 (a.m.) for moon computations.
* Year: the year whose base is to be computed.

### MILESIAN_DOOMSDAY (Year, Option)
The day of the week that is common to all "key-day" of the Milesian or Gregorian year.
* Year: the year for which the doomsday is computed
* Option: a number (default 0), used as DAYOFWEEK_Ext

### JULIAN_EPOCH_COUNT (Date)
Decimal Julian Day from Excel time stamp, deemed UTC date. 
* Date: the date to convert.

### JULIAN_EPOCH_DATE (Count)
Excel time stamp (Date type) representing the UTC Date from a fractional Julian Day.
* Count: fractional Julian Day to convert.

### DAYOFWEEK_Ext (Date, Option)
The day of the week for the Date, with another default option.
* Date: the date whose day of week is computed
* Option: a number; default or 0 means 0 = Sunday, Monday = 1, etc., Saturday = 6; 
1 is Excel's DAYOFWEEK's default option meaning 1 = Sunday, 2 = Monday, etc., Saturday = 7;
2, 3, 11 to 17, are the same as Excel's DAYOFWEEK's options.

### GREGORIAN_EPACT (Year)
The Gregorian Epact, i.e. the Ecclesiastic Moon age the eve of 1 March (and of 1 January of a "Julian" Year, i.e. a year of 365,25 days).
THe Gregorian Epact is an integer number in the range 0 to 29 inclusive. The date of Easter is computed after this number.

### MILESIAN_EPACT (Year)
The Gregorian Epact shifted to the eve of 1 1m. Always (Gregorian_Epact - 11) modulo 30.

### EASTER_SUNDAY (Year)
The day of Easter under Gregorian computus.
* Year: the year for which Easter Sunday is computed. An integer number, greater than 1582.

### LastMoonPhase (FromDate, Moonphase)
Date of last new moon, or of other specified moon phase. Result is in Terrestrial Time.
* FromDate: Base Excel date (deemed UTC);
* MoonPhase (0 by default): 0 for new moon, 1 for 1st quarter, 2 for full moon, 3 for last quarter.
### NextMoonPhase (FromDate, Moonphase)
Similar, but computes date of next moon phase.

## DateParse
This module has only a string parser, that converts a (numeric) Gregorian or Milesian date or date-time expression 
into an Excel time stamp. 
### DATE_PARSE (String)
Date (Excel time stamp) corresponding to a date expression
* String: holds the date expression. 
This parser recognises a date expression, Gregorian or Milesian. 
It is a Milesian date expression if either the month number ends with "m" (and without leading 0), 
or if the complete string begins with "M", in which case elements must be in the order year, month, date.
Due to Excel, no date before year 100 can be handled. 
Separators between date elements must be the same (except comma with spaces). 
It is possible to specify only 2 date elements, including the month. 
If specified, the year is 3-digits. If missing, it is considered "current year".
If day of month is not specified, it is set to 1.
This function applied to *string* date expressions from 1/1/1900 to 28/02/1900 yields correct dates.
