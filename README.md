# Milesian-calendar-VBA
Excel VBA functions for Milesian calendar computations

Copyright (c) Miletus, Louis-Aimé de Fouquières, 2017-2018

MIT licence applies

MAC OS users: 
1. This package only works on versions of Excel that handle VBA, i.e. from 2013 on.
1. Check date conversion starting from 1904 if you use this epoch. Package might work from 2 January 1904 only.

## Installation
1. Create a new Excel file, save as "Excel file with macros".
1. If you don't see the "developer" menu, get access to this menu through Excel Options.
1. In "developer" menu, choose "Visual Basic" (leftmost).
1. A Visual Basic sheet opens. In "File" menu of this sheet, choose "Import file".
1. Import the suitable .bas files of the release. You may import one or several.
1. After importing, you see new "modules", one for each .bas file. You can see the source code and comments.
1. In "File" menu, hit "Close and back to Excel".
1. In Excel file, you can call the fucntions of the modules you selected.

## Using the functions
* Hit "insert function" near the input bar.
* Choose "custom" - you can see the functions.
* If you choose one function, the parameter list appear (sorry, no help in this version).
* Functions are sensitive to "1904 Calendar" (by default on MacOS in old versions of Excel)

## Considerations on date expressions (strings representing a date) in Excel
* These functions might not work with Excel version prior to 2013.
* Excel Timestamp is a fractional number of days counted from 30 Dec 1899 00:00 Gregorian, 
with no local time consideration, even with new MacOS sheets. 
* Date expression from 1/1/1900 to 29/2/1900: by default, Excel wrongly converts these expressions 
into a time stamp representing the day before, 
and conversely displays the corresponding time stamp into wrong date expression. 
However, if specified as strings (starting with a quote) and passed to VBA,
those date expression are converted without error.
* Excel does not convert a date expression (without time) from 1/1/100 to 31/12/1899
but VBA converts it into a negative integer time stamp.
* VBA wrongly converts any date-time expression prior to 31/12/1899 00:00:00, 
the fractional part representing the time in the day is subtracted, instead of added.
* Lowest date handled is 1 Jan 100 00:00 (Gregorian), highest is 31 Dec 9999 23:59:59 (Gregorian).

Some Excel-specific functions take care of these issues.

## MilesianCalendar
Compute a system date with Milesian date elements, or retrieve Milesian date elements from a system date.

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

### MILESIAN_IS_LONG_YEAR (Year)
Boolean, whether the year is long (366 days) or not. 
* Year, the year in question.

A long Milesian year is just before a leap year, e.g. 2015 is a long year because 2016 is a leap year. 
The Milesian calendar use the Gergorian rules for leap years, with one additional rule: 
years -4000, -800, 2400, 5600 etc. (every 3200 years) are *not* leap years, 
hence -4001, -801, 2399, 5599 are *not* long Milesian years. 
Remember that by mistake, dates 1/1/1900 to 29/2/1900 are wrong under Microsoft Windows.

### MILESIAN_YEAR_BASE (Year) 
Date of the day before the 1 1m of year Y, i.e. the "doomsday".
* Year: the year whose base is to be computed.

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

### EASTER_SUNDAY (Year)
The day of Easter under Gregorian computus.
* Year: the year for which Easter Sunday is computed. An integer number, greater than 1582.

### DATE_Exceltype (Cell_Date)
Convert any cell with a date item, including from a Date1904 sheet, into a proper Excel timestamp. 
* Cell_Date: date argument, most often a date cell

## MilesianMoonPhase
Next or last mean moon. Error is +/- 6 hours for +/- 3000 years from year 2000.
### LastMoonPhase (FromDate, Moonphase)
Date of last new moon, or of other specified moon phase. Result is in Terrestrial Time.
* FromDate: Base Excel date (deemed UTC);
* MoonPhase (0 by default): 0 for new moon, 1 for 1st quarter, 2 for full moon, 3 for last quarter.
### NextMoonPhase (FromDate, Moonphase)
Similar, but computes next moon phase.

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
If specified, the year is 3-digits. Elsewhise, it is considered "current year".
If day of month is not specified, it is set to 1.
This function applied to *string* date expressions from 1/1/1900 to 28/02/1900 yields correct dates.
