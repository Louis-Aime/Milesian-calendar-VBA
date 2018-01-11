# Milesian-calendar-VBA
Excel VBA functions for Milesian calendar computations

Copyright (c) Miletus, Louis-Aimé de Fouquières, 2017

MIT licence applies

MAC OS users: 
1. This package only works on versions of Excel that handle VBA, i.e. from 2013 on.
1. Check date conversion starting from 1904 if you use this epoch. Package might work from 2 January 1904 only.

## Installation
1. Create a new Excel file, save as "Excel file with macros".
1. If you don't see the "developer" menu, get access to this menu through Excel Options.
1. In "developer" menu, choose "Visual Basic" (lefmost).
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

## MilesianCalendar
Compute a system date with Milesian date elements, or retrieve Milesian date elements from a system date.
### MILESIAN_IS_LONG_YEAR (Year)
Boolean, whether the year is long (366 days) or not. 
* Year, the year in question.

A long Milesian year is just before a leap year, e.g. 2015 is a long year because 2016 is a leap year. 
With the Milesian calendar, a proposed rule is this:
years -4001, -801, 2399, 5599 etc. are *not* long. Elsewise the Gregorian rules are applied, 
e.g. 1899 is *not* long whereas 1999 is.
Remember that by mistake, dates 1/1/1900 to 29/2/1900 are wrong under Microsoft Windows.

### MILESIAN_YEAR_BASE (Year) 
Date of the day before the 1 1m of year Y, i.e. the "doomsday".

### Other functions of this module 
They work like the standard date-time functions of Excel. 
Under Microsoft (1900-calendar), negative results are handled. Excel does not display those date.
The minimum negative value is around (julian) January of year 100. 
Under MacOS, no date before 1 Jan 1904 can be handled.

* MILESIAN_YEAR, MILESIAN_MONTH, MILESIAN_DAY: the Milesian date elements of an Excel date-time stamp.
* MILESIAN_DISPLAY (D) : a string that expresses a date in Milesian.
* MILESIAN_MONTH_END : works like MONTH.END.
* MILESIAN_MONTH_SHIFT : works like MONTH.SHIFT.

## Milesian moon
Next or last mean moon. Error is +/- 6 hours for +/- 3000 years from year 2000.
### LastMoonPhase (FromDate, Moonphase)
Date of last new moon, or of other specified moon phase. Result is in Terrestrial Time.
* FromDate: Base Excel date;
* MoonPhase (0 by default): 0 for new moon, 1 for 1st quarter, 2 for full moon, 3 for last quarter.
### NextMoonPhase (FromDate, Moonphase)
Similar, but computes next moon phase.




