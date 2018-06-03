Attribute VB_Name = "VBADateLaboratory"
'Laboratory functions, to test capacity and behavior of Date type.
'Copyright Miletus 2018
'Version M2018/02/10
'Display value of Date expressions, just as they are
Function DATEVAL_EXT(Arg As Date)
Attribute DATEVAL_EXT.VB_Description = "Extract any date value, even before 1900, as elaborated by MS Excel"
    DATEVAL_EXT = Arg
End Function

'Excel only: a function to enable computations on dates
'even before Excel epoch, and despite Microsoft software engineers' garbage
'Extract time component and set positive, as is it were added to plain date
'Date-time values parsed by Excel from a date before 30/12/1899 (Windows real 0)
'Have the time component substracted instead of added.
'The PARSE function avoids this problem, as it separates Integer Date and Time components.

Function DATE_TIME_MICROSOFT_FIX(TheDate As Date) As Date
Attribute DATE_TIME_MICROSOFT_FIX.VB_Description = "Convert into Date any Date Expression string, and fix the MS-generated problem of time reversion for dates before 30/12/1899"
If TheDate >= 0 Then      'This is the easy case
  DATE_TIME_MICROSOFT_FIX = TheDate
Else
  Dim D, H
  D = Int(TheDate)  'Highest Integer lower than TheDate
  H = 1 - (TheDate - D) 'Time part reconstructed.
  If H = 1 Then
    DATE_TIME_MICROSOFT_FIX = TheDate 'Again a good situation
  Else
    DATE_TIME_MICROSOFT_FIX = D + 1 + H 'Undo MS garbage
  End If
End If
End Function
