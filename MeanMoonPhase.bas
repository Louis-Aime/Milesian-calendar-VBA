Attribute VB_Name = "MeanMoonPhase"
'MeanMoonPhase: Find a moon phase near a date.
'Copyright Miletus SARL 2017-2018. www.calendriermilesien.org
'For use as an MS Excel VBA module.
'Tested under MS Excel 2007 (Windows) and 2016 (Windows and MacOS).
'No warranty.
'If transmitted, even with changes, present header shall be maintained in full.
'This package uses no other element.
'Note: the moon computed here is a mean moon, with no Terrestrial Time correction.
'The error on mean moon is less than one day for three milliniums before and after year 2000.
'Functions:
'   MOON_PHASE_LAST : last time the mean moon reached a given phase
'   MOON_PHASE_NEXT : next time where the mean moon reaches a given phase
'Parameters
'   FromDate: a date and time UTC expressed in Excel
'   MoonPhase: 0 or omitted: new moon; 1: first quarter; 2: full moon; 3: last quarter; <0 or >3: error.
'Version V2.0 M2018-02-10

Const MeanSynodicMoon As Double = 29.53058883 'Mean duration of a synodic month in decimal days = 29d 12h 44mn 2s 7/8s
Const MeanNewMoon2000 As Double = 36531.59773 'G2000-01-06T14-20-44 TT, conventional date of a mean new moon

Function MOON_PHASE_LAST(FromDate As Date, Optional MoonPhase As Integer) As Date
Attribute MOON_PHASE_LAST.VB_Description = "Date of last mean moon phase. Phase 0 or omitted: New Moon, 1: First Q, 2: Full Moon, 3: Last Q."
'Date of last mean moon phase before FromDate (in UTC).
'MoonPhase: 0 or omitted, new moon; 1: first quarter; 2: full moon; 3: last quarter; else: error.
Dim Phase As Double
If IsMissing(MoonPhase) Then MoonPhase = 0
If MoonPhase < 0 Or MoonPhase > 3 Then Err.Raise 1
Phase = FromDate        'Force conversion to Double in order to avoid Date type controls
Phase = Phase - MeanNewMoon2000 - (MoonPhase / 4) * MeanSynodicMoon
 While Phase < 0
    Phase = Phase + MeanSynodicMoon
 Wend
 While Phase >= MeanSynodicMoon
    Phase = Phase - MeanSynodicMoon
 Wend
MOON_PHASE_LAST = FromDate - Phase
End Function

Function MOON_PHASE_NEXT(FromDate As Date, Optional MoonPhase As Integer) As Date
Attribute MOON_PHASE_NEXT.VB_Description = "Date of next mean moon phase. Phase 0 or omitted: New Moon, 1: First Q, 2: Full Moon, 3: Last Q."
'Date of next mean moon phase after FromDate (in UTC).
'MoonPhase: 0 or omitted, new moon; 1: first quarter; 2: full moon; 3: last quarter; else: error.
Dim Phase As Double
If IsMissing(MoonPhase) Then MoonPhase = 0
If MoonPhase < 0 Or MoonPhase > 3 Then Err.Raise 1
Phase = FromDate        'Force conversion to Double in order to avoid Date type controls
Phase = MeanNewMoon2000 - Phase + (MoonPhase / 4) * MeanSynodicMoon
 While Phase < 0
    Phase = Phase + MeanSynodicMoon
 Wend
 While Phase >= MeanSynodicMoon
    Phase = Phase - MeanSynodicMoon
 Wend
MOON_PHASE_NEXT = FromDate + Phase
End Function
