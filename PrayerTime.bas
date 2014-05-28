' Copyright 2014 Moustapha Cherri

' This file is part of OPT (Outlook Prayer Times).

' OPT is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as
' published by the Free Software Foundation, either version 3 of
' the License, or (at your option) any later version.

' OPT is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU Lesser General Public License for more details.

' You should have received a copy of the GNU Lesser General Public License
' along with Foobar.  If not, see <http://www.gnu.org/licenses/>.

Option Explicit

' ---------------------- Global Variables --------------------
Private PCalcMethod As String ' caculation method
Private PAsrJuristic As Integer ' Juristic method for Asr
Private PDhuhrMinutes As Integer ' minutes after mid-day for Dhuhr
Private PAdjustHighLats As Integer ' adjusting method for higher latitudes
Private PTimeFormat As Integer ' time format
Private PLat As Double ' latitude
Private PLng As Double ' longitude
Private PTimeZone As Double ' time-zone
Private PJDate As Double ' Julian date

' ------------------------------------------------------------
' Calculation Methods
Private PKarachi As String ' University of Islamic Sciences, Karachi
Private PISNA As String ' Islamic Society of North America (ISNA)
Private PMWL As String ' Muslim World League (MWL)
Private PMakkah As String ' Umm al-Qura, Makkah
Private PEgypt As String ' Egyptian General Authority of Survey
Private PCustom As String ' Custom Setting

' Juristic Methods
Private PShafii As Integer ' Shafii (standard)
Private PHanafi As Integer ' Hanafi

' Adjusting Methods for Higher Latitudes
Private PNone As Integer ' No adjustment
Private PMidNight As Integer ' middle of night
Private POneSeventh As Integer ' 1/7th of night
Private PAngleBased As Integer ' angle/60th of night

' Time Formats
Private PTime24 As Integer ' 24-hour format
Private PTime12 As Integer ' 12-hour format
Private PTime12NS As Integer ' 12-hour format with no suffix
Private PFloating As Integer ' floating point number

' Time Names
Private PTimeNames(0 To 6) As String
Private InvalidTime As String ' The string used for invalid times
    
' --------------------- Technical Settings --------------------
Private PNumIterations As Integer ' number of iterations needed to compute times

' ------------------- Calc Method Parameters --------------------
Private MethodParams As Collection

' this.methodParams[methodNum] = new Array(fa, ms, mv, is, iv);
'
' fa : fajr angle ms : maghrib selector (0 = angle; 1 = minutes after
' sunset) mv : maghrib parameter value (in angle or minutes) is : isha
' selector (0 = angle; 1 = minutes after maghrib) iv : isha parameter value
' (in angle or minutes)
Private Offsets(0 To 6)
    
Private Sub Class_Initialize()
    ' Initialize vars
    CalcMethod = "0"
    AsrJuristic = 0
    DhuhrMinutes = 0
    AdjustHighLats = 1
    TimeFormat = 0

    ' Calculation Methods
    Karachi = "1" ' University of Islamic Sciences, Karachi
    ISNA = "2" ' Islamic Society of North America (ISNA)
    MWL = "3" ' Muslim World League (MWL)
    Makkah = "4" ' Umm al-Qura, Makkah
    Egypt = "5" ' Egyptian General Authority of Survey
    Custom = "7" ' Custom Setting

    ' Juristic Methods
    Shafii = 0 ' Shafii (standard)
    Hanafi = 1 ' Hanafi

    ' Adjusting Methods for Higher Latitudes
    None = 0 ' No adjustment
    MidNight = 1 ' middle of night
    OneSeventh = 2 ' 1/7th of night
    AngleBased = 3 ' angle/60th of night

    ' Time Formats
    Time24 = 0 ' 24-hour format
    Time12 = 1 ' 12-hour format
    Time12NS = 2 ' 12-hour format with no suffix
    Floating = 3 ' floating point number

    ' Time Names
    PTimeNames(0) = "Fajr"
    PTimeNames(1) = "Sunrise"
    PTimeNames(2) = "Zuhr"
    PTimeNames(3) = "Asr"
    PTimeNames(4) = "Sunset"
    PTimeNames(5) = "Maghrib"
    PTimeNames(6) = "Isha"
    
    InvalidTime = "-----" ' The string used for invalid times
    
    ' --------------------- Technical Settings --------------------

    NumIterations = 3 ' number of iterations needed to compute times

    ' ------------------- Calc Method Parameters --------------------

    ' Tuning Offsets {fajr, sunrise, dhuhr, asr, sunset, maghrib, isha}
    Offsets(0) = 0
    Offsets(1) = 0
    Offsets(2) = 0
    Offsets(3) = 0
    Offsets(4) = 0
    Offsets(5) = 0
    Offsets(6) = 0
    
    Set MethodParams = New Collection
    ' Karachi
    ' double[] Kvalues = {18,1,0,0,18};
    ' methodParams.put(Integer.valueOf(this.getKarachi()), Kvalues);

    ' ISNA
    ' double[] Ivalues = {15,1,0,0,15};
    ' methodParams.put(Integer.valueOf(this.getISNA()), Ivalues);

    ' MWL
    ' double[] MWvalues = {18,1,0,0,17};
    ' methodParams.put(Integer.valueOf(this.getMWL()), MWvalues);

    ' Makkah
    ' double[] MKvalues = {18.5,1,0,1,90};
    ' methodParams.put(Integer.valueOf(this.getMakkah()), MKvalues);

    ' Egypt
    Dim Evalues(0 To 4) As Double
    Evalues(0) = 19.5
    Evalues(1) = 1
    Evalues(2) = 0
    Evalues(3) = 0
    Evalues(4) = 17.5
    MethodParams.Add Evalues, Egypt

    ' Custom
    ' double[] Cvalues = {18,1,0,0,17};
    ' methodParams.put(Integer.valueOf(this.getCustom()), Cvalues);
        
End Sub
' ---------------------- Trigonometric Functions -----------------------
' range reduce angle in degrees.
Private Function FixAngle(A As Double) As Double
    A = A - (360 * (Int(A / 360#)))
    If A < 0 Then
        A = A + 360
    End If
    
    FixAngle = A
End Function

' range reduce hours to 0..23
Private Function FixHour(A As Double) As Double
    A = A - 24# * Int(A / 24#)
    If A < 0 Then
        A = A + 24
    End If

    FixHour = A
End Function

' radian to degree
Private Function RadiansToDegrees(Alpha As Double) As Double
    RadiansToDegrees = ((Alpha * 180#) / PI)
End Function

' deree to radian
Private Function DegreesToRadians(Alpha As Double) As Double
    DegreesToRadians = ((Alpha * PI) / 180#)
End Function

' degree sin
Private Function DSin(D As Double) As Double
    DSin = (Sin(DegreesToRadians(D)))
End Function

' degree cos
Private Function DCos(D As Double) As Double
    DCos = (Cos(DegreesToRadians(D)))
End Function

' degree tan
Private Function DTan(D As Double) As Double
    DTan = (Tan(DegreesToRadians(D)))
End Function

Private Function Atn2(Y As Double, X As Double) As Double

    If X > 0 Then
        Atn2 = Atn(Y / X)
    ElseIf X < 0 Then
        Atn2 = Sgn(Y) * (PI - Atn(Abs(Y / X)))
    ElseIf Y = 0 Then
        Atn2 = 0
    Else
        Atn2 = Sgn(Y) * PI / 2
    End If

End Function

Public Function ArcSin(X As Double) As Double
    If (Sqr(1 - X * X) <= 0.000000000001) And (Sqr(1 - X * X) >= -0.000000000001) Then
        ArcSin = PI / 2
    Else
        ArcSin = Atn(X / Sqr(-X * X + 1))
    End If
End Function

Public Function ArcCos(X) As Double
 
    If Round(X, 8) = 1# Then ArcCos = 0#: Exit Function
    If Round(X, 8) = -1# Then ArcCos = PI: Exit Function
    ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)

End Function

' degree arcsin
Private Function DArcSin(X As Double) As Double
    DArcSin = RadiansToDegrees(ArcSin(X))
End Function

' degree arccos
Private Function DArcCos(X As Double) As Double
    DArcCos = RadiansToDegrees(ArcCos(X))
End Function

' degree arctan
Private Function DArcTan(X As Double) As Double
    DArcTan = RadiansToDegrees(Atn(X))
End Function

' degree arctan2
Private Function DArcTan2(Y As Double, X As Double) As Double
    DArcTan2 = RadiansToDegrees(Atn2(Y, X))
End Function

' degree arccot
Private Function DArcCot(X As Double) As Double
    DArcCot = RadiansToDegrees(Atn2(1#, X))
End Function

' ---------------------- Time-Zone Functions -----------------------
' compute local time-zone for a specific date
Private Function GetTimeZone1() As Double
    GetTimeZone1 = FillCalendar.TimeZone
End Function


' compute base time-zone of the system
Private Function GetBaseTimeZone1() As Double
    GetBaseTimeZone1 = FillCalendar.TimeZone
End Function

' detect daylight saving in a given date
Private Function DetectDaylightSaving() As Double
    DetectDaylightSaving = FillCalendar.DayLightSaving
End Function

' ---------------------- Julian Date Functions -----------------------
' calculate julian date from a calendar date
Private Function JulianDate(Year As Integer, Month As Integer, Day As Integer) As Double
    If (Month <= 2) Then
        Year = Year - 1
        Month = Month + 12
    End If
    
    Dim A As Double, B As Double
    
    A = Int(Year / 100#)
    B = 2 - A + Int(A / 4#)
    
    JulianDate = Int(365.25 * (Year + 4716)) + Int(30.6001 * (Month + 1)) + Day + B - 1524.5
    
End Function

' ---------------------- Calculation Functions -----------------------
' References:
' http://www.ummah.net/astronomy/saltime
' http://aa.usno.navy.mil/faq/docs/SunApprox.html
' compute declination angle of sun and equation of time
Private Function SunPosition(JD As Double) As Double()
    Dim D As Double, G As Double, Q As Double, L As Double
    
    D = JD - 2451545
    G = FixAngle(357.529 + 0.98560028 * D)
    Q = FixAngle(280.459 + 0.98564736 * D)
    L = FixAngle(Q + (1.915 * DSin(G)) + (0.02 * DSin(2 * G)))
    
    Dim E As Double, D2 As Double, RA As Double
    E = 23.439 - (0.00000036 * D)
    D2 = DArcSin(DSin(E) * DSin(L))
    RA = (DArcTan2((DCos(E) * DSin(L)), (DCos(L)))) / 15#
    RA = FixHour(RA)
    
    Dim EqT As Double, SPosition(0 To 1) As Double
    EqT = Q / 15# - RA
    SPosition(0) = D2
    SPosition(1) = EqT
    
    SunPosition = SPosition
    
End Function

' compute equation of time
Private Function EquationOfTime(JD As Double) As Double
    EquationOfTime = SunPosition(JD)(1)
End Function

' compute declination angle of sun
Private Function SunDeclination(JD As Double) As Double
    SunDeclination = SunPosition(JD)(0)
End Function
    
' compute mid-day (Dhuhr, Zawal) time
Private Function ComputeMidDay(T As Double) As Double
    Dim T2 As Double
    T2 = EquationOfTime(JDate + T)
    ComputeMidDay = FixHour(12 - T2)
End Function

' compute time for a given angle G
Private Function ComputeTime(G As Double, T As Double) As Double
    Dim D As Double, Z As Double, Beg As Double, Mid As Double, V As Double
    
    D = SunDeclination(JDate + T)
    Z = ComputeMidDay(T)
    Beg = -DSin(G) - DSin(D) * DSin(Lat)
    Mid = DCos(D) * DCos(Lat)
    V = DArcCos(Beg / Mid) / 15#
    
    If G > 90 Then
        ComputeTime = Z - V
    Else
        ComputeTime = Z + V
    End If
End Function

' compute the time of Asr
' Shafii: step=1, Hanafi: step=2
Private Function ComputeAsr(Step As Double, T As Double) As Double
    Dim D As Double, G As Double
    D = SunDeclination(JDate + T)
    G = -DArcCot(Step + DTan(Abs(Lat - D)))
    ComputeAsr = ComputeTime(G, T)
End Function

' ---------------------- Misc Functions -----------------------
' compute the difference between two times
Private Function TimeDiff(Time1 As Double, Time2 As Double) As Double
    TimeDiff = FixHour(Time2 - Time1)
End Function

' convert hours to day portions
Private Function DayPortion(Times() As Double) As Double()
    Dim I As Integer
    For I = 0 To 6
        Times(I) = Times(I) / 24
    Next
    
    DayPortion = Times
End Function

' ---------------------- Compute Prayer Times -----------------------
' compute prayer times at given julian date
Private Function ComputeTimes(Times() As Double) As Double()
    Dim T() As Double, D As Double

    T = DayPortion(Times)

    Dim Fajr As Double, Sunrise As Double, Dhuhr As Double, Asr As Double, Sunset As Double
    Dim Maghrib As Double, Isha As Double

    Fajr = ComputeTime(180 - MethodParams.Item(CalcMethod)(0), T(0))
    Sunrise = ComputeTime(180 - 0.833, T(1))
    Dhuhr = ComputeMidDay(T(2))
    Asr = ComputeAsr(1 + AsrJuristic, T(3))
    Sunset = ComputeTime(0.833, T(4))
    D = MethodParams.Item(CalcMethod)(2)
    Maghrib = ComputeTime(D, T(5))
    D = MethodParams.Item(CalcMethod)(4)
    Isha = ComputeTime(D, T(6))

    Dim CTimes(0 To 6) As Double
    CTimes(0) = Fajr
    CTimes(1) = Sunrise
    CTimes(2) = Dhuhr
    CTimes(3) = Asr
    CTimes(4) = Sunset
    CTimes(5) = Maghrib
    CTimes(6) = Isha

    ComputeTimes = CTimes

End Function
    
' the night portion used for adjusting times in higher latitudes
Private Function NightPortion(Angle As Double) As Double
    Dim Calc As Double
    
    If AdjustHighLats = AngleBased Then
        NightPortion = Angle / 60
    ElseIf AdjustHighLats = MidNight Then
        NightPortion = 0.5
    ElseIf AdjustHighLats = OneSeventh Then
        NightPortion = 0.14286
    Else
        NightPortion = 0
    End If
    
End Function

' adjust Fajr, Isha and Maghrib for locations in higher latitudes
Private Function AdjustHighLatTimes(Times() As Double) As Double()
    Dim NightTime As Double, FajrDiff As Double, D As Double
    
    NightTime = TimeDiff(Times(4), Times(1))
    
    ' Adjust Fajr
    D = MethodParams.Item(CalcMethod)(0)
    FajrDiff = NightPortion(D) * NightTime
    
    If TimeDiff(Times(0), Times(1)) > FajrDiff Then
        Times(0) = Times(1) - FajrDiff
    End If
    
    ' Adjust Isha
    Dim IshaAngle As Double, IshaDiff As Double
    If MethodParams.Item(CalcMethod)(3) = 0 Then
        IshaAngle = MethodParams.Item(CalcMethod)(4)
    Else
        IshaAngle = 18
    End If
    IshaDiff = NightPortion(IshaAngle) * NightTime
    If TimeDiff(Times(4), Times(6)) > IshaDiff Then
        Times(6) = Times(4) + IshaDiff
    End If
    
    ' Adjust Maghrib
    Dim MaghribAngle As Double, MaghribDiff As Double
    If MethodParams.Item(CalcMethod)(1) = 0 Then
        MaghribAngle = MethodParams.Item(CalcMethod)(2)
    Else
        MaghribAngle = 4
    End If
    MaghribDiff = NightPortion(MaghribAngle) * NightTime
    If TimeDiff(Times(4), Times(5)) > MaghribDiff Then
        Times(5) = Times(4) + MaghribDiff
    End If
    
    AdjustHighLatTimes = Times
    
End Function

' adjust times in a prayer time array
Private Function AdjustTimes(Times() As Double) As Double()
    Dim I As Integer
    For I = 0 To 6
        Times(I) = Times(I) + TimeZone - Lng / 15
    Next
    
    Times(2) = Times(2) + DhuhrMinutes / 60 ' Dhuhr
    If MethodParams.Item(CalcMethod)(1) = 1 Then ' Maghrib
        Times(5) = Times(4) + MethodParams.Item(CalcMethod)(2) / 60
    End If
    
    If MethodParams.Item(CalcMethod)(3) = 1 Then ' Isha
        Times(6) = Times(5) + MethodParams.Item(CalcMethod)(4) / 60
    End If
    
    If AdjustHighLats <> None Then
        Times = AdjustHighLatTimes(Times)
    End If
    
    AdjustTimes = Times
    
End Function

Private Function TuneTimes(Times() As Double) As Double()
    Dim I As Integer
    For I = 0 To 6
        Times(I) = Times(I) + Offsets(I) / 60#
    Next
    
    TuneTimes = Times
End Function

' compute prayer times at given julian date
Private Function ComputeDayTimes() As Double()
    Dim Times() As Double
    ReDim Times(0 To 6) As Double
    Times(0) = 5
    Times(1) = 6
    Times(2) = 12
    Times(3) = 13
    Times(4) = 18
    Times(5) = 18
    Times(6) = 18
    
    Dim I As Integer
    For I = 1 To NumIterations
        Times = ComputeTimes(Times)
    Next
    
    Times = AdjustTimes(Times)
    Times = TuneTimes(Times)
    
    ComputeDayTimes = Times
    
End Function

' -------------------- Interface Functions --------------------
' return prayer times for a given date
Private Property Get DatePrayerTimes(Year As Integer, Month As Integer, Day As Integer, Latitude As Double, Longitude As Double, TZone As Double) As Double()
    Lat = Latitude
    Lng = Longitude
    TimeZone = TZone
    JDate = JulianDate(Year, Month, Day)
    
    Dim LonDiff As Double
    LonDiff = Longitude / (15# * 24#)
    JDate = JDate - LonDiff
    DatePrayerTimes = ComputeDayTimes
End Property

' return prayer times for a given date
Public Property Get PrayerTimes(DT As Date) As Double()
    PrayerTimes = DatePrayerTimes(Year(DT), Month(DT), Day(DT), FillCalendar.Latitude, FillCalendar.Longitude, FillCalendar.TimeZone)
End Property

Public Property Get CalcMethod() As String
    CalcMethod = PCalcMethod
End Property

Public Property Let CalcMethod(Value As String)
    PCalcMethod = Value
End Property

Public Property Get AsrJuristic() As Integer
    AsrJuristic = PAsrJuristic
End Property

Public Property Let AsrJuristic(Value As Integer)
    PAsrJuristic = Value
End Property

Public Property Get DhuhrMinutes() As Integer
    AsrJuristic = PDhuhrMinutes
End Property

Public Property Let DhuhrMinutes(Value As Integer)
    PDhuhrMinutes = Value
End Property

Public Property Get AdjustHighLats() As Integer
    AdjustHighLats = PAdjustHighLats
End Property

Public Property Let AdjustHighLats(Value As Integer)
    PAdjustHighLats = Value
End Property

Public Property Get TimeFormat() As Integer
    TimeFormat = PTimeFormat
End Property

Public Property Let TimeFormat(Value As Integer)
    PTimeFormat = Value
End Property

Public Property Get Lat() As Double
    Lat = PLat
End Property

Public Property Let Lat(Value As Double)
    PLat = Value
End Property

Public Property Get Lng() As Double
    Lng = PLng
End Property

Public Property Let Lng(Value As Double)
    PLng = Value
End Property

Public Property Get TimeZone() As Double
    TimeZone = PTimeZone
End Property

Public Property Let TimeZone(Value As Double)
    PTimeZone = Value
End Property

Public Property Get JDate() As Double
    JDate = PJDate
End Property

Public Property Let JDate(Value As Double)
    PJDate = Value
End Property

Public Property Get Karachi() As String
    Karachi = PKarachi
End Property

Public Property Let Karachi(Value As String)
    PKarachi = Value
End Property

Public Property Get ISNA() As String
    ISNA = PISNA
End Property

Public Property Let ISNA(Value As String)
    PISNA = Value
End Property

Public Property Get MWL() As String
    MWL = PMWL
End Property

Public Property Let MWL(Value As String)
    PMWL = Value
End Property

Public Property Get Makkah() As String
    Makkah = PMakkah
End Property

Public Property Let Makkah(Value As String)
    PMakkah = Value
End Property

Public Property Get Egypt() As String
    Egypt = PEgypt
End Property

Public Property Let Egypt(Value As String)
    PEgypt = Value
End Property

Public Property Get Custom() As String
    Custom = PCustom
End Property

Public Property Let Custom(Value As String)
    PCustom = Value
End Property

Public Property Get Shafii() As Integer
    Shafii = PShafii
End Property

Public Property Let Shafii(Value As Integer)
    PShafii = Value
End Property

Public Property Get Hanafi() As Integer
    Hanafi = PHanafi
End Property

Public Property Let Hanafi(Value As Integer)
    PHanafi = Value
End Property

Public Property Get None() As Integer
    None = PNone
End Property

Public Property Let None(Value As Integer)
    PNone = Value
End Property

Public Property Get MidNight() As Integer
    MidNight = PMidNight
End Property

Public Property Let MidNight(Value As Integer)
    PMidNight = Value
End Property

Public Property Get OneSeventh() As Integer
    OneSeventh = POneSeventh
End Property

Public Property Let OneSeventh(Value As Integer)
    POneSeventh = Value
End Property

Public Property Get AngleBased() As Integer
    AngleBased = PAngleBased
End Property

Public Property Let AngleBased(Value As Integer)
    PAngleBased = Value
End Property

Public Property Get Time24() As Integer
    Time24 = PTime24
End Property

Public Property Let Time24(Value As Integer)
    PTime24 = Value
End Property

Public Property Get Time12() As Integer
    Time12 = PTime12
End Property

Public Property Let Time12(Value As Integer)
    PTime12 = Value
End Property

Public Property Get Time12NS() As Integer
    Time12NS = PTime12NS
End Property

Public Property Let Time12NS(Value As Integer)
    PTime12NS = Value
End Property

Public Property Get Floating() As Integer
    Floating = PFloating
End Property

Public Property Let Floating(Value As Integer)
    PFloating = Value
End Property

Public Property Get NumIterations() As Integer
    NumIterations = PNumIterations
End Property

Public Property Let NumIterations(Value As Integer)
    PNumIterations = Value
End Property

Public Property Get TimeNames() As String()
    TimeNames = PTimeNames
End Property
