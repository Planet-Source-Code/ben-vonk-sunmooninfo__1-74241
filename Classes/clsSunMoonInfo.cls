VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSunMoonInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Calculations are made with "Astronomical Algorithms" by Jean Meeuws
'The routines for MoonPhase and MoonPosition are adopted from
'Thomas A. Swift, Lunar Calculator 2.1 and can be found at:
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=66957&lngWId=1

'SunMoonInfo Class
'
'Author Ben Vonk
'15-01-2012 First version, calculates: Sunrise, Suntransit, Sunset, Equation of Time and Sun Declination for the given Latitude, Longitude and TimeZone
'31-01-2012 Second version, fixed some bugs, add: SunTime, Twilight times and SunPosition for given time
'03-03-2012 Third version, add: MoonInfo, name SunInfo changed in SunMoonInfo

Option Explicit

' Public Events
Public Event ResultCityInfo(Name As String, Index As Long, Latitude As String, Longitude As String, TimeZone As Double)
Public Event ResultMoonCalculate(IsDate As Date, Moonrise As Date, Moontransit As Date, Moonset As Date, MoonsetInfo As MoonStatus, MoonriseInfo As MoonStatus, MoonStatus As MoonStatus, MoonAge As Double)
Public Event ResultMoonPhases(IsDate As Date, NewMoon As Date, FirstQuarter As Date, FullMoon As Date, LastQuarter As Date, PreviousPhases As Boolean)
Public Event ResultMoonPosition(IsDate As Date, MoonIllumination As Double, MoonAngle As Double, MoonAge As Double)
Public Event ResultSunCalculate(IsDate As Date, Sunrise As Date, Suntransit As Date, Sunset As Date, SunTime As Date, CivilTwilightBegin As Date, CivilTwilightEnd As Date, NauticalTwilightBegin As Date, NauticalTwilightEnd As Date, AstronomicalTwilightBegin As Date, AstronomicalTwilightEnd As Date, EquationOfTime As Double, SunDeclination As Double)
Public Event ResultSunPosition(IsDate As Date, SunAzimuth As Double, SunZenith As Double, SunElevation As Double)

' Private Constants
Private Const MOON_SYNODIC_PERIOD As Double = 29.5305888843999
Private Const PI                  As Double = 3.14159265358979
Private Const PI_DOUBLE           As Double = 6.28318530717958
Private Const ERROR_VALUE         As Long = &HFC19
Private Const FILENAME_CITYINFO   As String = "CityInfo.dat"

' Public Enumerations
Public Enum Hemispheres
   North
   South
End Enum

Public Enum MoonDescriptions
   WaxingGibbous
   WaningGibbous
   WaningCrescent
   WaxingCrescent
End Enum

Public Enum MoonStatus
   Done
   None
   AboveHorizon
   BelowHorizon
   Up
   Down
   Rising
   Setting
   UpAllDay
   DownAllDay
End Enum

' Private Type
Private Type CityInfoType
   Name                           As String
   Longitude                      As Double
   Latitude                       As Double
   TimeZone                       As Double
End Type

' Private Variables
Private LatitudeDegMinSec         As Boolean
Private LongitudeDegMinSec        As Boolean
Private m_DaylightSavingTime      As Boolean
Private m_TimeInUTC               As Boolean
Private CityBuffer                As CityInfoType
Private CityInfo()                As CityInfoType
Private m_DateCalculate           As Date
Private TimeDaylightSavingTime    As Double
Private EquationOfTime            As Double
Private m_Longitude               As Double
Private m_Latitude                As Double
Private m_TimeZone                As Double
Private RadianAngle               As Double
Private SunDeclination            As Double
Private SunDepression             As Double
Private m_Hemisphere              As Hemispheres
Private m_CityCount               As Long
Private m_MapCityInfo             As String

Public Property Get DateCalculate() As Date

   DateCalculate = m_DateCalculate

End Property

Public Property Let DateCalculate(ByVal NewDateCalculate As Date)

   m_DateCalculate = NewDateCalculate

End Property

Public Property Get DaylightSavingTime() As Boolean

   DaylightSavingTime = m_DaylightSavingTime

End Property

Public Property Let DaylightSavingTime(ByVal NewDaylightSavingTime As Boolean)

   m_DaylightSavingTime = NewDaylightSavingTime

End Property

Public Property Get Latitude() As Variant

   If LatitudeDegMinSec Then
      Latitude = DegMinSecToDecDeg(CheckCoordinate(m_Latitude, True))
      
   Else
      Latitude = m_Latitude
   End If

End Property

Public Property Let Latitude(ByVal NewLatitude As Variant)

   If IsNumeric(NewLatitude) Then
      m_Latitude = CDbl(ValidateValue(NewLatitude))
      LatitudeDegMinSec = False
      
   ElseIf VarType(NewLatitude) = vbString Then
      m_Latitude = DegMinSecToDecDeg(CheckCoordinate(NewLatitude, True))
      LatitudeDegMinSec = True
      
   Else
      Exit Property
   End If
   
   If Not CheckDistance("Latitude", m_Latitude) Then
      m_Latitude = 0
      m_Hemisphere = North
      LatitudeDegMinSec = False
      
   Else
      m_Hemisphere = CalcHemisphere
   End If

End Property

Public Property Get Longitude() As Variant

   If LongitudeDegMinSec Then
      Longitude = DegMinSecToDecDeg(CheckCoordinate(m_Longitude))
      
   Else
      Longitude = m_Longitude
   End If

End Property

Public Property Let Longitude(ByVal NewLongitude As Variant)

   If IsNumeric(NewLongitude) Then
      m_Longitude = CDbl(ValidateValue(NewLongitude))
      LongitudeDegMinSec = False
      
   ElseIf VarType(NewLongitude) = vbString Then
      m_Longitude = DegMinSecToDecDeg(CheckCoordinate(NewLongitude))
      LongitudeDegMinSec = True
      
   Else
      Exit Property
   End If
   
   If Not CheckDistance("Longitude", m_Longitude) Then
      m_Longitude = 0
      LongitudeDegMinSec = False
   End If

End Property

Public Property Get MapCityInfo() As String

   MapCityInfo = Left(m_MapCityInfo, Len(m_MapCityInfo) - 1)

End Property

Public Property Let MapCityInfo(ByVal NewMapCityInfo As String)

   If Right(NewMapCityInfo, 1) <> "\" Then NewMapCityInfo = NewMapCityInfo & "\"
   
   If Len(Dir(NewMapCityInfo & FILENAME_CITYINFO)) Then
      m_MapCityInfo = NewMapCityInfo
      
      Call LoadCityInfo
   End If

End Property

Public Property Get TimeInUTC() As Boolean

   TimeInUTC = m_TimeInUTC

End Property

Public Property Let TimeInUTC(ByVal NewTimeInUTC As Boolean)

   m_TimeInUTC = NewTimeInUTC

End Property

Public Property Get TimeZone() As Double

   TimeZone = m_TimeZone

End Property

Public Property Let TimeZone(ByVal NewTimeZone As Double)

   m_TimeZone = NewTimeZone
   
   If CheckDistance("TimeZone", m_TimeZone) Then
      m_TimeZone = CheckTimeZone(m_TimeZone, m_Longitude)
      
   Else
      m_TimeZone = 0
   End If

End Property

Public Function CityAdd(ByVal Name As String, ByVal Latitude As Variant, ByVal Longitude As Variant, ByVal TimeZone As Double, Optional ByVal Save As Boolean) As Variant

Dim strError As String

   If CityIndex(Name) > -1 Then
      strError = "Name: " & Name & " already exist!"
      
   Else
      strError = FillCityInfo(Name, Latitude, Longitude, TimeZone)
   End If
   
   If Len(strError) Then
      CityAdd = "Error - " & Trim(strError)
      
   Else
      Call AddCityInfo
      
      If Save Then Call SaveCityInfo
      
      CityAdd = CityIndex(Name)
   End If

End Function

Public Function CityChange(ByRef City As Variant, ByVal Name As String, ByVal Latitude As Variant, ByVal Longitude As Variant, ByVal TimeZone As Double, Optional ByVal Save As Boolean) As Variant

Dim lngIndex As Long
Dim strError As String

   If IsNumeric(City) Then
      lngIndex = City
      
   ElseIf VarType(City) = vbString Then
      lngIndex = CityIndex(City)
      
   Else
      lngIndex = -1
   End If
   
   If (lngIndex < 0) And (lngIndex > m_CityCount - 1) Then
      strError = "City: " & City & "?"
      
   Else
      strError = FillCityInfo(Name, Latitude, Longitude, TimeZone)
   End If
   
   If Len(strError) Then
      CityChange = "Error - " & Trim(strError)
      
   Else
      CityInfo(lngIndex) = CityBuffer
      QuickSort CityInfo(), LBound(CityInfo), UBound(CityInfo)
      
      If Save Then Call SaveCityInfo
      
      CityChange = CityIndex(Name)
   End If

End Function

Public Function CityCount() As Long

   CityCount = m_CityCount

End Function

Public Function CityDelete(ByVal City As Variant, Optional ByVal Save As Boolean) As Variant

Dim lngIndex As Long

   If IsNumeric(City) Then
      lngIndex = CLng(City)
      
   ElseIf VarType(City) = vbString Then
      lngIndex = CityIndex(City)
      
      If (lngIndex < 0) Or (lngIndex > m_CityCount - 1) Then lngIndex = -1
      
   Else
      lngIndex = -1
   End If
   
   If m_CityCount = 0 Then
      CityDelete = "Error - No more cities available!"
      
   ElseIf lngIndex = -1 Then
      CityDelete = "Error - City: " & City & "?"
      
   Else
      CityInfo(lngIndex).Name = String(255, 255)
      QuickSort CityInfo(), LBound(CityInfo), UBound(CityInfo)
      m_CityCount = m_CityCount - 1
      
      If m_CityCount > 0 Then
         ReDim Preserve CityInfo(m_CityCount - 1) As CityInfoType
         
      Else
         ReDim CityInfo(0) As CityInfoType
      End If
      
      If Save Then Call SaveCityInfo
      
      CityDelete = True
   End If

End Function

Public Function CityGet(Optional ByVal City As Variant = True, Optional ByVal CoordinatesAsDegreesMinutesSeconds As Boolean, Optional ByVal CoordinatesMoreAccurate As Boolean = True) As Boolean

Dim lngIndex     As Long
Dim lngStart     As Long
Dim lngStop      As Long
Dim strLatitude  As String
Dim strLongitude As String

   If IsNumeric(City) Then
      lngIndex = CLng(City)
      
      ' Get info for all cities
      If City = True Then
         lngStart = 0
         lngStop = m_CityCount - 1
         
      ElseIf (lngIndex < 0) Or (lngIndex > m_CityCount - 1) Then
         Exit Function
         
      Else
         lngStart = lngIndex
         lngStop = lngStart
      End If
      
   ElseIf VarType(City) = vbString Then
      lngStart = CityIndex(City)
      lngStop = lngStart
      
      If lngStart = -1 Then Exit Function
      
   Else
      Exit Function
   End If
   
   For lngIndex = lngStart To lngStop
      With CityInfo(lngIndex)
         strLatitude = CStr(.Latitude)
         strLongitude = CStr(.Longitude)
         
         If CoordinatesAsDegreesMinutesSeconds Then
            strLatitude = CheckCoordinate(DecDegToDegMinSec(strLatitude, CoordinatesMoreAccurate), True)
            strLongitude = CheckCoordinate(DecDegToDegMinSec(strLongitude, CoordinatesMoreAccurate))
         End If
         
         RaiseEvent ResultCityInfo(.Name, lngIndex, strLatitude, strLongitude, .TimeZone)
      End With
   Next 'lngIndex
   
   CityGet = True

End Function

Public Function CityIndex(ByVal CityName As String) As Long

Dim lngIndex As Long

   CityName = UCase(Trim(CityName))
   
   For lngIndex = 0 To UBound(CityInfo)
      If Left(UCase(CityInfo(lngIndex).Name), Len(CityName)) = CityName Then Exit For
   Next 'lngIndex
   
   If lngIndex > m_CityCount - 1 Then lngIndex = -1
   
   CityIndex = lngIndex

End Function

Public Function CityName(ByVal CityIndex As Long) As Variant

   If (CityIndex < 0) Or (CityIndex > m_CityCount) Then
      CityName = False
      
   Else
      CityName = CityInfo(CityIndex).Name
   End If

End Function

Public Function CitySet(ByRef City As Variant) As Boolean

Dim lngIndex As Long

   If IsNumeric(City) Then
      lngIndex = City
      
   ElseIf VarType(City) = vbString Then
      lngIndex = CityIndex(City)
      
   Else
      lngIndex = -1
   End If
   
   m_DaylightSavingTime = False
   
   If (lngIndex < 0) Or (lngIndex >= m_CityCount) Then
      m_Latitude = 0
      m_Longitude = 0
      m_TimeZone = 0
      m_Hemisphere = North
      
   Else
      With CityInfo(lngIndex)
         m_Latitude = .Latitude
         m_Longitude = .Longitude
         m_TimeZone = .TimeZone
         m_Hemisphere = CalcHemisphere
      End With
      
      CitySet = True
   End If

End Function

Public Function Hemisphere() As Hemispheres

   Hemisphere = m_Hemisphere

End Function

' Validates a value for the decimal sign
Public Function ValidateValue(ByVal Value As String) As String

   If IsNumeric(Value) Then
      If Len(CStr(CDbl(Value))) <> Len(Value) Then
         If InStr(Value, ".") Then
            ValidateValue = Replace(Value, ".", ",")
            
         ElseIf InStr(Value, ",") Then
            ValidateValue = Replace(Value, ",", ".")
            
         Else
            ValidateValue = ERROR_VALUE
         End If
         
      Else
         ValidateValue = Value
      End If
      
   Else
      ValidateValue = ERROR_VALUE
   End If

End Function

Public Function MoonDescription() As MoonDescriptions

Dim dtePhase(2) As Date
Dim intPhase    As Integer

   If Not LimitsForDate Then Exit Function
   
   For intPhase = 0 To 2
      dtePhase(intPhase) = CalcJulianDayToTimeUTC(CalcMoonPhase(intPhase + 1, CalcLunation(m_DateCalculate)))
   Next 'intPhase
   
   If m_DateCalculate >= dtePhase(2) Then
      MoonDescription = WaningCrescent
      
   ElseIf m_DateCalculate >= dtePhase(1) Then
      MoonDescription = WaningGibbous
      
   ElseIf m_DateCalculate >= dtePhase(0) Then
      MoonDescription = WaxingGibbous
      
   Else
      MoonDescription = WaxingCrescent
   End If
   
   Erase dtePhase

End Function

Public Sub MoonCalculate()

Dim blnAbove       As Boolean
Dim blnMoonrise    As Boolean
Dim blnMoonset     As Boolean
Dim dteMoonrise    As Date
Dim dteMoonset     As Date
Dim dteMoontransit As Date
Dim dblCosLatitude As Double
Dim dblDepression  As Double
Dim dblJuliandDay  As Double
Dim dblLongitude   As Double
Dim dblMoonAge     As Double
Dim dblMoonrise    As Double
Dim dblMoonset     As Double
Dim dblSinLatitude As Double
Dim dblY(3)        As Double
Dim dblZero(2)     As Double
Dim intHour        As Integer
Dim mstMoonrise    As MoonStatus
Dim mstMoonset     As MoonStatus
Dim mstMoonStatus  As MoonStatus

   dblLongitude = -m_Longitude
   dblJuliandDay = CalcJulianDayModified
   dblJuliandDay = CalcJulianDayModified - m_TimeZone / 24
   dblSinLatitude = Sin(m_Latitude * RadianAngle)
   dblCosLatitude = Cos(m_Latitude * RadianAngle)
   dblDepression = Sin((8 / 60) * RadianAngle)
   dblY(0) = CalcSinAltitude(dblJuliandDay, 0, dblLongitude, dblCosLatitude, dblSinLatitude) - dblDepression
   blnAbove = (dblY(0) > 0)
   intHour = 1
   
   Do
      dblY(1) = CalcSinAltitude(dblJuliandDay, intHour, dblLongitude, dblCosLatitude, dblSinLatitude) - dblDepression
      dblY(2) = CalcSinAltitude(dblJuliandDay, intHour + 1, dblLongitude, dblCosLatitude, dblSinLatitude) - dblDepression
      dblY(3) = 0
      dblZero(0) = 0
      dblZero(1) = 0
      dblZero(2) = 0
      
      Call CalcQuadratic(dblY(), dblZero())
      
      If dblZero(0) = 1 Then
         If dblY(0) < 0 Then
            dblMoonrise = intHour + dblZero(1)
            blnMoonrise = True
            
         Else
            dblMoonset = intHour + dblZero(1)
            blnMoonset = True
         End If
         
      ElseIf dblZero(0) = 2 Then
         If dblY(3) < 0 Then
            dblMoonrise = intHour + dblZero(2)
            dblMoonset = intHour + dblZero(1)
            
         Else
            dblMoonrise = intHour + dblZero(1)
            dblMoonset = intHour + dblZero(2)
         End If
         
         blnMoonrise = True
         blnMoonset = True
      End If
      
      dblY(0) = dblY(2)
      intHour = intHour + 2
   Loop Until (intHour = 25) Or (blnMoonrise And blnMoonset)
   
   If blnMoonrise Or blnMoonset Then
      If blnMoonrise Then dteMoonrise = MakeTime(dblMoonrise - m_TimeZone)
      If blnMoonset Then dteMoonset = MakeTime(dblMoonset - m_TimeZone)
      
      If Not blnMoonrise Then
         mstMoonrise = None
         dteMoonrise = TimeValue("00:00:00")
         
         If Abs(DateDiff("n", TimeValue(dteMoonset), Time)) < 10 Then
            mstMoonStatus = Setting
            
         ElseIf Time > TimeValue(dteMoonset) Then
            mstMoonStatus = Down
            
         Else
            mstMoonStatus = Up
         End If
         
      ElseIf Not blnMoonset Then
         mstMoonset = None
         dteMoonset = TimeValue("00:00:00")
         
         If Abs(DateDiff("n", TimeValue(dteMoonrise), Time)) < 10 Then
            mstMoonStatus = Rising
            
         ElseIf Time > TimeValue(dteMoonrise) Then
            mstMoonStatus = Up
            
         Else
            mstMoonStatus = Down
         End If
         
      ElseIf Abs(DateDiff("n", TimeValue(dteMoonrise), Time)) < 10 Then
         mstMoonStatus = Rising
         
      ElseIf Abs(DateDiff("n", TimeValue(dteMoonset), Time)) < 10 Then
         mstMoonStatus = Setting
         
      ElseIf (Time > TimeValue(dteMoonrise)) And (Time < TimeValue(dteMoonset)) Then
         mstMoonStatus = Up
         
      ElseIf ((Time > TimeValue(dteMoonset)) And (Time < TimeValue(dteMoonrise))) Or (TimeValue(dteMoonrise) < TimeValue(dteMoonset)) Then
         mstMoonStatus = Down
         
      Else
         mstMoonStatus = Up
      End If
      
   ElseIf blnAbove Then
      dteMoonrise = TimeValue("00:00:00")
      mstMoonrise = AboveHorizon
      mstMoonStatus = UpAllDay
      
   Else
      dteMoonrise = TimeValue("00:00:00")
      mstMoonrise = BelowHorizon
      mstMoonStatus = DownAllDay
   End If
   
   dteMoontransit = MakeTime(CalcMoonTransit)
   dblMoonAge = CalcMoonAge(m_DateCalculate, False)
   RaiseEvent ResultMoonCalculate(m_DateCalculate, dteMoonrise, dteMoontransit, dteMoonset, mstMoonrise, mstMoonset, mstMoonStatus, dblMoonAge)
   Erase dblY, dblZero

End Sub

Public Sub MoonPhases(ByVal PreviousPhases As Boolean)

Dim dtePhase(3) As Date
Dim intPhase    As Integer

   ' 0 = NewMoon
   ' 1 = FirstQuarter
   ' 2 = FullMoon
   ' 3 = LastQuarter
   For intPhase = 0 To 3
      dtePhase(intPhase) = CalcJulianDayToTimeUTC(CalcMoonPhase(intPhase, CalcLunation(m_DateCalculate)))
      
      If PreviousPhases Then
         If dtePhase(intPhase) >= m_DateCalculate Then dtePhase(intPhase) = CalcJulianDayToTimeUTC(CalcMoonPhase(intPhase, CalcLunation(m_DateCalculate) - 1))
         
      ElseIf dtePhase(intPhase) <= m_DateCalculate Then
         dtePhase(intPhase) = CalcJulianDayToTimeUTC(CalcMoonPhase(intPhase, CalcLunation(m_DateCalculate) + 1))
      End If
      
      If Not m_TimeInUTC Then dtePhase(intPhase) = DateAdd("s", m_TimeZone * 3600, dtePhase(intPhase))
   Next 'intPhase
   
   RaiseEvent ResultMoonPhases(m_DateCalculate, dtePhase(0), dtePhase(1), dtePhase(2), dtePhase(3), PreviousPhases)
   Erase dtePhase

End Sub

Public Sub MoonPosition(ByVal IsDate As Date)

Dim dblMoonAge          As Double
Dim dblMoonAngle        As Double
Dim dblMoonIllumination As Double

   dblMoonAge = CalcMoonAge(IsDate, True)
   dblMoonAngle = dblMoonAge * 360 / MOON_SYNODIC_PERIOD
   dblMoonIllumination = (1 - Cos((PI_DOUBLE * dblMoonAge) / MOON_SYNODIC_PERIOD)) / 2 * 100
   RaiseEvent ResultMoonPosition(IsDate, dblMoonIllumination, dblMoonAngle, dblMoonAge)

End Sub

Public Sub SunCalculate()

Dim dteSunrise         As Date
Dim dteSunset          As Date
Dim dteSunTime         As Date
Dim dteSuntransit      As Date
Dim dteTwilights(2, 1) As Date
Dim dblTime            As Double
Dim intSunDepression   As Integer
Dim intTwilight        As Integer
Dim lngSunHours        As Long
Dim lngSunMinutes      As Long
Dim lngSunSeconds      As Long

   ' Calculate Sunrise
   dblTime = CalcSunState(False)
   
   If dblTime = ERROR_VALUE Then
      dteSunrise = CalcForNoSunState(False)
      
   Else
      dteSunrise = MakeDate(dblTime)
   End If
   
   ' Calculate Suntransit
   dblTime = CalcSunState(False, True)
   dteSuntransit = MakeDate(dblTime)
   
   ' Calculate Sunset
   dblTime = CalcSunState(True)
   
   If dblTime = ERROR_VALUE Then
      dteSunset = CalcForNoSunState(True)
      
   Else
      dteSunset = MakeDate(dblTime)
   End If
   
   ' Calculate SunTime
   dblTime = DateDiff("s", dteSunrise, dteSunset)
   lngSunHours = dblTime \ 3600
   lngSunMinutes = dblTime \ 60 - lngSunHours * 60
   lngSunSeconds = dblTime - lngSunHours * 3600 - lngSunMinutes * 60
   dteSunTime = TimeSerial(lngSunHours, lngSunMinutes, lngSunSeconds)
   
   ' Calculate Twilights
   ' 0 - Civil Twilight        =  6
   ' 1 - Nautical Twilight     = 12
   ' 2 - Astronomical Twilight = 18
   For intTwilight = 0 To 2
      intSunDepression = 6 + intTwilight * 6
      dblTime = CalcTwilightState(intSunDepression, False)
      
      If dblTime = ERROR_VALUE Then
         dteTwilights(intTwilight, 0) = DateSerial(Year(dteSuntransit), Month(dteSuntransit), Day(dteSuntransit))
         
      Else
         dteTwilights(intTwilight, 0) = MakeDate(dblTime)
      End If
   
      dblTime = CalcTwilightState(intSunDepression, True)
      
      If dblTime = ERROR_VALUE Then
         dteTwilights(intTwilight, 1) = DateSerial(Year(dteSuntransit), Month(dteSuntransit), Day(dteSuntransit))
         
      Else
         dteTwilights(intTwilight, 1) = MakeDate(dblTime)
      End If
   Next 'intTwilight
   
   RaiseEvent ResultSunCalculate(m_DateCalculate, dteSunrise, dteSuntransit, dteSunset, dteSunTime, dteTwilights(0, 0), dteTwilights(0, 1), dteTwilights(1, 0), dteTwilights(1, 1), dteTwilights(2, 0), dteTwilights(2, 1), EquationOfTime, SunDeclination)
   Erase dteTwilights

End Sub

Public Sub SunPosition(ByVal IsDate As Date)

Dim dteDateCalculate        As Date
Dim dblAzimuth              As Double
Dim dblElevation            As Double
Dim dblHourAngle            As Double
Dim dblHours                As Double
Dim dblJulianDay            As Double
Dim dblLatitude             As Double
Dim dblLongitude            As Double
Dim dblMinutes              As Double
Dim dblRefractionCorrection As Double
Dim dblResult               As Double
Dim dblSeconds              As Double
Dim dblSunDeclination       As Double
Dim dblTime                 As Double
Dim dblTimeZone             As Double
Dim dblTrueSunTime          As Double
Dim dblZenith               As Double

   dblLatitude = CalcLatitude
   dblLongitude = -m_Longitude
   dblTimeZone = CalcTimeZone
   dteDateCalculate = m_DateCalculate
   m_DateCalculate = DateSerial(Year(IsDate), Month(IsDate), Day(IsDate))
   dblHours = Hour(IsDate) - CalcDaylightSavingTime / 60
   dblMinutes = Minute(IsDate)
   dblSeconds = Second(IsDate)
   dblJulianDay = CalcJulianDay
   dblTime = dblHours + dblMinutes / 60 + dblSeconds / 3600 - dblTimeZone
   dblTime = CalcTimeJulianCenturies(dblJulianDay + dblTime / 24)
   dblSunDeclination = CalcSunDeclination(dblTime)
   dblTrueSunTime = dblHours * 60 + dblMinutes + dblSeconds / 60 + (CalcEquationOfTime(dblTime) - 4 * dblLongitude - dblTimeZone * 60)
   
   Do While dblTrueSunTime > 1440
      dblTrueSunTime = dblTrueSunTime - 1440
   Loop
   
   dblHourAngle = dblTrueSunTime / 4 - 180
   
   If dblHourAngle < -180 Then dblHourAngle = dblHourAngle + 360
   
   dblResult = Sin(DegreesToRadians(dblLatitude)) * Sin(DegreesToRadians(dblSunDeclination)) + Cos(DegreesToRadians(dblLatitude)) * Cos(DegreesToRadians(dblSunDeclination)) * Cos(DegreesToRadians(dblHourAngle))
   
   If Abs(dblResult) > 1 Then dblResult = Sgn(dblResult)
   
   dblZenith = RadiansToDegrees(Acos(dblResult))
   dblResult = (Cos(DegreesToRadians(dblLatitude)) * Sin(DegreesToRadians(dblZenith)))
   
   If Abs(dblResult) > 0.001 Then
      dblResult = ((Sin(DegreesToRadians(dblLatitude)) * Cos(DegreesToRadians(dblZenith))) - Sin(DegreesToRadians(dblSunDeclination))) / dblResult
      
      If Abs(dblResult) > 1 Then dblResult = Sgn(dblResult)
      
      dblAzimuth = 180 - RadiansToDegrees(Acos(dblResult))
      
      If dblHourAngle > 0 Then dblAzimuth = -dblAzimuth
      
   ElseIf dblLatitude > 0 Then
      dblAzimuth = 180
      
   Else
      dblAzimuth = 0
   End If
   
   If dblAzimuth < 0 Then dblAzimuth = dblAzimuth + 360
   
   dblElevation = 90 - dblZenith
   
   If dblElevation > 85 Then
      dblRefractionCorrection = 0
      
   Else
      dblResult = Tan(DegreesToRadians(dblElevation))
      
      If dblElevation > 5 Then
         dblRefractionCorrection = 58.1 / dblResult - 0.07 / (dblResult ^ 3) + 0.000086 / (dblResult ^ 5)
         
      ElseIf dblElevation > -0.575 Then
         dblRefractionCorrection = 103.4 + dblElevation * (-12.79 + dblElevation * 0.711)
         dblRefractionCorrection = 1735 + dblElevation * (-518.2 + dblElevation * dblRefractionCorrection)
         
      Else
         dblRefractionCorrection = -20.774 / dblResult
      End If
      
      dblRefractionCorrection = dblRefractionCorrection / 3600
   End If
   
   dblZenith = dblZenith - dblRefractionCorrection
   m_DateCalculate = dteDateCalculate
   RaiseEvent ResultSunPosition(IsDate, dblAzimuth, dblZenith, 90 - dblZenith)

End Sub

Private Function Acos(ByVal Value As Double) As Double

   On Local Error GoTo ErrorAcos
   
   If Abs(Value) <> 1 Then
      Acos = PI / 2 - Atn(Value / Sqr(1 - Value * Value))
      
   ElseIf Value = -1 Then
      Acos = PI
      
   ElseIf Value = 1 Then
      Acos = 0
   End If
   
   GoTo ExitFunction
   
ErrorAcos:
   Acos = ERROR_VALUE
   
ExitFunction:
   On Local Error GoTo 0

End Function

Private Function Asin(ByVal Value As Double) As Double

   If Abs(Value) <> 1 Then
      Asin = Atn(Value / Sqr(1 - Value * Value))
      
   Else
      Asin = PI / 2 * Sgn(Value)
   End If

End Function

Private Function Atan2(ByVal Y As Double, ByVal X As Double) As Double

   If Y > 0 Then
      If X >= Y Then
         Atan2 = Atn(Y / X)
         
      ElseIf X <= -Y Then
         Atan2 = Atn(Y / X) + PI
         
      Else
         Atan2 = PI / 2 - Atn(X / Y)
      End If
      
   Else
      If X >= -Y Then
         Atan2 = Atn(Y / X)
         
      ElseIf X <= Y Then
         Atan2 = Atn(Y / X) - PI
         
      Else
         Atan2 = -Atn(X / Y) - PI / 2
      End If
   End If

End Function

Private Function CalcDaylightSavingTime() As Double

   CalcDaylightSavingTime = Abs(m_DaylightSavingTime) * 60

End Function

Private Function CalcEccentricityEarthOrbit(ByVal Time As Double) As Double

   CalcEccentricityEarthOrbit = 0.016708634 - Time * (0.000042037 + Time * 0.0000001267)

End Function

Private Function CalcEphemeris(ByVal Phase As Integer, ByVal Ephemeris As Double, ByVal EccentricityAnomaly As Double, ByVal MoonLatitude As Double, ByVal MoonLongitude As Double, ByVal MeanAnomalyMoon As Double, ByVal MeanAnomalySun As Double, ByRef Arguments() As Double) As Double

Dim dblCorrection As Double
Dim dblEphemeris  As Double

   On Local Error GoTo ErrorCalulation
   dblEphemeris = Ephemeris
   
   ' New Moon
   If Phase = 0 Then
      dblEphemeris = dblEphemeris - 0.4072 * Sin(MeanAnomalyMoon)
      dblEphemeris = dblEphemeris + 0.17241 * EccentricityAnomaly * Sin(MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.01608 * Sin(2 * MeanAnomalyMoon)
      dblEphemeris = dblEphemeris + 0.01039 * Sin(2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00739 * EccentricityAnomaly * Sin(MeanAnomalyMoon - MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.00514 * EccentricityAnomaly * Sin(MeanAnomalyMoon + MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00208 * EccentricityAnomaly * EccentricityAnomaly * Sin(2 * MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.00111 * Sin(MeanAnomalyMoon - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00057 * Sin(MeanAnomalyMoon + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00056 * EccentricityAnomaly * Sin(2 * MeanAnomalyMoon + MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.00042 * Sin(3 * MeanAnomalyMoon)
      dblEphemeris = dblEphemeris + 0.00042 * EccentricityAnomaly * Sin(MeanAnomalySun + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00038 * EccentricityAnomaly * Sin(MeanAnomalySun - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00024 * EccentricityAnomaly * Sin(2 * MeanAnomalyMoon - MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.00017 * Sin(MoonLongitude)
      dblEphemeris = dblEphemeris - 0.00007 * Sin(MeanAnomalyMoon + 2 * MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00004 * Sin(2 * MeanAnomalyMoon - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00004 * Sin(3 * MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00003 * Sin(MeanAnomalyMoon + MeanAnomalySun - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00003 * Sin(2 * MeanAnomalyMoon + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00003 * Sin(MeanAnomalyMoon + MeanAnomalySun + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00003 * Sin(MeanAnomalyMoon - MeanAnomalySun + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00002 * Sin(MeanAnomalyMoon - MeanAnomalySun - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00002 * Sin(3 * MeanAnomalyMoon + MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00002 * Sin(4 * MeanAnomalyMoon)
      
   ' Full Moon
   ElseIf Phase = 2 Then
      dblEphemeris = dblEphemeris - 0.40614 * Sin(MeanAnomalyMoon)
      dblEphemeris = dblEphemeris + 0.17302 * EccentricityAnomaly * Sin(MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.01614 * Sin(2 * MeanAnomalyMoon)
      dblEphemeris = dblEphemeris + 0.01043 * Sin(2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00734 * EccentricityAnomaly * Sin(MeanAnomalyMoon - MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.00515 * EccentricityAnomaly * Sin(MeanAnomalyMoon + MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00209 * EccentricityAnomaly * EccentricityAnomaly * Sin(2 * MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.00111 * Sin(MeanAnomalyMoon - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00057 * Sin(MeanAnomalyMoon + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00056 * EccentricityAnomaly * Sin(2 * MeanAnomalyMoon + MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.00042 * Sin(3 * MeanAnomalyMoon)
      dblEphemeris = dblEphemeris + 0.00042 * EccentricityAnomaly * Sin(MeanAnomalySun + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00038 * EccentricityAnomaly * Sin(MeanAnomalySun - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00024 * EccentricityAnomaly * Sin(2 * MeanAnomalyMoon - MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.00017 * Sin(MoonLongitude)
      dblEphemeris = dblEphemeris - 0.00007 * Sin(MeanAnomalyMoon + 2 * MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00004 * Sin(2 * MeanAnomalyMoon - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00004 * Sin(3 * MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00003 * Sin(MeanAnomalyMoon + MeanAnomalySun - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00003 * Sin(2 * MeanAnomalyMoon + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00003 * Sin(MeanAnomalyMoon + MeanAnomalySun + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00003 * Sin(MeanAnomalyMoon - MeanAnomalySun + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00002 * Sin(MeanAnomalyMoon - MeanAnomalySun - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00002 * Sin(3 * MeanAnomalyMoon + MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00002 * Sin(4 * MeanAnomalyMoon)
      
   ' First Quarter or Last Quarter (Phase = 1 or Phase = 3)
   Else
      dblEphemeris = dblEphemeris - 0.62801 * Sin(MeanAnomalyMoon)
      dblEphemeris = dblEphemeris + 0.17172 * EccentricityAnomaly * Sin(MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.01183 * EccentricityAnomaly * Sin(MeanAnomalyMoon + MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00862 * Sin(2 * MeanAnomalyMoon)
      dblEphemeris = dblEphemeris + 0.00804 * Sin(2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00454 * EccentricityAnomaly * Sin(MeanAnomalyMoon - MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00204 * EccentricityAnomaly * EccentricityAnomaly * Sin(2 * MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.0018 * Sin(MeanAnomalyMoon - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.0007 * Sin(MeanAnomalyMoon + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.0004 * Sin(3 * MeanAnomalyMoon)
      dblEphemeris = dblEphemeris - 0.00034 * EccentricityAnomaly * Sin(2 * MeanAnomalyMoon - MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00032 * EccentricityAnomaly * Sin(MeanAnomalySun + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00032 * EccentricityAnomaly * Sin(MeanAnomalySun - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00028 * EccentricityAnomaly * EccentricityAnomaly * Sin(MeanAnomalyMoon + 2 * MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00027 * EccentricityAnomaly * Sin(2 * MeanAnomalyMoon + MeanAnomalySun)
      dblEphemeris = dblEphemeris - 0.00017 * Sin(MoonLongitude)
      dblEphemeris = dblEphemeris - 0.00005 * Sin(MeanAnomalyMoon - MeanAnomalySun - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00004 * Sin(2 * MeanAnomalyMoon + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00004 * Sin(MeanAnomalyMoon + MeanAnomalySun + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00004 * Sin(MeanAnomalyMoon - 2 * MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00003 * Sin(MeanAnomalyMoon + MeanAnomalySun - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00003 * Sin(3 * MeanAnomalySun)
      dblEphemeris = dblEphemeris + 0.00002 * Sin(2 * MeanAnomalyMoon - 2 * MoonLatitude)
      dblEphemeris = dblEphemeris + 0.00002 * Sin(MeanAnomalyMoon - MeanAnomalySun + 2 * MoonLatitude)
      dblEphemeris = dblEphemeris - 0.00002 * Sin(3 * MeanAnomalyMoon + MeanAnomalySun)
      dblCorrection = 0.00306
      dblCorrection = dblCorrection - 0.00038 * EccentricityAnomaly * Cos(MeanAnomalySun)
      dblCorrection = dblCorrection + 0.00026 * Cos(MeanAnomalyMoon)
      dblCorrection = dblCorrection - 0.00002 * Cos(MeanAnomalyMoon - MeanAnomalySun)
      dblCorrection = dblCorrection + 0.00002 * Cos(MeanAnomalyMoon + MeanAnomalySun)
      dblCorrection = dblCorrection + 0.00002 * Cos(2 * MoonLatitude)
      
      ' Last Quarter
      If Phase = 3 Then dblCorrection = -dblCorrection
      
      dblEphemeris = dblEphemeris + dblCorrection
   End If
   
   dblEphemeris = dblEphemeris + 0.000325 * Sin(Arguments(0))
   dblEphemeris = dblEphemeris + 0.000165 * Sin(Arguments(1))
   dblEphemeris = dblEphemeris + 0.000164 * Sin(Arguments(2))
   dblEphemeris = dblEphemeris + 0.000126 * Sin(Arguments(3))
   dblEphemeris = dblEphemeris + 0.00011 * Sin(Arguments(4))
   dblEphemeris = dblEphemeris + 0.000062 * Sin(Arguments(5))
   dblEphemeris = dblEphemeris + 0.00006 * Sin(Arguments(6))
   dblEphemeris = dblEphemeris + 0.000056 * Sin(Arguments(7))
   dblEphemeris = dblEphemeris + 0.000047 * Sin(Arguments(8))
   dblEphemeris = dblEphemeris + 0.000042 * Sin(Arguments(9))
   dblEphemeris = dblEphemeris + 0.00004 * Sin(Arguments(10))
   dblEphemeris = dblEphemeris + 0.000037 * Sin(Arguments(11))
   dblEphemeris = dblEphemeris + 0.000035 * Sin(Arguments(12))
   CalcEphemeris = dblEphemeris + 0.000023 * Sin(Arguments(13))
   
ErrorCalulation:
   On Local Error GoTo 0

End Function

Private Function CalcEquationOfTime(ByVal Time As Double) As Double

Dim dblEccentricity  As Double
Dim dblEpsilon       As Double
Dim dblMeanAnomaly   As Double
Dim dblMeanLongitude As Double
Dim dblResult(5)     As Double

   dblEpsilon = CalcObliquityCorrection(Time)
   dblEccentricity = CalcEccentricityEarthOrbit(Time)
   dblMeanLongitude = CalcGeoMeanLongitudeSun(Time)
   dblMeanAnomaly = CalcGeoMeanAnomalySun(Time)
   dblResult(0) = Tan(DegreesToRadians(dblEpsilon) / 2) ^ 2
   dblResult(1) = Sin(DegreesToRadians(dblMeanLongitude) * 2)
   dblResult(2) = Sin(DegreesToRadians(dblMeanAnomaly))
   dblResult(3) = Cos(DegreesToRadians(dblMeanLongitude) * 2)
   dblResult(4) = Sin(DegreesToRadians(dblMeanLongitude) * 4)
   dblResult(5) = Sin(DegreesToRadians(dblMeanAnomaly) * 2)
   dblEccentricity = dblResult(0) * dblResult(1) - 2 * dblEccentricity * dblResult(2) + 4 * dblEccentricity * dblResult(0) * dblResult(2) * dblResult(3) - 0.5 * dblResult(0) * dblResult(0) * dblResult(4) - 1.25 * dblEccentricity * dblEccentricity * dblResult(5)
   CalcEquationOfTime = RadiansToDegrees(dblEccentricity) * 4
   Erase dblResult

End Function

Private Function CalcForNoSunState(ByVal Sunset As Boolean) As Double

Dim dblTimeUTC As Double

   dblTimeUTC = CalcForNoSunStateUTC(Sunset)
   
   If (dblTimeUTC <> ERROR_VALUE) And Not m_TimeInUTC Then dblTimeUTC = dblTimeUTC + CalcTimeZoneAsTime
   
   CalcForNoSunState = dblTimeUTC

End Function

Private Function CalcForNoSunStateUTC(ByVal Sunset As Boolean) As Double

Dim dblDayOfYear As Double
Dim dblJulianDay As Double
Dim dblLatitude  As Double
Dim dblLongitude As Double
Dim dblTimeUTC   As Double

   dblLatitude = CalcLatitude
   dblLongitude = -m_Longitude
   dblJulianDay = CalcJulianDay
   dblDayOfYear = DatePart("y", DateSerial(Year(m_DateCalculate), Month(m_DateCalculate), Day(m_DateCalculate)))
   
   If ((dblLatitude > 66.4) And (dblDayOfYear > 79) And (dblDayOfYear < 267)) Or ((dblLatitude < -66.4) And ((dblDayOfYear < 83) Or (dblDayOfYear > 263))) Then
      dblJulianDay = FindSunState(dblJulianDay, Not Sunset, Sunset)
      
   ElseIf ((dblLatitude > 66.4) And ((dblDayOfYear < 83) Or (dblDayOfYear > 263))) Or ((dblLatitude < -66.4) And (dblDayOfYear > 79) And (dblDayOfYear < 267)) Then
      dblJulianDay = FindSunState(dblJulianDay, Not Sunset, Not Sunset)
   End If
   
   dblTimeUTC = CalcSunStateUTC(dblJulianDay, Sunset) * 1440
   
   If dblTimeUTC > 1440 Then
      dblTimeUTC = dblTimeUTC - 1440
      dblJulianDay = dblJulianDay + 1
   End If
   
   If dblTimeUTC < 0 Then
      dblTimeUTC = dblTimeUTC + 1440
      dblJulianDay = dblJulianDay - 1
   End If
   
   CalcForNoSunStateUTC = (dblJulianDay - CalcJulianDay) / 1440

End Function

Private Function CalcGeoMeanAnomalySun(ByVal Time As Double) As Double

   CalcGeoMeanAnomalySun = 357.52911 + Time * (35999.05029 - Time * 0.0001537)

End Function

Private Function CalcGeoMeanLongitudeSun(ByVal Time As Double) As Double

Dim dblMeanLongitude As Double

   dblMeanLongitude = 280.46646 + Time * (36000.76983 + Time * 0.0003032)
   
   Do While dblMeanLongitude > 360
      dblMeanLongitude = dblMeanLongitude - 360
   Loop
   
   Do While dblMeanLongitude < 0
      dblMeanLongitude = dblMeanLongitude + 360
   Loop
   
   CalcGeoMeanLongitudeSun = dblMeanLongitude

End Function

Private Function CalcHemisphere() As Hemispheres

Dim hspNorthSouth As Hemispheres

   If Sgn(m_Latitude) = 1 Then
      hspNorthSouth = North
      
   Else
      hspNorthSouth = South
   End If
   
   CalcHemisphere = hspNorthSouth

End Function

Private Function CalcHourAngleSun(ByVal Latitude As Double, ByVal SunDeclination As Double, ByVal Sunset As Boolean, Optional ByVal SunDepression As Double = 0.833333333333333) As Double

Dim dblHourAngle As Double
Dim dblResult(1) As Double

   dblResult(0) = DegreesToRadians(Latitude)
   dblResult(1) = DegreesToRadians(SunDeclination)
   dblHourAngle = Cos(DegreesToRadians(90 + SunDepression)) / (Cos(dblResult(0)) * Cos(dblResult(1))) - Tan(dblResult(0)) * Tan(dblResult(1))
   
   If Abs(dblHourAngle) > 1 Then
      dblHourAngle = ERROR_VALUE
      
   Else
      dblHourAngle = Acos(dblHourAngle)
   End If
   
   If Sunset And (dblHourAngle <> ERROR_VALUE) Then dblHourAngle = -dblHourAngle
   
   CalcHourAngleSun = dblHourAngle
   Erase dblResult

End Function

Private Function CalcJulianDay() As Double

Dim dblCentury As Double
Dim intMonth   As Integer
Dim intYear    As Integer

   intYear = Year(m_DateCalculate)
   intMonth = Month(m_DateCalculate)
   
   If intMonth <= 2 Then
      intYear = intYear - 1
      intMonth = intMonth + 12
   End If
   
   dblCentury = intYear / 100
   dblCentury = Fix(2 - dblCentury + (dblCentury / 4))
   CalcJulianDay = Fix((365.25 * (intYear + 4716))) + Fix((30.6001 * (intMonth + 1))) + Day(m_DateCalculate) + dblCentury - 1524.5

End Function

Private Function CalcJulianDayFromJulianCenturies(ByVal Time As Double) As Double

   CalcJulianDayFromJulianCenturies = 2451545 + Time * 36525

End Function

Private Function CalcJulianDayModified() As Double

Dim dblResult As Double
Dim lngDay    As Long
Dim lngMonth  As Long
Dim lngYear   As Long

   lngYear = Year(m_DateCalculate)
   lngMonth = Month(m_DateCalculate)
   lngDay = Day(m_DateCalculate)
   dblResult = lngYear * 10000 + lngMonth * 100 + lngDay
   
   If lngYear < 0 Then lngYear = lngYear + 1
   
   If lngMonth <= 2 Then
      lngMonth = lngMonth + 12
      lngYear = lngYear - 1
   End If
   
   If dblResult <= 15821004.1 Then
      dblResult = -2 + Int((lngYear + 4716) / 4) - 1179
      
   Else
      dblResult = Int(lngYear / 400) - Int(lngYear / 100) + Int(lngYear / 4)
   End If
   
   CalcJulianDayModified = (lngYear * 365 - 679004) + dblResult + Sgn((lngMonth + 1) * 30.6001) * Int(Abs((lngMonth + 1) * 30.6001)) + lngDay

End Function

Private Function CalcJulianDayToTimeUTC(ByVal JulianDay As Double) As Date

Dim dblDate   As Double
Dim dblResult As Double

   If Not LimitsForJulianDay(JulianDay) Then Exit Function
   
   If JulianDay >= 2299160.5 Then
      dblDate = JulianDay - 2415018.5
      
   Else
      dblDate = JulianDay - 2415028.5
   End If
   
   ' Dates from 100 A.D. to 1900
   If JulianDay < 2415020.5 Then
      dblResult = dblDate - Fix(dblDate)
      
      If dblResult = 0 Then
         dblDate = Fix(dblDate)
         
      Else
         dblDate = Fix(dblDate - 3) + 1 - dblResult
      End If
   End If
   
   CalcJulianDayToTimeUTC = CDate(dblDate)

End Function

Private Function CalcLatitude() As Double

Dim dblLatitude As Double

   dblLatitude = m_Latitude
   
   If dblLatitude < -89.8 Then dblLatitude = -89.8
   If dblLatitude > 89.8 Then dblLatitude = 89.8
   
   CalcLatitude = dblLatitude

End Function

Private Function CalcLocalSiderialTime(ByVal JulianDay As Double, ByVal Longitude As Double) As Double

Dim dblJulianDay        As Double
Dim dblMeanSiderialTime As Double
Dim dblTime             As Double
Dim dblTimeUTC          As Double

   dblJulianDay = Sgn(JulianDay) * Int(Abs(JulianDay))
   dblTimeUTC = (JulianDay - dblJulianDay) * 24
   dblTime = (dblJulianDay - 51544.5) / 36525
   dblMeanSiderialTime = 6.697374558 + dblTimeUTC * 1.0027379093
   dblMeanSiderialTime = dblMeanSiderialTime + (8640184.812866 + dblTime * (0.093104 - dblTime * 0.0000062)) * dblTime / 3600
   CalcLocalSiderialTime = FracPart((dblMeanSiderialTime - Longitude / 15) / 24) * 24

End Function

Private Function CalcLunation(ByVal IsDate As Date) As Long

Dim dblJulianDay  As Double
Dim dblLowerLimit As Double
Dim dblUpperLimit As Double
Dim lngLunation   As Long

   lngLunation = (CDbl(IsDate) - 8388.51399305556) / MOON_SYNODIC_PERIOD
   dblJulianDay = CalcTimeUTCToJulianDay(IsDate) + 0.0000057875
   
   Do
      DoEvents
      dblLowerLimit = CalcMoonPhase(0, lngLunation)
      dblUpperLimit = CalcMoonPhase(0, lngLunation + 1)
      
      If dblLowerLimit > dblJulianDay Then
         lngLunation = lngLunation - 1
         
      ElseIf dblUpperLimit < dblJulianDay Then
         lngLunation = lngLunation + 1
      End If
   Loop Until (dblLowerLimit <= dblJulianDay) And (dblUpperLimit > dblJulianDay)
   
   CalcLunation = lngLunation

End Function

Private Function CalcMeanObliquityOfEcliptic(ByVal Time As Double) As Double

Dim dblSeconds As Double

   dblSeconds = 21.448 - Time * (46.815 + Time * (0.00059 - Time * (0.001813)))
   CalcMeanObliquityOfEcliptic = 23 + (26 + (dblSeconds / 60)) / 60

End Function

Private Function CalcMoonAge(ByVal IsDate As Date, ByVal MoreAccurate As Boolean) As Double

Dim dblJulianDay As Double
Dim dblMoonAge   As Double

   If MoreAccurate Then
      dblMoonAge = CalcTimeUTCToJulianDay(IsDate) - CalcMoonPhase(0, CalcLunation(IsDate))
      
   Else
      dblJulianDay = (1 + CalcJulianDay - 2451550.1) / MOON_SYNODIC_PERIOD
      dblJulianDay = dblJulianDay - Int(dblJulianDay)
      
      If dblJulianDay < 0 Then dblJulianDay = dblJulianDay + 1
      
      dblMoonAge = dblJulianDay * MOON_SYNODIC_PERIOD
   End If
   
   CalcMoonAge = dblMoonAge

End Function

Private Function CalcMoonPhase(ByVal Phase As Integer, ByVal Lunation As Long) As Double

Dim dblArguments(13)       As Double
Dim dblEccentricityAnomaly As Double
Dim dblEphemeris           As Double
Dim dblFraction            As Double
Dim dblMeanAnomalyMoon     As Double
Dim dblMeanAnomalySun      As Double
Dim dblMoonLatitude        As Double
Dim dblMoonLongitude       As Double
Dim dblTime                As Double
Dim intArgument            As Integer

   dblFraction = Lunation - 953 + (Phase / 4)
   dblTime = dblFraction / 1236.85
   dblEphemeris = 2451550.09765 + (29.530588853 * dblFraction) + dblTime * dblTime * (0.0001337 + dblTime * (-0.00000015 + 0.00000000073 * dblTime))
   dblEccentricityAnomaly = 1 + dblTime * (-0.002516 + -0.0000074 * dblTime)
   dblMeanAnomalySun = 2.5534 + 29.10535669 * dblFraction + dblTime * dblTime * (-0.0000218 + -0.00000011 * dblTime)
   dblMeanAnomalyMoon = 201.5643 + 385.81693528 * dblFraction + dblTime * dblTime * (0.0107438 + dblTime * (0.00001239 + -0.000000058 * dblTime))
   dblMoonLatitude = 160.7108 + 390.67050274 * dblFraction + dblTime * dblTime * (-0.0016341 * dblTime * (-0.00000227 + 0.000000011 * dblTime))
   dblMoonLongitude = 124.7746 - 1.5637558 * dblFraction + dblTime * dblTime * (0.0020691 + 0.00000215 * dblTime)
   dblArguments(0) = 299.77 + 0.107408 * dblFraction - 0.009173 * dblTime * dblTime
   dblArguments(1) = 251.88 + 0.016321 * dblFraction
   dblArguments(2) = 251.83 + 26.651886 * dblFraction
   dblArguments(3) = 349.42 + 36.412478 * dblFraction
   dblArguments(4) = 84.66 + 18.206239 * dblFraction
   dblArguments(5) = 141.74 + 53.303771 * dblFraction
   dblArguments(6) = 207.14 + 2.453732 * dblFraction
   dblArguments(7) = 154.84 + 7.30686 * dblFraction
   dblArguments(8) = 34.52 + 27.261239 * dblFraction
   dblArguments(9) = 207.19 + 0.121824 * dblFraction
   dblArguments(10) = 291.34 + 1.844379 * dblFraction
   dblArguments(11) = 161.72 + 24.198154 * dblFraction
   dblArguments(12) = 239.56 + 25.513099 * dblFraction
   dblArguments(13) = 331.55 + 3.592518 * dblFraction
   dblMeanAnomalySun = Radians(dblMeanAnomalySun)
   dblMeanAnomalyMoon = Radians(dblMeanAnomalyMoon)
   dblMoonLatitude = Radians(dblMoonLatitude)
   dblMoonLongitude = Radians(dblMoonLongitude)
   
   For intArgument = 0 To 13
      dblArguments(intArgument) = Radians(dblArguments(intArgument))
   Next 'intArgument
   
   CalcMoonPhase = CalcEphemeris(Phase, dblEphemeris, dblEccentricityAnomaly, dblMoonLatitude, dblMoonLongitude, dblMeanAnomalyMoon, dblMeanAnomalySun, dblArguments())
   Erase dblArguments

End Function

Private Function CalcMoonTransit() As Double

Dim blnTimeUTC    As Boolean
Dim dteSuntransit As Date
Dim dblTime       As Double
Dim strTime       As String

   blnTimeUTC = m_TimeInUTC
   m_TimeInUTC = True
   dblTime = CalcSunState(False, True)
   dteSuntransit = MakeDate(dblTime)
   m_TimeInUTC = blnTimeUTC
   dblTime = (Fix(CalcMoonAge(m_DateCalculate, False)) * (1440 / MOON_SYNODIC_PERIOD))
   strTime = TimeValue(DateAdd("s", dblTime * 60, TimeValue(dteSuntransit)))
   CalcMoonTransit = Hour(strTime) + (Minute(strTime) * 60 + Second(strTime)) / 3600

End Function

Private Function CalcObliquityCorrection(ByVal Time As Double) As Double

   CalcObliquityCorrection = CalcMeanObliquityOfEcliptic(Time) + Cos(DegreesToRadians(CalcOmega(Time))) * 0.00256

End Function

Private Function CalcOmega(ByVal Time As Double) As Double

   CalcOmega = 125.04 - Time * 1934.136

End Function

Private Function CalcSinAltitude(ByVal JulianDay As Double, ByVal Hour As Integer, ByVal Longitude As Double, ByVal CosLatitude As Double, ByVal SinLatitude As Double) As Double

Const ARC                   As Double = 206264.8062
Const EPSILON_COS           As Double = 0.91748
Const EPSILON_SIN           As Double = 0.39778

Dim dblArguments(3)         As Double
Dim dblDeclination          As Double
Dim dblDifference           As Double
Dim dblHourAngle            As Double
Dim dblInstant              As Double
Dim dblMeanAnomalyMoon      As Double
Dim dblMeanAnomalySun       As Double
Dim dblMeanArgumentLatitude As Double
Dim dblMeanLongitudeMoon    As Double
Dim dblLatitudeMoon         As Double
Dim dblLongitude            As Double
Dim dblLongitudeMoon        As Double
Dim dblRightAscension       As Double
Dim dblRho                  As Double
Dim dblTime                 As Double

   dblInstant = JulianDay + Hour / 24
   dblTime = (dblInstant - 51544.5) / 36525
   dblMeanLongitudeMoon = FracPart(0.606433 + dblTime * 1336.855225)
   dblMeanAnomalyMoon = PI_DOUBLE * FracPart(0.374897 + dblTime * 1325.55241)
   dblMeanAnomalySun = PI_DOUBLE * FracPart(0.993133 + dblTime * 99.997361)
   dblDifference = PI_DOUBLE * FracPart(0.827361 + dblTime * 1236.853086)
   dblMeanArgumentLatitude = PI_DOUBLE * FracPart(0.259086 + dblTime * 1342.227825)
   dblLongitude = 22640 * Sin(dblMeanAnomalyMoon)
   dblLongitude = dblLongitude - 4586 * Sin(dblMeanAnomalyMoon - dblDifference * 2)
   dblLongitude = dblLongitude + 2370 * Sin(dblDifference * 2)
   dblLongitude = dblLongitude + 769 * Sin(dblMeanAnomalyMoon * 2)
   dblLongitude = dblLongitude - 668 * Sin(dblMeanAnomalySun)
   dblLongitude = dblLongitude - 412 * Sin(dblMeanArgumentLatitude * 2)
   dblLongitude = dblLongitude - 212 * Sin(dblMeanAnomalyMoon * 2 - dblDifference * 2)
   dblLongitude = dblLongitude - 206 * Sin(dblMeanAnomalyMoon + dblMeanAnomalySun - dblDifference * 2)
   dblLongitude = dblLongitude + 192 * Sin(dblMeanAnomalyMoon + dblDifference * 2)
   dblLongitude = dblLongitude - 165 * Sin(dblMeanAnomalySun - dblDifference * 2)
   dblLongitude = dblLongitude - 125 * Sin(dblDifference)
   dblLongitude = dblLongitude - 110 * Sin(dblMeanAnomalyMoon + dblMeanAnomalySun)
   dblLongitude = dblLongitude + 148 * Sin(dblMeanAnomalyMoon - dblMeanAnomalySun)
   dblLongitude = dblLongitude - 55 * Sin(dblMeanArgumentLatitude * 2 - dblDifference * 2)
   dblArguments(0) = dblMeanArgumentLatitude - dblDifference * 2
   dblArguments(1) = -526 * Sin(dblArguments(0))
   dblArguments(1) = dblArguments(1) + 44 * Sin(dblMeanAnomalyMoon + dblArguments(0))
   dblArguments(1) = dblArguments(1) - 31 * Sin(dblArguments(0) - dblMeanAnomalyMoon)
   dblArguments(1) = dblArguments(1) - 23 * Sin(dblMeanAnomalySun + dblArguments(0))
   dblArguments(1) = dblArguments(1) + 11 * Sin(dblArguments(0) - dblMeanAnomalySun)
   dblArguments(1) = dblArguments(1) - 25 * Sin(dblMeanArgumentLatitude - dblMeanAnomalyMoon * 2)
   dblArguments(1) = dblArguments(1) + 21 * Sin(dblMeanArgumentLatitude - dblMeanAnomalyMoon)
   dblArguments(0) = dblMeanArgumentLatitude + (dblLongitude + 412 * Sin(dblMeanArgumentLatitude * 2) + 541 * Sin(dblMeanAnomalySun)) / ARC
   dblLatitudeMoon = PI_DOUBLE * FracPart(dblMeanLongitudeMoon + dblLongitude / 1296000)
   dblLongitudeMoon = (18520 * Sin(dblArguments(0)) + dblArguments(1)) / ARC
   dblArguments(1) = Cos(dblLongitudeMoon) * Cos(dblLatitudeMoon)
   dblArguments(2) = Cos(dblLongitudeMoon) * Sin(dblLatitudeMoon)
   dblArguments(3) = EPSILON_COS * dblArguments(2) - EPSILON_SIN * Sin(dblLongitudeMoon)
   dblArguments(2) = EPSILON_SIN * dblArguments(2) + EPSILON_COS * Sin(dblLongitudeMoon)
   dblRho = Sqr(1 - dblArguments(2) * dblArguments(2))
   dblDeclination = (360 / PI_DOUBLE) * Atn(dblArguments(2) / dblRho)
   dblRightAscension = (48 / PI_DOUBLE) * Atn(dblArguments(3) / (dblArguments(1) + dblRho))
   
   If dblRightAscension < 0 Then dblRightAscension = dblRightAscension + 24
   
   dblHourAngle = 15 * (CalcLocalSiderialTime(dblInstant, Longitude) - dblRightAscension)
   CalcSinAltitude = SinLatitude * Sin(dblDeclination * RadianAngle) + CosLatitude * Cos(dblDeclination * RadianAngle) * Cos(dblHourAngle * RadianAngle)
   Erase dblArguments

End Function

Private Function CalcSunApparentLongitude(ByVal Time As Double) As Double

   CalcSunApparentLongitude = CalcSunTrueLongitude(Time) - 0.00569 - Sin(DegreesToRadians(CalcOmega(Time))) * 0.00478

End Function

Private Function CalcSunDeclination(ByVal Time As Double) As Double

Dim dblEpsilon As Double
Dim dblLambda  As Double
Dim dblTananum As Double

   dblEpsilon = CalcObliquityCorrection(Time)
   dblLambda = CalcSunApparentLongitude(Time)
   dblTananum = Sin(DegreesToRadians(dblEpsilon)) * Sin(DegreesToRadians(dblLambda))
   CalcSunDeclination = RadiansToDegrees(Asin(dblTananum))

End Function

Private Function CalcSunEquationOfCenter(ByVal Time As Double) As Double

Dim dblResult(3) As Double

   dblResult(0) = DegreesToRadians(CalcGeoMeanAnomalySun(Time))
   dblResult(1) = Sin(dblResult(0))
   dblResult(2) = Sin(dblResult(0) * 2)
   dblResult(3) = Sin(dblResult(0) * 3)
   CalcSunEquationOfCenter = dblResult(1) * (1.914602 - Time * (0.004817 + Time * 0.000014)) + dblResult(2) * (0.019993 - Time * 0.000101) + dblResult(3) * 0.000289
   Erase dblResult

End Function

Private Function CalcSunRadiusVector(ByVal Time As Double) As Double

Dim dblEccentricity As Double
Dim dblTrueAnomaly  As Double

   dblTrueAnomaly = CalcSunTrueAnomaly(Time)
   dblEccentricity = CalcEccentricityEarthOrbit(Time)
   CalcSunRadiusVector = (1.000001018 * (1 - dblEccentricity * dblEccentricity)) / (1 + dblEccentricity * Cos(DegreesToRadians(dblTrueAnomaly)))

End Function

Private Function CalcSunRightAscension(ByVal Time As Double) As Double

Dim dblEpsilon   As Double
Dim dblLambda    As Double
Dim dblTanadenom As Double
Dim dblTananum   As Double

   dblEpsilon = CalcObliquityCorrection(Time)
   dblLambda = CalcSunApparentLongitude(Time)
   dblTananum = Cos(DegreesToRadians(dblEpsilon)) * Sin(DegreesToRadians(dblLambda))
   dblTanadenom = Cos(DegreesToRadians(dblLambda))
   CalcSunRightAscension = RadiansToDegrees(Atan2(dblTanadenom, dblTananum))

End Function

Private Function CalcSunState(ByVal Sunset As Boolean, Optional ByVal Transit As Boolean) As Double

Dim dblTimeUTC As Double

   dblTimeUTC = CalcSunStateUTC(CalcJulianDay, Sunset, Transit)
   
   If (dblTimeUTC <> ERROR_VALUE) And Not m_TimeInUTC Then dblTimeUTC = dblTimeUTC + CalcTimeZoneAsTime
   
   CalcSunState = dblTimeUTC

End Function

Private Function CalcSunStateUTC(ByVal JulianDay As Double, ByVal Sunset As Boolean, Optional ByVal Transit As Boolean) As Double

Dim dblHourAngle As Double
Dim dblLatitude  As Double
Dim dblLongitude As Double
Dim dblTime      As Double
Dim dblTimeUTC   As Double

   dblLatitude = CalcLatitude
   dblLongitude = -m_Longitude
   dblTime = CalcTimeJulianCenturies(JulianDay)
   
   If Transit Then
      dblTime = CalcTimeJulianCenturies(CalcJulianDayFromJulianCenturies(dblTime) + 0.5 + dblLongitude / 360)
      dblTimeUTC = dblLongitude
      
   Else
      EquationOfTime = CalcEquationOfTime(dblTime)
      SunDeclination = CalcSunDeclination(dblTime)
      dblHourAngle = CalcHourAngleSun(dblLatitude, SunDeclination, Sunset)
      
      If dblHourAngle = ERROR_VALUE Then
         dblTimeUTC = ERROR_VALUE
         
      Else
         dblTimeUTC = (720 + (dblLongitude - RadiansToDegrees(dblHourAngle)) * 4 - EquationOfTime) / 1440
         dblTime = CalcTimeJulianCenturies(CalcJulianDayFromJulianCenturies(dblTime) + dblTimeUTC)
         dblHourAngle = CalcHourAngleSun(dblLatitude, CalcSunDeclination(dblTime), Sunset)
         
         If dblHourAngle = ERROR_VALUE Then
            dblTimeUTC = ERROR_VALUE
            
         Else
            dblTimeUTC = (dblLongitude - RadiansToDegrees(dblHourAngle))
         End If
      End If
   End If
   
   If dblTimeUTC <> ERROR_VALUE Then dblTimeUTC = (720 + dblTimeUTC * 4 - CalcEquationOfTime(dblTime)) / 1440
   
   CalcSunStateUTC = dblTimeUTC

End Function

Private Function CalcSunTrueAnomaly(ByVal Time As Double) As Double

   CalcSunTrueAnomaly = CalcGeoMeanAnomalySun(Time) + CalcSunEquationOfCenter(Time)

End Function

Private Function CalcSunTrueLongitude(ByVal Time As Double) As Double

   CalcSunTrueLongitude = CalcGeoMeanLongitudeSun(Time) + CalcSunEquationOfCenter(Time)

End Function

Private Function CalcTimeJulianCenturies(ByVal JulianDay As Double) As Double

   CalcTimeJulianCenturies = (JulianDay - 2451545) / 36525

End Function

Private Function CalcTimeUTCToJulianDay(ByVal IsDate As Date) As Double

Dim dblDate      As Double
Dim dblJulianDay As Double
Dim dblResult    As Double

   dblDate = CDbl(IsDate)
   
   If Sgn(dblDate) = -1 Then
      dblResult = dblDate - Fix(dblDate)
      
      If dblDate - dblResult < CDbl(DateSerial(1582, 10, 15)) Then
         dblJulianDay = Fix(dblDate) + 2415028
         
      Else
         dblJulianDay = Fix(dblDate) + 2415018
      End If
      
      dblJulianDay = dblJulianDay + 0.5 - dblResult
      
   Else
      dblJulianDay = dblDate + 2415018.5
   End If
   
   CalcTimeUTCToJulianDay = dblJulianDay

End Function

Private Function CalcTimeZone() As Double

Dim dblTimeZone As Double

   dblTimeZone = m_TimeZone
   
   If dblTimeZone > 12.5 Then
      dblTimeZone = dblTimeZone - 24
      
   ElseIf dblTimeZone < -12 Then
      dblTimeZone = dblTimeZone + 24
   End If
   
   CalcTimeZone = dblTimeZone

End Function

Private Function CalcTimeZoneAsTime() As Double

   CalcTimeZoneAsTime = (CalcTimeZone * 60 + CalcDaylightSavingTime) / 1440

End Function

Private Function CalcTwilightState(ByVal SunDepression As Double, ByVal Dusk As Boolean) As Double

Dim dblTimeUTC As Double

   dblTimeUTC = CalcTwilightStateUTC(SunDepression, Dusk)
   
   If (dblTimeUTC <> ERROR_VALUE) And Not m_TimeInUTC Then dblTimeUTC = dblTimeUTC + CalcTimeZoneAsTime
   
   CalcTwilightState = dblTimeUTC

End Function

Private Function CalcTwilightStateUTC(ByVal SunDepression As Double, ByVal Dusk As Boolean) As Double

Dim dblHourAngle As Double
Dim dblLatitude  As Double
Dim dblLongitude As Double
Dim dblTime      As Double
Dim dblTimeUTC   As Double

   dblLatitude = CalcLatitude
   dblLongitude = -m_Longitude
   dblTime = CalcTimeJulianCenturies(CalcJulianDay)
   dblHourAngle = CalcHourAngleSun(dblLatitude, CalcSunDeclination(dblTime), Dusk)
   
   If dblHourAngle = ERROR_VALUE Then
      dblTimeUTC = ERROR_VALUE
      
   Else
      dblTimeUTC = (720 + (dblLongitude - RadiansToDegrees(dblHourAngle)) * 4 - CalcEquationOfTime(dblTime)) / 1440
      dblTime = CalcTimeJulianCenturies(CalcJulianDayFromJulianCenturies(dblTime) + dblTimeUTC)
      dblHourAngle = CalcHourAngleSun(dblLatitude, CalcSunDeclination(dblTime), Dusk, SunDepression)
      
      If dblHourAngle = ERROR_VALUE Then
         dblTimeUTC = ERROR_VALUE
         
      Else
         dblTimeUTC = (720 + (dblLongitude - RadiansToDegrees(dblHourAngle)) * 4 - CalcEquationOfTime(dblTime)) / 1440
      End If
   End If
   
   CalcTwilightStateUTC = dblTimeUTC

End Function

' Check Latitude, Longitude for North, South, East or West
Private Function CheckCoordinate(ByVal Coordinate As String, Optional ByVal Latitude As Boolean) As String

Dim strDirection As String

   If Val(Coordinate) = ERROR_VALUE Then
      CheckCoordinate = Coordinate
      Exit Function
   End If
   
   Select Case UCase(Right(Coordinate, 1))
      Case "N", "E"
         Coordinate = Trim(Left(Coordinate, Len(Coordinate) - 1))
         
      Case "S", "W"
         Coordinate = "-" & Trim(Left(Coordinate, Len(Coordinate) - 1))
         
      Case Else
         If Val(Coordinate) < 0 Then
            strDirection = " " & Mid("SW", Latitude + 2, 1)
            Coordinate = Mid(Coordinate, 2)
            
         Else
            strDirection = " " & Mid("NE", Latitude + 2, 1)
         End If
   End Select
   
   CheckCoordinate = Coordinate & strDirection

End Function

' Check Latitude = -90 to 90 or Longitude = -180 to 180 or Timezone = -12 to 14
Private Function CheckDistance(ByVal DistanceType As String, ByVal Value As Double) As Boolean

Dim blnError As Boolean

   Select Case DistanceType
      Case "Latitude"
         If (Value < -90) Or (Value > 90) Then blnError = True
         
      Case "Longitude"
         If (Value < -180) Or (Value > 180) Then blnError = True
         
      Case "TimeZone"
         If (Value < -12) Or (Value > 14) Then blnError = True
   End Select
   
   CheckDistance = Not blnError

End Function

Private Function CheckTimeZone(ByVal TimeZone As Double, ByVal Longitude As Double) As Double

   If ((TimeZone > 0) And (Longitude < 0)) Or ((TimeZone < 0) And (Longitude >= 0)) Then TimeZone = -TimeZone
   
   CheckTimeZone = TimeZone

End Function

' Decimal Degree to Degrees, Minutes, Seconds
Private Function DecDegToDegMinSec(ByVal Coordinate As String, Optional ByVal MoreAccurate As Boolean = True, Optional Delimiter As String = ":") As String

Dim dblDegrees As Double
Dim dblMinutes As Double
Dim dblSeconds As Double
Dim intDegrees As Integer
Dim intMinutes As Integer
Dim intSign    As Integer
Dim strSeconds As String
Dim strFormat  As String

   Coordinate = UCase(Trim(Coordinate))
   
   If Right(Coordinate, 1) = "N" Or Right(Coordinate, 1) = "E" Then
      Coordinate = Trim(Left(Coordinate, Len(Coordinate) - 1))
      
   ElseIf Right(Coordinate, 1) = "S" Or Right(Coordinate, 1) = "W" Then
      Coordinate = Trim(Left(Coordinate, Len(Coordinate) - 1))
      
      If Left(Coordinate, 1) <> "-" Then Coordinate = "-" & Coordinate
   End If
   
   If Not IsNumeric(Coordinate) Then
      DecDegToDegMinSec = CStr(ERROR_VALUE)
      Exit Function
   End If
   
   dblDegrees = CDbl(ValidateValue(CStr(Coordinate)))
   intSign = Sgn(dblDegrees)
   dblDegrees = Abs(dblDegrees)
   dblDegrees = (dblDegrees - (Int(dblDegrees / 3600) * 3600))
   intDegrees = Int(dblDegrees)
   dblMinutes = Abs(dblDegrees - intDegrees) * 60
   intMinutes = Int(dblMinutes)
   
   If MoreAccurate Then
      strFormat = dblDegrees
      strFormat = "." & String(Len(strFormat) - InStr(strFormat, ","), vbKey0)
      
      If strFormat = "." Then strFormat = ""
   End If
   
   dblSeconds = (dblMinutes - intMinutes) * 60
   
   If Int(dblSeconds) > 59 Then
      dblSeconds = dblSeconds - 60
      intMinutes = intMinutes + 1
   End If
   
   If intMinutes > 59 Then
      intMinutes = intMinutes - 60
      intDegrees = intDegrees + 1
   End If
   
   strSeconds = Format(dblSeconds, "00" & strFormat)
   
   Do While Right(strSeconds, 1) = "0"
      strSeconds = Left(strSeconds, Len(strSeconds) - 1)
   Loop
   
   If Right(strSeconds, 1) = "," Then strSeconds = Left(strSeconds, Len(strSeconds) - 1)
   If Len(strSeconds) < 2 Then strSeconds = CStr(Val(strSeconds)) & "0"
   
   DecDegToDegMinSec = intSign * intDegrees & "°" & Delimiter & Format(intMinutes, "00") & "'" & Delimiter & strSeconds & Chr(34)

End Function

' Degrees, Minutes, Seconds to Decimal Degree
Private Function DegMinSecToDecDeg(ByVal Coordinate As String, Optional ByVal MoreAccurate As Boolean = True, Optional Delimiter As String = ":") As Double

Dim dblDegrees As Double
Dim dblSeconds As Double
Dim intDegrees As Integer
Dim intMinutes As Integer
Dim intSign    As Integer
Dim strValue() As String

   Coordinate = Trim(Coordinate)
   strValue() = Split(Coordinate, Delimiter)
   
   For intSign = 0 To UBound(strValue)
      If Not IsNumeric(Right(strValue(intSign), 1)) Then strValue(intSign) = Left(strValue(intSign), Len(strValue(intSign)) - 1)
      If Not IsNumeric(strValue(intSign)) Then Exit For
   Next 'intSign
   
   If intSign < 3 Then
      dblDegrees = ERROR_VALUE
      
   Else
      intDegrees = CInt(strValue(0))
      intSign = Sgn(intDegrees)
      
      If Right(Coordinate, 1) = "S" Or Right(Coordinate, 1) = "W" Then intSign = -1
      
      intDegrees = Abs(intDegrees)
      intMinutes = CInt(strValue(1))
      dblSeconds = CDbl(strValue(2))
      dblDegrees = Round(((intMinutes * 60) + dblSeconds) / 3600, ((Len(strValue(2)) - InStr(strValue(2), ",") + 3) And MoreAccurate))
      dblDegrees = (intDegrees + dblDegrees) * intSign
   End If
   
   DegMinSecToDecDeg = dblDegrees
   Erase strValue

End Function

Private Function DegreesToRadians(ByVal AngleDegrees As Double) As Double

   DegreesToRadians = AngleDegrees / 180 * PI

End Function

Private Function FillCityInfo(ByVal Name As String, ByVal Latitude As Variant, ByVal Longitude As Variant, ByVal TimeZone As Double) As String

Dim strError As String

   With CityBuffer
      TimeZone = CheckTimeZone(TimeZone, Longitude)
      .Name = Name
      .TimeZone = TimeZone
      
      If IsNumeric(Latitude) Then
         .Latitude = CDbl(ValidateValue(Latitude))
         
      ElseIf VarType(Latitude) = vbString Then
         .Latitude = DegMinSecToDecDeg(CheckCoordinate(Latitude, True))
         
      Else
         .Latitude = ERROR_VALUE
      End If
      
      If IsNumeric(Longitude) Then
         .Longitude = CDbl(ValidateValue(Longitude))
         
      ElseIf VarType(Longitude) = vbString Then
         .Longitude = DegMinSecToDecDeg(CheckCoordinate(Longitude))
         
      Else
         .Longitude = ERROR_VALUE
      End If
      
      If Not CheckDistance("Latitude", .Latitude) Then
         .Latitude = 0
         strError = "Latitude: " & Latitude & "? (-90/90)  "
      End If
      
      If Not CheckDistance("Longitude", .Longitude) Then
         .Longitude = 0
         strError = strError & "Longitude: " & Longitude & "? (-180/180)  "
      End If
      
      If Not CheckDistance("TimeZone", .TimeZone) Then
         .TimeZone = 0
         strError = strError & "TimeZone: " & TimeZone & "? (-12/14)"
      End If
   End With
   
   FillCityInfo = strError

End Function

Private Function FindSunState(ByVal JulianDay As Double, ByVal Sunset As Boolean, ByVal NextState As Boolean) As Double

Dim dblJulianDay As Double
Dim dblTime      As Double
Dim dblValue     As Double
Dim intDay       As Integer

   dblValue = -1 + (2 And NextState)
   dblJulianDay = JulianDay
   dblTime = CalcSunStateUTC(dblJulianDay, Sunset)
   
   Do While (dblTime = ERROR_VALUE) And (intDay < 367)
      intDay = intDay + 1
      dblJulianDay = dblJulianDay + dblValue
      dblTime = CalcSunStateUTC(dblJulianDay, Sunset)
   Loop
   
   FindSunState = dblJulianDay

End Function

Private Function FracPart(ByVal X As Double) As Double

   X = X - Int(X)
   
   If X < 0 Then X = X + 1
   
   FracPart = X

End Function

Private Function LimitsForDate() As Boolean

Dim dblDate As Double

   dblDate = CDbl(m_DateCalculate)
   
   If (dblDate > -657418.176956019) And (dblDate <= 2958465.99999) And (dblDate <> -657418) Then
      If (dblDate <= -115869) Or (dblDate > -115859.99999) Then LimitsForDate = True
   End If

End Function

Private Function LimitsForJulianDay(ByVal JulianDay As Double) As Boolean

   If (JulianDay >= 1757593.50001) And (JulianDay <= 5373484.49998843) Then LimitsForJulianDay = True

End Function

Private Function MakeDate(ByVal Time As Double) As Date

   MakeDate = DateAdd("s", Time * 86400, DateSerial(Year(m_DateCalculate), Month(m_DateCalculate), Day(m_DateCalculate)))

End Function

Private Function MakeTime(ByVal Time As Double) As Date

Dim dblTime   As Double
Dim lngHour   As Long
Dim lngMinute As Long
Dim lngSecond As Long

   dblTime = (Time + ((m_TimeZone + CalcDaylightSavingTime / 60) And Not m_TimeInUTC)) * 3600
   lngHour = Int(dblTime / 3600)
   dblTime = dblTime - lngHour * 3600
   lngMinute = Int(dblTime / 60)
   lngSecond = (dblTime - lngMinute * 60)
   MakeTime = DateValue(m_DateCalculate) & " " & Format(TimeSerial(lngHour, lngMinute, lngSecond), "hh:mm:ss")

End Function

Private Function QuickSort(ByRef DataArray() As CityInfoType, ByVal LowerBound As Long, ByVal UpperBound As Long)

Dim lngLower  As Long
Dim lngUpper  As Long
Dim strMiddle As String

   If UpperBound <= LowerBound Then Exit Function
   
   lngLower = LowerBound
   lngUpper = UpperBound
   strMiddle = DataArray((LowerBound + UpperBound) \ 2).Name
   
   Do While (lngLower <= lngUpper)
      Do While DataArray(lngLower).Name < strMiddle
         lngLower = lngLower + 1
         
         If lngLower = UpperBound Then Exit Do
      Loop
      
      Do While strMiddle < DataArray(lngUpper).Name
         lngUpper = lngUpper - 1
         
         If lngUpper = LowerBound Then Exit Do
      Loop
      
      If lngLower <= lngUpper Then
         CityBuffer = DataArray(lngLower)
         DataArray(lngLower) = DataArray(lngUpper)
         DataArray(lngUpper) = CityBuffer
         lngLower = lngLower + 1
         lngUpper = lngUpper - 1
      End If
   Loop
   
   If LowerBound < lngUpper Then QuickSort DataArray(), LowerBound, lngUpper
   If lngLower < UpperBound Then QuickSort DataArray(), lngLower, UpperBound

End Function

Private Function Radians(ByVal X As Double) As Double

   Radians = (Sgn(X) * Abs(X - (Int(X / 360) * 360))) * RadianAngle

End Function

Private Function RadiansToDegrees(ByVal AngleRadians As Double) As Double

   RadiansToDegrees = AngleRadians * 180 / PI

End Function

Private Sub AddCityInfo()

   With CityBuffer
      .TimeZone = CheckTimeZone(.TimeZone, .Longitude)
   End With
   
   If m_CityCount > UBound(CityInfo) Then ReDim Preserve CityInfo(UBound(CityInfo) + 1) As CityInfoType
   
   CityInfo(m_CityCount) = CityBuffer
   m_CityCount = m_CityCount + 1
   QuickSort CityInfo(), LBound(CityInfo), UBound(CityInfo)

End Sub

Private Sub CalcQuadratic(ByRef Y() As Double, ByRef Zero() As Double)

Dim dblDiscriminant As Double
Dim dblResult(1)    As Double
Dim dblX            As Double

   dblResult(0) = (Y(0) + Y(2)) / 2 - Y(1)
   dblResult(1) = (Y(2) - Y(0)) / 2
   dblX = -dblResult(1) / (dblResult(0) * 2)
   Y(3) = (dblResult(0) * dblX + dblResult(1)) * dblX + Y(1)
   dblDiscriminant = dblResult(1) * dblResult(1) - dblResult(0) * Y(1) * 4
   
   If dblDiscriminant > 0 Then
      dblDiscriminant = Sqr(dblDiscriminant) / Abs(dblResult(0)) / 2
      Zero(1) = dblX - dblDiscriminant
      Zero(2) = dblX + dblDiscriminant
      
      If Abs(Zero(1)) <= 1 Then Zero(0) = Zero(0) + 1
      If Abs(Zero(2)) <= 1 Then Zero(0) = Zero(0) + 1
      If Zero(1) < -1 Then Zero(1) = Zero(2)
   End If
   
   Erase dblResult

End Sub

Private Sub LoadCityInfo()

Dim blnGetRecord As Boolean
Dim intFields    As Integer
Dim intForInput  As Integer
Dim strField()   As String
Dim strLine      As String

   ReDim CityInfo(0) As CityInfoType
   
   m_CityCount = 0
   intForInput = FreeFile
   
   Open m_MapCityInfo & FILENAME_CITYINFO For Input As #intForInput
      Do While Not EOF(intForInput)
         Line Input #intForInput, strLine
         
         If blnGetRecord And Len(Trim(strLine)) Then
            strField() = Split(strLine, "=")
            
            If UBound(strField) = 1 Then
               strField(0) = UCase(Trim(strField(0)))
               strField(1) = Trim(strField(1))
               
               With CityBuffer
                  If strField(0) = "NAME" Then
                     .Name = strField(1)
                     intFields = intFields + 1
                  End If
                  
                  If strField(0) = "LATITUDE" Then
                     .Latitude = CDbl(ValidateValue(strField(1)))
                     intFields = intFields + 1
                  End If
                  
                  If strField(0) = "LONGITUDE" Then
                     .Longitude = CDbl(ValidateValue(strField(1)))
                     intFields = intFields + 1
                  End If
                  
                  If strField(0) = "TIMEZONE" Then
                     .TimeZone = CDbl(ValidateValue(strField(1)))
                     intFields = intFields + 1
                  End If
               End With
               
               If intFields = 4 Then
                  intFields = 0
                  
                  Call AddCityInfo
               End If
            End If
         End If
         
         If UCase(strLine) = "[CITY_INFO]" Then blnGetRecord = True
      Loop
   Close #intForInput
   
   Erase strField

End Sub

Private Sub SaveCityInfo()

Dim intForOutput As Integer
Dim lngIndex     As Long

   intForOutput = FreeFile
   
   Open m_MapCityInfo & FILENAME_CITYINFO For Output As #intForOutput
      Print #intForOutput, "[City_Info]"
      
      For lngIndex = 0 To m_CityCount - 1
         With CityInfo(lngIndex)
            Print #intForOutput, "Name      = " & .Name
            Print #intForOutput, "Latitude  = " & .Latitude
            Print #intForOutput, "Longitude = " & .Longitude
            Print #intForOutput, "TimeZone  = " & .TimeZone
            
            If lngIndex < m_CityCount - 1 Then Print #intForOutput, ""
         End With
      Next 'lngIndex
   Close #intForOutput

End Sub

Private Sub Class_Initialize()

   m_MapCityInfo = App.Path
   
   If Right(m_MapCityInfo, 1) <> "\" Then m_MapCityInfo = m_MapCityInfo & "\"
   
   m_MapCityInfo = m_MapCityInfo & "Data\"
   
   If Len(Dir(m_MapCityInfo & FILENAME_CITYINFO)) Then
      Call LoadCityInfo
      
   Else
      ReDim CityInfo(0) As CityInfoType
   End If
   
   m_DateCalculate = Now
   RadianAngle = Atn(1) / 45

End Sub

Private Sub Class_Terminate()

   Erase CityInfo

End Sub
