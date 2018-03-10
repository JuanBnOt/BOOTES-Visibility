''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MoonImage,
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const CCDnumber = 0 'CCD Number
Const appError = 0.00015166666666 'error given on degrees per second

'Geographical data of the site

Const latitude = 36.75916666666667
Const longitude = -4.0411111111111
Const height = 0


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public connStr, objConn


'Initialization of the Data Base
Sub DataBaseInitialization()
	'Database initialization
	connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Telescope\scripts\bootes.accdb"
	 'Define object type
	Set objConn = CreateObject("ADODB.Connection")
	 'Open Connection
	objConn.open connStr
End Sub



'this sub inserts an observation on the BOOTES DataBase
Sub insertObservation (tgtName, tgtRA, tgtDEC, tgtCCD, tgtExpTime, tgtDate, tgtTime)	
		SQLInst="INSERT INTO OP ( [tgtName], [RA], [DEC], [CCD], [ExpTime], [DatePlan], [TimePlan] ) VALUES ('" + cstr(tgtName) + "','" + Cstr(tgtRA) + "','" + Cstr(tgtDEC) + "','" + Cstr(tgtCCD) + "','" + cstr(tgtExpTime) + "','" + cstr(tgtDate) + "','" + cstr(tgtTime) + "');"
			objConn.Execute(SQLInst)
End Sub



'this function returns the value of the angle between 0 and 359 degrees
Function Angle0to360(angle)
	If angle >= 360 Then
		Angle0to360 = angle - Fix(angle/360) * 360
	ElseIf angle <= 0 Then 
		Angle0to360 = angle - Int(angle/360) * 360
	Else
		Angle0to360 = angle
	End If
End Function



'This function return the number of days of the year until the selected date
Function NumberDaysYear(year1, month1, day1)
	count = 1
	While count <= month1 
	'if we got a even number we sum 30 days, unless is August that we sum 31 or February 28
		If count Mod 2 = 0 Then 
			If count = 8 Then
				totaldays = totaldays + 31
			ElseIf count = 2 Then
				totaldays = totaldays + 28
			Else
				totaldays = totaldays + 30
			End If
		Else
			totaldays = totaldays + 31
		End If
		count = count + 1
	Wend
	'We need to subtract some days because we have added the whole month
	If month1 Mod 2 = 0 then
		If month1 = 8 Then
			totaldays = totaldays + day1 - 31
		Else
			totaldays = totaldays + day1 - 30
		End If
	Else
		totaldays = totaldays + day1 - 31
	End If
	NumberDaysYear = totaldays
End Function 



'Returns the Hour in HH:MM:SS format from a fractional hour
Function FractionalToHMS(fractionalHour)
	splitHora = Split(fractionalHour,",")
	horaTemp = CInt(splitHora(0))
	minTemp = CInt((fractionalHour - horaTemp) * 60)
	FractionalToHMS = CStr(TimeSerial(horaTemp, minTemp, 0))
End Function 





'Returns the position of the moon (RA and DEC on degrees) and the bright given as a fraction of unity
Function MoonData(year1, month1, day1, hour1, ByRef bright, ByRef moonRA, ByRef moonDEC)
	objV.MoonDistanceAndData_SelectDay year1, month1, day1, hour1, objRA, objDEC, distancia, brillo, moonRA, moonDEC
End Function 





'This sub gives the nearest day of the beginning of the lunar cicle and the end of it
Sub MoonCycle(selectedDate, ByRef bDate, ByRef eDate)
	bDate = selectedDate	
	Do 
		bDate = DateAdd("d", 1, bDate)
		JD = apASCOM1.JulianDate(Year(bDate), Month(bDate), Day(bDate), 0)
		phase = apUtils.MoonPhase(JD)
	Loop Until phase > -180 And phase < -160
		eDate = bDate
	Do
		eDate = DateAdd("d", 1, eDate)
		JD = apASCOM1.JulianDate(Year(eDate), Month(eDate), Day(eDate), 0)
		phase = apUtils.MoonPhase(JD)
	Loop Until phase > 160 And phase < 180
End Sub



'This function returns the name of the phase of the moon
Function MoonPhase(dateSelected)
	JD = apASCOM1.JulianDate(Year(dateSelected), Month(dateSelected), Day(dateSelected), 0)
	phase = apUtils.MoonPhase(JD)
     	If phase >= -180.0 And phase < -135.0 Then 
        	MoonPhase = "Full Moon"
     	ElseIf phase >= -135.0 And  phase < -90.0 Then
       	 	MoonPhase = "Waning Gibbous"
     	ElseIf phase >= -90.0 And  phase < -45.0 Then
     	    MoonPhase = "Last Quarter"
     	ElseIf phase >= -45.0 And phase < 0.0 Then
     	    MoonPhase = "Waning Crescent"
    	ElseIf phase >= 0.0 And phase < 45.0 Then
    	    MoonPhase = "New Moon"
     	ElseIf phase >= 45.0 And phase < 90.0 Then
        	MoonPhase = "Waxing Crescent"
     	ElseIf phase >= 90.0 And phase < 135.0 Then
     	    MoonPhase = "First Quarter"
     	ElseIf phase >= 135.0 And phase < 180.0 Then
     	    MoonPhase = "Waxing Gibbous"
	End If
End Function







'This function search in the DataBase the hour of sunRise and sunSet
Function GetHourNight(daysYear, ByRef sunSet, ByRef sunRise)
	sqlRise = "SELECT Sunrise FROM Timetable WHERE Numday = " + CStr(daysYear) + " "
	sqlSet = "SELECT Sunset FROM Timetable WHERE Numday = " + CStr(daysYear) + " "
	
	Set sunSetTemp = objConn.Execute(sqlSet)
	Set sunRiseTemp = objConn.Execute(sqlRise)

	sunSet = sunSetTemp(0)
	sunRise = sunRiseTemp(0)
End Function






'Returns the Topocentric Elevation of the moon, data must be in degrees
Function MoonElevation(year1, month1, day1, hour1, latitude, longitude, height, moonRA, moonDEC)		
	JD = apASCOM1.JulianDate(year1, month1, day1 ,hour1)

	apTransform.SiteLatitude = latitude
	apTransform.SiteLongitude = longitude
	apTransform.SiteTemperature = 15
	apTransform.SiteElevation = height
	apTransform.JulianDateUTC = JD
	
	moonRA = moonRA / 15
	
	apTransform.SetJ2000 moonRA, moonDEC
	
	MoonElevation = apTransform.ElevationTopocentric
End Function



'This function returns the higher Elevation of the Moon during the night, also gives the MoonRA and MoonDEC on degrees for that value of Elevation. dateSelected is Date type variable
Function HigherMoonEleveationNight(dateSelected, latitude, longitude, height, ByRef bRA, ByRef bDEC, ByRef bHour, ByRef lowElevation)
	Dim inc
	Dim bElevation, moonRA, moonDEC, elevationLoop
	
	'increment of 10 minutes for each cycle
	inc = 1 / 6

	'Getting information of the night selected
	daysYear = NumberDaysYear(Year(dateSelected), Month(dateSelected), Day(dateSelected))
	GetHourNight daysYear, sunSet, sunRise
	
	'night limit hours, we sum 1 and subtract 1 on sunSet and sunRise to get better visibility
	sunSet = Hour(sunSet) + Minute(sunRise)/60 + 1
	sunRise = Hour(sunRise) + Minute(sunRise)/60 - 1
	
	'initialization of variables
	bHour = sunSet
	bRA = 0
	bDEC = 0
	bElevation = -90
	raLoop = 0
	decLoop = 0
	elevationLoop = 0
	
	'Calculamos la altura más alta de la luna durante la noche, para ellos cada 10 minutos
	'calculamos la altura. Para ello calculamos desde la hora de Pûesta hasta las 00:00
	' y luego desde las 00:00 hasta la salida del Sol
	
	'Calculation of the highest elevation of the moon during the night. To do that, we sum 10 min on each iteration
	'To make right the calculus, after the 23:59 we must sum a day to dateSelected and after that we can make the
	'calculation until the sunRise.
	
	Do Until sunSet > 24
			MoonData Year(dateSelected), Month(dateSelected), Day(dateSelected), sunSet, bright, moonRA, moonDEC
			elevationLoop = MoonElevation(Year(dateSelected), Month(dateSelected), Day(dateSelected), sunSet, latitude, longitude, height, moonRA, moonDEC)
		If elevationLoop > bElevation Then
			bElevation = elevationLoop
			bRA = moonRA
			bDEC = moonDEC
			bHour = sunSet
		End If
		sunSet = sunSet + inc
	Loop

	'We must change the Date to continue with the calculation of the best Elevetion
	sunSet = 0

	'Must sum a day to our Date
	newDate = DateAdd("d", 1, dateSelected)
	
	'Now, we study the second part from 00:00 to sunrise
	Do Until sunSet > sunRise
		MoonData Year(newDate), Month(newDate), Day(newDate), sunRise, bright, moonRA, moonDEC
		elevationLoop = MoonElevation(Year(newDate), Month(newDate), Day(newDate), sunSet, latitude, longitude, height, moonRA, moonDEC)
		If elevationLoop > bElevation Then
			bElevation = elevationLoop
			bRA = moonRA
			bDEC = moonDEC
			bHour = sunSet
		End If
		sunSet = sunSet + inc
	Loop
	
	'We put a condition to exit the script if the elevation of the moon is less than 10º during the night
	If bElevation < 10 Then
		lowElevation = 1
	End If
	HigherMoonEleveationNight = bElevation
	
	  
End Function














'This function programs the observation plan for one night, ExpTime must be in seconds
Sub OneNightObservation(dateSelected,latitude, longitude, height, CCDangle, CCD, ExpTime)
	
	'Initialization and declaration of variables
	Dim infLim, supLim, steps, count, moonRA, moonDEC
	count = 0
	
	'Calculation of JulianDate
	JD = apASCOM1.JulianDate(Year(dateSelected), Month(dateSelected), Day(dateSelected), 0)
	
	'Calculation of highest Elevation and RA-DEC for that moment, and update dateSelected
	altura = HigherMoonEleveationNight(dateSelected, latitude, longitude,height, moonRA, moonDEC, bHour, lowElevation)
	
	If lowElevation = 1 Then
	lowElevation = 0
	WScript.Echo "Altura mayor menor de 10 grados, estudiando el siguiente dia. Elevetion less than 10 degrees, studying the next day"
		Exit Sub
	End If

	newDate = DateAdd("s",Int(bHour * 3600),dateSelected)
	
	'Calculo del número de ímagenes según la CCD y los límites inferiores y superiores
	If CCDAngle < 5 Then
		infLim = 5
	Else
		infLim = CCDAngle
	End If

		supLim = 35
	
	'Number of steps of the For Loop depending on the CCD
		steps = Int((supLim - infLim) / CCDAngle + 1)
		Dim objRightRA(), objLeftRA(),objRightDEC(), objLeftDEC(), distancia(), totalError, updateTime
		
		ReDim distancia(steps - 1)
		ReDim objRightRA(steps - 1)
		ReDim objLeftRA(steps - 1)
		ReDim objRightDEC(steps - 1)
		ReDim objLeftDEC(steps - 1)
		
		totalError = 0
		
		
	'objective RA calculation depending on the CCD
	For i=1 To steps
		If totalError > 0.45 Then
			newDate = DateAdd("s",updateTime,newDate)
			updateTime = 0
			totalError = 0
			newHour = Hour(newDate) + Minute(newDate)/60
			MoonData Year(newDate), Month(newDate), Day(newDate), newHour, bright, moonRA, moonDEC
			WScript.Echo "datos luna actualizados para evitar error"

		End If
		distancia(i -1) = i * CCDangle
		objRightDEC(i -1) = moonDEC
		objRightRA(i - 1) = Angle0to360(distancia(i-1) / Cos(moonDEC) + moonRA)
		updateTime = updateTime + ExpTime
		totalError = totalError + ExpTime * appError
	Next
	
	For i=1 To steps
		If totalError > 0.45 Then
			newDate = DateAdd("s",updateTime,newDate)
			updateTime = 0
			totalError = 0
			newHour = Hour(newDate) + Minute(newDate)/60
			MoonData Year(newDate), Month(newDate), Day(newDate), newHour, bright, moonRA, moonDEC
			WScript.Echo "datos luna actualizados para evitar error"
		End If
		distancia(i -1) = i * CCDangle
		objLeftDEC(i -1) = moonDEC
		objLeftRA(i -1) = Angle0to360(distancia(i-1) / Cos(moonDEC) - moonRA)
		updateTime = updateTime + ExpTime
		totalError = totalError + ExpTime * appError
	Next
	
	'Get obsevations into de BOOTES DataBase
	'Shared variables between 2 sides of the observation
	
	tgtCCD = CCD
	tgtExpTime = ExpTime
	tgtDate = CStr(newDate)
	tgtTime = FractionalToHMS(bHour)
	phase = MoonPhase(newDate)
	bright = Fix(apUtils.MoonIllumination(JD) * 100)
	
	'Giving the name and inserting the observations on the DataBase
	For Each item In objRightRA
		tgtName =  "Moon_Der_bright_" + CStr(bright) + "%_d" + CStr(distancia(count)) + "_" +  CStr(phase)
		tgtRA = CStr(item)
		tgtDEC = CStr(objRightDEC(count))
		insertObservation tgtName, tgtRA, tgtDEC, tgtCCD, tgtExpTime, tgtDate, tgtTime
		count = count + 1
	Next
	count = 0
	For Each item In objLeftRA
		tgtName =  "Moon_Left_bright_" + CStr(bright) + "%_d" + CStr(distancia(count)) + "_" +  CStr(phase)
		tgtRA = CStr(item)
		tgtDEC = CStr(objLeftDEC(count))
		insertObservation tgtName, tgtRA, tgtDEC, tgtCCD, tgtExpTime, tgtDate, tgtTime
		count = count + 1
	Next
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'			MAIN PROGRAM
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Initialization

'Creating object from ASCOM and Visibility DLL
Set objV = CreateObject("BOOTES_Visibility.Visibility")
Set apASCOM1 = CreateObject("ASCOM.Astrometry.NOVAS.NOVAS31")
Set apTransform = CreateObject("ASCOM.Astrometry.Transform.Transform")
Set visibilityObj = CreateObject("BOOTES_Visibility.Visibility")
Set apUtils = CreateObject("ASCOM.Astrometry.AstroUtils.AstroUtils")

'initialization of variables 
Dim dateLoop, CCDangle, CCD, TimeExp

'Getting information about the Telescope and Exposure Time
CCDangle = 5
CCD = 0
TimeExp = 60

'Sub to Initialize the DataBase
DataBaseInitialization()

'Getting the dates when the moon Cycle begins and finishes
selectedDate = CDate(#24/04/2018#)
MoonCycle selectedDate, bDate, eDate
dateLoop = bDate

'Doing the observation during the whole Cycle
Do Until dateLoop > eDate

WScript.Echo "dia de estudio: ", dateLoop

OneNightObservation dateLoop, latitude, longitude, height, CCDangle, CCD, TimeExp

dateLoop = DateAdd("d", 1, dateLoop)

Loop





