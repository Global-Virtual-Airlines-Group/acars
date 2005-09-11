Attribute VB_Name = "modFDR"
Option Explicit

Private isTurboProp As Boolean
Private EngineCount As Integer

'Flight status bit position constants.
Public Const FLIGHTPAUSED = 1
Public Const FLIGHTSLEWING = 2
Public Const FLIGHTPARKED = 4
Public Const FLIGHTONGROUND = 8

Public Const FLIGHT_SP_ARM = &H10
Public Const FLIGHT_GEAR_DOWN = &H20

Public Const FLIGHT_AP_GPS = &H100
Public Const FLIGHT_AP_NAV = &H200
Public Const FLIGHT_AP_HDG = &H400
Public Const FLIGHT_AP_APR = &H800
Public Const FLIGHT_AP_ALT = &H1000
Public Const FLIGHT_AT_IAS = &H2000
Public Const FLIGHT_AT_MACH = &H4000

Public Function RecordFlightData() As PositionData

Dim lngResult As Long
Dim x As Integer

Dim fsTime(1) As Byte
Dim curLat As Currency
Dim curLon As Currency

Dim terrElevation As Long
Dim altMSL As Long

Dim magVar As Integer
Dim hdg As Long

Dim ASpeed As Long
Dim GSpeed As Long
Dim VSpeed As Long
Dim Mach As Integer

Dim flapsL As Integer
Dim flapsR As Integer
Dim Spoilers As Long
Dim gearPos As Long

Dim RawSimRate As Integer
Dim isPaused As Integer
Dim isSlew As Integer

'Make sure we're connected to FSUIPC
If Not config.FSUIPCConnected Then Exit Function

'Set critical error handler
On Error GoTo FatalError

'Get sim time
Call FSUIPC_Read(&H23B, 2, VarPtr(fsTime(0)), lngResult)

'Get latitude/longitude
Call FSUIPC_Read(&H560, 8, VarPtr(curLat), lngResult)
Call FSUIPC_Read(&H568, 8, VarPtr(curLon), lngResult)

'Get terrain elevation/altitude MSL/heading/magVariation
Call FSUIPC_Read(&H20, 4, VarPtr(terrElevation), lngResult)
Call FSUIPC_Read(&H574, 4, VarPtr(altMSL), lngResult)
Call FSUIPC_Read(&H2A0, 2, VarPtr(magVar), lngResult)
Call FSUIPC_Read(&H580, 4, VarPtr(hdg), lngResult)

'Get air speed/ground speed/vertical speed/Mach
Call FSUIPC_Read(&H2BC, 4, VarPtr(ASpeed), lngResult)
Call FSUIPC_Read(&H2B4, 4, VarPtr(GSpeed), lngResult)
Call FSUIPC_Read(&H2C8, 4, VarPtr(VSpeed), lngResult)
Call FSUIPC_Read(&H11C6, 2, VarPtr(Mach), lngResult)
    
'Get left/right inboard flaps positions/spoilers/gear positions
Call FSUIPC_Read(&H30F0, 2, VarPtr(flapsL), lngResult)
Call FSUIPC_Read(&H30F4, 2, VarPtr(flapsR), lngResult)
Call FSUIPC_Read(&HBD0, 4, VarPtr(Spoilers), lngResult)
Call FSUIPC_Read(&HBE8, 4, VarPtr(gearPos), lngResult)

'Get sim rate/pause/slew
Call FSUIPC_Read(&HC1A, 2, VarPtr(RawSimRate), lngResult)
Call FSUIPC_Read(&H264, 2, VarPtr(isPaused), lngResult)
Call FSUIPC_Read(&H5DC, 2, VarPtr(isSlew), lngResult)

Dim apMode As Long
Dim GPSmode As Long
Dim NAVmode As Long
Dim HDGmode As Long
Dim APRmode As Long
Dim ALTmode As Long
Dim IASmode As Long
Dim MCHmode As Long

'Get AP/NAV/HDG/APR/ALT/IAS/MCH autopilot modes
Call FSUIPC_Read(&H7BC, 4, VarPtr(apMode), lngResult)
Call FSUIPC_Read(&H132C, 4, VarPtr(GPSmode), lngResult)
Call FSUIPC_Read(&H7C4, 4, VarPtr(NAVmode), lngResult)
Call FSUIPC_Read(&H7C8, 4, VarPtr(HDGmode), lngResult)
Call FSUIPC_Read(&H800, 4, VarPtr(APRmode), lngResult)
Call FSUIPC_Read(&H7D0, 4, VarPtr(ALTmode), lngResult)
Call FSUIPC_Read(&H7DC, 4, VarPtr(IASmode), lngResult)
Call FSUIPC_Read(&H7E4, 4, VarPtr(MCHmode), lngResult)

'Get fuel remaining
Dim FuelWeight As Integer
Dim ZeroFuelWeight As Long

Dim tankOffsets As Variant
Dim tankSizeOffsets As Variant
Dim TankLevel(10) As Long
Dim TankCapacity(10) As Long

'Get tank capacity and sizes
tankOffsets = Array(&HB74, &HB7C, &HB84, &HB8C, &HB94, &HB9C, &HBA4, &H1244, &H124C, &H1254, &H125C)
tankSizeOffsets = Array(&HB78, &HB80, &HB88, &HB90, &HB98, &HBA0, &HBA8, &H1248, &H1250, &H1258, &H1260)
For x = 0 To UBound(tankOffsets)
    Call FSUIPC_Read(CLng(tankOffsets(x)), 4, VarPtr(TankLevel(x)), lngResult)
    Call FSUIPC_Read(CLng(tankSizeOffsets(x)), 4, VarPtr(TankCapacity(x)), lngResult)
Next

'Get fuel/ZF weight
Call FSUIPC_Read(&HAF4, 2, VarPtr(FuelWeight), lngResult)
Call FSUIPC_Read(&H3BFC, 4, VarPtr(ZeroFuelWeight), lngResult)

Dim EngineType As Byte
Dim EngineN1(3) As Double
Dim EngineN2(3) As Double
Dim EngineThrottle(3) As Integer
Dim EngineRunning(3) As Integer

'Get engine type
Call FSUIPC_Read(&H609, 1, VarPtr(EngineType), lngResult)

'Get number of engines
If (EngineCount = 0) Then
    Call FSUIPC_Read(&HAEC, 2, VarPtr(EngineCount), lngResult)
    EngineCount = 4 'Set temporary value for the loop
End If

'Get engine firing status, engine N1, and engine throttle position.
'0AEC    2   Number of Engines
'0898    2   Engine 1 running flag
'0930    2   Engine 2 running flag
'09C8    2   Engine 3 running flag
'0A60    2   Engine 4 running flag
For x = 0 To (EngineCount - 1)
    Call FSUIPC_Read(&H894 + (&H98 * x), 2, VarPtr(EngineRunning(x)), lngResult)
    Call FSUIPC_Read(&H2000 + (&H100 * x), 8, VarPtr(EngineN1(x)), lngResult)
    Call FSUIPC_Read(&H2008 + (&H100 * x), 8, VarPtr(EngineN2(x)), lngResult)
    Call FSUIPC_Read(&H88C + (&H98 * x), 2, VarPtr(EngineThrottle(x)), lngResult)
Next

Dim isOnGround As Integer
Dim isParkBrake As Integer

'Is the aircraft on the ground/parked?
Call FSUIPC_Read(&H366, 2, VarPtr(isOnGround), lngResult)
Call FSUIPC_Read(&HBC8, 2, VarPtr(isParkBrake), lngResult)

'Call FSUIPC and create the flight data bean
Dim data As New PositionData
data.Phase = info.Phase
If Not FSUIPC_Process(lngResult) Then
    FSError lngResult
    Exit Function
End If

'Build sim time string and update status bar.
frmMain.sbMain.Panels(4).Text = "Sim Time: " & Format(fsTime(0), "00") & ":" & Format(fsTime(1), "00") & " Z"

'Calculate latitude/longitude
data.Latitude = (curLat * 10000#) * 90# / (10001750# * 65536# * 65536#)
data.Longitude = (curLon * 10000#) * 360# / (65536# * 65536# * 65536# * 65536#)

'Calculate speeds
data.AirSpeed = CInt(ASpeed / 128#)
data.GroundSpeed = CInt(GSpeed * 3600# / 65536# / 1852#)
data.VerticalSpeed = CInt(VSpeed * 60# * 3.28084 / 256#)
data.Mach = CDbl(Mach / 20480#)

'Calculate heading
magVar = magVar * 360# / 65536#
data.Heading = CInt(hdg * (360# / (65536# * 65536#)))
data.Heading = data.Heading - magVar
If data.Heading < 0 Then data.Heading = (360 + data.Heading)

'Calculate Altitude AGL/MSL
terrElevation = (terrElevation * 3.28084) / 256
data.AltitudeMSL = altMSL * 3.28084
data.AltitudeAGL = data.AltitudeMSL - terrElevation

'Calculate flaps setting
120 data.Flaps = (flapsL / 512) + (flapsR / 512)

'Calculate total fuel remaining in all tanks.
Dim FuelGallons As Double
Dim TankPct As Double
For x = 0 To UBound(TankCapacity)
    If TankCapacity(x) > 0 Then
        TankPct = TankLevel(x) / 128# / 65536#
        FuelGallons = FuelGallons + (TankPct * TankCapacity(x))
    End If
Next

'Get fuel/total weight
data.Fuel = FuelGallons * (FuelWeight / 256)
data.Weight = (ZeroFuelWeight / 256) + data.Fuel

'Count how many engines are running and calculate average N1/N2
isTurboProp = (EngineType = 5)
data.EngineCount = EngineCount
Dim EnginesRunning As Integer
For x = 0 To 3
    If EngineRunning(x) Then
        data.setN1 x, EngineN1(x)
        data.setN2 x, EngineN2(x)
        EnginesRunning = EnginesRunning + 1
    Else
        data.setN1 x, 0
        data.setN2 x, 0
    End If
    
    data.setThrottle x, (EngineThrottle(x) / 163.84)
Next

'Calculate Average values, with a DivZero check
data.EnginesStarted = (EnginesRunning > 0)

'Build flags
Dim isAP As Boolean
isAP = (apMode = 1)

data.Paused = (isPaused = 1)
data.Slewing = (isSlew = 1)
data.Parked = (isParkBrake = 32767)
data.onGround = (isOnGround = 1)
data.Spoilers = (Spoilers = 4800)
data.GearDown = (gearPos > 8192)
data.AP_GPS = isAP And (NAVmode = 1) And (GPSmode = 1)
data.AP_NAV = isAP And (NAVmode = 1) And (GPSmode = 0)
data.AP_HDG = isAP And (HDGmode = 1)
data.AP_APR = isAP And (APRmode = 1)
data.AP_ALT = isAP And (ALTmode = 1)
data.AT_IAS = (IASmode = 1)
data.AT_MCH = (MCHmode = 1)

'Save sim rate
data.simRate = RawSimRate

'Return flight position data
Set RecordFlightData = data

ExitFunc:
Exit Function

FatalError:
ShowMessage Error$(err) + " at Line " + CStr(Erl) + " of FDR", ACARSERRORCOLOR
Resume ExitFunc

End Function

Public Function PhaseChanged(cPos As PositionData) As Boolean
    Static TakeoffCheckCount As Integer

    Select Case info.Phase
    Case "Preflight"
            'If the parking brake is released, we enter the Pushback phase.
            If Not cPos.Parked Then
                info.Phase = "Pushback"
                info.TimeOut = Now
                TakeoffCheckCount = 0
                PhaseChanged = True
            End If
        
        Case "Pushback"
            'If an engine has been started, we enter the Taxi Out phase.
            If cPos.EnginesStarted Then
                info.Phase = "Taxi Out"
                info.TaxiOutTime = Now
                info.TaxiFuel = cPos.Fuel
                info.TaxiWeight = cPos.Weight
                PhaseChanged = True
            End If

        Case "Taxi Out"
            'Check if the average throttle > 75%. If so, increment a counter. If
            'that counter reaches 16, then we must be taking off. Also check if we're
            'airborne. If so, jump to takeoff phase, which will record some values
            'then immediately jump to airborne phase.
            If cPos.AverageThrottle > 75 Then
                TakeoffCheckCount = TakeoffCheckCount + 1
            Else
                TakeoffCheckCount = 0
            End If
            
            If ((TakeoffCheckCount > 15) Or (cPos.AirSpeed > 60)) Then
                info.Phase = "Takeoff"
                info.TakeoffTime = Now
                If config.SB3Connected Then SB3Transponder (True)
                PhaseChanged = True
            ElseIf Not cPos.onGround And (cPos.AltitudeAGL > 7) Then
                info.Phase = "Airborne"
                info.TimeOff = Now
                info.TakeoffTime = Now
                If config.SB3Connected Then SB3Transponder (True)
                info.TakeoffSpeed = cPos.AirSpeed
                info.TakeoffFuel = cPos.Fuel
                info.TakeoffWeight = cPos.Weight
                info.TakeoffN1 = cPos.AverageN1
                PhaseChanged = True
            End If

        Case "Takeoff"
            'If we're off the ground, then we enter the Airborne phase.
            If Not cPos.onGround Then
                info.Phase = "Airborne"
                info.TimeOff = Now
                info.TakeoffSpeed = cPos.AirSpeed
                info.TakeoffFuel = cPos.Fuel
                info.TakeoffWeight = cPos.Weight
                info.TakeoffN1 = cPos.AverageN1
                PhaseChanged = True
            End If
            
        Case "Airborne"
            'If we've been in the Airborne phase for more than 30 seconds, and now
            'we're back on the ground, then we enter the Landed phase. The 30 seconds
            'is for debouncing.
            If ((DateDiff("s", info.TimeOff, Now) > 30) And cPos.onGround) Then
                info.Phase = "Landed"
                info.LandingTime = Now
                info.LandingSpeed = cPos.AirSpeed
                info.LandingVSpeed = cPos.VerticalSpeed
                info.LandingFuel = cPos.Fuel
                info.LandingWeight = cPos.Weight
                info.LandingN1 = cPos.AverageN1
                PhaseChanged = True
            End If

        Case "Landed"
            'If ground speed falls below 40 knots, we enter the Taxi In phase.
            If (cPos.GroundSpeed < 40) Then
                info.Phase = "Taxi In"
                info.TaxiInTime = Now
                If config.SB3Connected Then SB3Transponder (False)
                PhaseChanged = True
            ElseIf Not cPos.onGround Then
                info.Phase = "Airborne"
                If config.SB3Connected Then SB3Transponder (True)
                PhaseChanged = True
            End If

        Case "Taxi In"
            'If parking brake is set, we enter the "At Gate" phase.
            If cPos.Parked Then
                info.Phase = "At Gate"
                info.GateTime = Now
                info.GateFuel = cPos.Fuel
                info.GateWeight = cPos.Weight
                info.FlightData = True
                PhaseChanged = True
            End If

        Case "At Gate"
            'If all engines are shut down, we enter the Shutdown phase.
             If Not cPos.EnginesStarted Then
                info.Phase = "Shutdown"
                info.ShutdownTime = Now
                info.ShutdownFuel = cPos.Fuel
                info.ShutdownWeight = cPos.Weight
                frmMain.tmrFlightTimeCounter.Enabled = False
                PhaseChanged = True
            End If
    End Select
End Function

Public Sub CheckSimRate(minRate As Integer, maxRate As Integer)
    Dim lngResult As Long
    Dim newSimRate As Long

    'Check the simulator rate
    If (pos.simRate > (maxRate * 256)) Then
        ShowMessage "Reset Sim Rate to " + CStr(maxRate) + "x", ACARSERRORCOLOR
        newSimRate = maxRate * 256
    ElseIf (pos.simRate < (minRate * 256)) Then
        ShowMessage "Reset Sim Rate to " + CStr(minRate) + "x", ACARSERRORCOLOR
        newSimRate = minRate * 256
    End If
    
    If (newSimRate <> 0) Then
        Call FSUIPC_Write(&HC1A, 2, VarPtr(newSimRate), lngResult)
        Call FSUIPC_Process(lngResult)
    End If
End Sub

Private Function UpdateJetPosInterval() As Integer
    If ((info.Phase = "Preflight") Or (info.Phase = "At Gate") Or (info.Phase = "Shutdown")) Then
        UpdateJetPosInterval = 90
    ElseIf (info.Phase = "Pushback") Then
        UpdateJetPosInterval = 20
    ElseIf ((info.Phase = "Taxi Out") Or (info.Phase = "Taxi In")) Then
        UpdateJetPosInterval = 8
    ElseIf ((info.Phase = "Takeoff") Or (info.Phase = "Landed")) Then
        UpdateJetPosInterval = 5
    ElseIf (pos.GroundSpeed < 175) Then
        UpdateJetPosInterval = 7
    ElseIf (pos.GroundSpeed < 235) Then
        UpdateJetPosInterval = 10
    ElseIf (pos.GroundSpeed < 255) Then
        UpdateJetPosInterval = 15
    ElseIf (pos.GroundSpeed < 295) Then
        UpdateJetPosInterval = 25
    ElseIf (pos.GroundSpeed < 560) Then
        UpdateJetPosInterval = 75
    Else
        UpdateJetPosInterval = 65
    End If
End Function

Private Function UpdateTurbopropPosInterval() As Integer
    If ((info.Phase = "Preflight") Or (info.Phase = "At Gate") Or (info.Phase = "Shutdown")) Then
        UpdateTurbopropPosInterval = 90
    ElseIf (info.Phase = "Pushback") Then
        UpdateTurbopropPosInterval = 20
    ElseIf ((info.Phase = "Taxi Out") Or (info.Phase = "Taxi In")) Then
        UpdateTurbopropPosInterval = 8
    ElseIf ((info.Phase = "Takeoff") Or (info.Phase = "Landed")) Then
        UpdateTurbopropPosInterval = 4
    ElseIf (pos.GroundSpeed < 175) Then
        UpdateTurbopropPosInterval = 9
    ElseIf (pos.GroundSpeed < 235) Then
        UpdateTurbopropPosInterval = 15
    ElseIf (pos.GroundSpeed < 255) Then
        UpdateTurbopropPosInterval = 30
    Else
        UpdateTurbopropPosInterval = 75
    End If
End Function

Public Function UpdatePositionInterval() As Integer
    If isTurboProp Then
        UpdatePositionInterval = UpdateTurbopropPosInterval
    Else
        UpdatePositionInterval = UpdateJetPosInterval
    End If
End Function
