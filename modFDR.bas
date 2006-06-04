Attribute VB_Name = "modFDR"
Option Explicit

Private EngineType As Byte
Private EngineCount As Integer

'Engine type constants
Private Const PISTON = 0
Private Const JET = 1
Private Const TURBOPROP = 5

'Fuel tank offsets
Private tankOffsets As Variant
Private tankSizeOffsets As Variant

'Flight phase constants
Public Const UNKNOWN = 0
Public Const PREFLIGHT = 1
Public Const PushBack = 2
Public Const TAXI_OUT = 3
Public Const TAKEOFF = 4
Public Const AIRBORNE = 5
Public Const ROLLOUT = 6
Public Const TAXI_IN = 7
Public Const ATGATE = 8
Public Const SHUTDOWN = 9
Public Const COMPLETE = 10
Public Const ABORTED = 11
Public Const ERROR = 12

'Flight status bit position constants.
Public Const FLIGHTPAUSED = &H1&
Public Const FLIGHTSLEWING = &H2&
Public Const FLIGHTPARKED = &H4&
Public Const FLIGHTONGROUND = &H8&

Public Const FLIGHT_SP_ARM = &H10&
Public Const FLIGHT_GEAR_DOWN = &H20&
Public Const FLIGHT_AFTERBURNER = &H40&

Public Const FLIGHT_AP_GPS = &H100&
Public Const FLIGHT_AP_NAV = &H200&
Public Const FLIGHT_AP_HDG = &H400&
Public Const FLIGHT_AP_APR = &H800&
Public Const FLIGHT_AP_ALT = &H1000&
Public Const FLIGHT_AT_IAS = &H2000&
Public Const FLIGHT_AT_MACH = &H4000&
Public Const FLIGHTPUSHBACK = &H8000&

'FS controls
Private Const KEY_PANEL_ID_TOGGLE = 66506
Private Const KEY_PANEL_ID_OPEN = 66507
Private Const KEY_PANEL_ID_CLOSE = 66508

'Load/Save flight codes
Public Const FLIGHT_LOAD = 0
Public Const FLIGHT_SAVE = 1

Private Type FSCONTROL
    ID As Long
    value As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal _
lpClassName As String, ByVal lpWindowName As String) As Long

Public Function RecordFlightData(aInfo As AircraftInfo) As PositionData

Dim lngResult As Long
Dim x As Integer

Dim curLat As Currency
Dim curLon As Currency

Dim terrElevation As Long
Dim altMSL As Long

Dim magVar As Integer
Dim hdg As Long

Dim ASpeed As Long
Dim GSpeed As Long
Dim VSpeed As Long
Dim TSpeed As Long
Dim Mach As Long
Dim Gs As Integer
Dim AofA As Integer

Dim Bank As Long
Dim Pitch As Long

Dim flapsL As Integer
Dim flapsR As Integer
Dim Spoilers As Long
Dim gearPos As Long

Dim RawSimRate As Integer
Dim isPaused As Integer
Dim isFrozen As Byte
Dim isPushBack As Long
Dim isSlew As Integer
Dim isReplay As Long
Dim isCockpitView As Byte

'Make sure we're connected to FSUIPC
If Not config.FSUIPCConnected Then Exit Function

'Set critical error handler
On Error GoTo FatalError

'Initialize the fuel tank readers
If IsEmpty(tankOffsets) Then InitFuelCheck

'Get latitude/longitude
10 Call FSUIPC_Read(&H560, 8, VarPtr(curLat), lngResult)
Call FSUIPC_Read(&H568, 8, VarPtr(curLon), lngResult)

'Get terrain elevation/altitude MSL/heading/magVariation/Gs/AoA
Call FSUIPC_Read(&H20, 4, VarPtr(terrElevation), lngResult)
Call FSUIPC_Read(&H574, 4, VarPtr(altMSL), lngResult)
Call FSUIPC_Read(&H2A0, 2, VarPtr(magVar), lngResult)
Call FSUIPC_Read(&H580, 4, VarPtr(hdg), lngResult)
Call FSUIPC_Read(&H11BA, 2, VarPtr(Gs), lngResult)
Call FSUIPC_Read(&H11BE, 2, VarPtr(AofA), lngResult)

'Get air speed/ground speed/vertical speed/Mach
Call FSUIPC_Read(&H2BC, 4, VarPtr(ASpeed), lngResult)
Call FSUIPC_Read(&H2B4, 4, VarPtr(GSpeed), lngResult)
Call FSUIPC_Read(&H2C8, 4, VarPtr(VSpeed), lngResult)
Call FSUIPC_Read(&H30C, 4, VarPtr(TSpeed), lngResult)
Call FSUIPC_Read(&H11C6, 2, VarPtr(Mach), lngResult)

    
'Get left/right inboard flaps positions/spoilers/gear positions
Call FSUIPC_Read(&H30F0, 2, VarPtr(flapsL), lngResult)
Call FSUIPC_Read(&H30F4, 2, VarPtr(flapsR), lngResult)
Call FSUIPC_Read(&HBD0, 4, VarPtr(Spoilers), lngResult)
Call FSUIPC_Read(&HBE8, 4, VarPtr(gearPos), lngResult)

'Get sim rate/pause/slew
20 Call FSUIPC_Read(&HC1A, 2, VarPtr(RawSimRate), lngResult)
Call FSUIPC_Read(&H264, 2, VarPtr(isPaused), lngResult)
Call FSUIPC_Read(&H3365, 1, VarPtr(isFrozen), lngResult)
Call FSUIPC_Read(&H5DC, 2, VarPtr(isSlew), lngResult)
Call FSUIPC_Read(&H628, 4, VarPtr(isReplay), lngResult)
Call FSUIPC_Read(&H8320, 1, VarPtr(isCockpitView), lngResult)
Call FSUIPC_Read(&H31F0, 4, VarPtr(isPushBack), lngResult)

'Get pitch/bank
Call FSUIPC_Read(&H578, 4, VarPtr(Pitch), lngResult)
Call FSUIPC_Read(&H57C, 4, VarPtr(Bank), lngResult)

Dim apMode As Long
Dim GPSmode As Long
Dim NAVmode As Long
Dim HDGmode As Long
Dim APRmode As Long
Dim ALTmode As Long
Dim IASmode As Long
Dim MCHmode As Long

'Get AP/NAV/HDG/APR/ALT/IAS/MCH autopilot modes
Call FSUIPC_Read(&H132C, 4, VarPtr(GPSmode), lngResult)
Call FSUIPC_Read(&H7C4, 4, VarPtr(NAVmode), lngResult)
Call FSUIPC_Read(&H800, 4, VarPtr(APRmode), lngResult)
Call FSUIPC_Read(&H7DC, 4, VarPtr(IASmode), lngResult)
Call FSUIPC_Read(&H7E4, 4, VarPtr(MCHmode), lngResult)

'Read PMDG 737 offsets
'If aInfo.PMDG737 Then
'    Call FSUIPC_Read(&H6226, 2, VarPtr(apMode), lngResult)
'    Call FSUIPC_Read(&H622C, 2, VarPtr(HDGmode), lngResult)
'    Call FSUIPC_Read(&H622E, 2, VarPtr(ALTmode), lngResult)
'Else
    Call FSUIPC_Read(&H7BC, 4, VarPtr(apMode), lngResult)
    Call FSUIPC_Read(&H7C8, 4, VarPtr(HDGmode), lngResult)
    Call FSUIPC_Read(&H7D0, 4, VarPtr(ALTmode), lngResult)
'End If

'Get fuel remaining
Dim FuelWeight As Integer
Dim ZeroFuelWeight As Long
Dim TankCapacity(10) As Long
Dim TankLevel(10) As Long

'Get tank capacity and sizes
For x = 0 To UBound(tankOffsets)
    Call FSUIPC_Read(CLng(tankOffsets(x)), 4, VarPtr(TankLevel(x)), lngResult)
    Call FSUIPC_Read(CLng(tankSizeOffsets(x)), 4, VarPtr(TankCapacity(x)), lngResult)
Next

'Get fuel/ZF weight
Call FSUIPC_Read(&HAF4, 2, VarPtr(FuelWeight), lngResult)
Call FSUIPC_Read(&H3BFC, 4, VarPtr(ZeroFuelWeight), lngResult)

Dim EngineN1(3) As Double
Dim EngineN2(3) As Double
Dim EngineThrottle(3) As Integer
Dim EngineRunning(3) As Integer
Dim EngineAB(3) As Long
Dim FuelFlow(3) As Double

'Get engine type
30 Call FSUIPC_Read(&H609, 1, VarPtr(EngineType), lngResult)

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
    Call FSUIPC_Read(&H88C + (&H98 * x), 2, VarPtr(EngineThrottle(x)), lngResult)
    Call FSUIPC_Read(&H918 + (&H98 * x), 8, VarPtr(FuelFlow(x)), lngResult)

    'Read turbine N1/N2
    If (EngineType <> PISTON) Then
        Call FSUIPC_Read(&H2000 + (&H100 * x), 8, VarPtr(EngineN1(x)), lngResult)
        Call FSUIPC_Read(&H2008 + (&H100 * x), 8, VarPtr(EngineN2(x)), lngResult)
        If (EngineType = JET) Then Call FSUIPC_Read(&H2048 + (&H100 * x), 4, VarPtr(EngineAB(x)), lngResult)
    Else
        Call FSUIPC_Read(&H2408 + (&H100 * x), 8, VarPtr(EngineN1(x)), lngResult)
    End If
Next

Dim isOnGround As Integer
Dim isParkBrake As Integer

'Is the aircraft on the ground/parked?
40 Call FSUIPC_Read(&H366, 2, VarPtr(isOnGround), lngResult)
Call FSUIPC_Read(&HBC8, 2, VarPtr(isParkBrake), lngResult)

Dim WindSpeed As Integer
Dim WindHeading As Long
Dim sWindCeiling As Integer

'Get wind data
Call FSUIPC_Read(&HE90, 2, VarPtr(WindSpeed), lngResult)
Call FSUIPC_Read(&HE92, 2, VarPtr(WindHeading), lngResult)
Call FSUIPC_Read(&HEEE, 2, VarPtr(sWindCeiling), lngResult)

'Call FSUIPC and create the flight data bean
Dim data As New PositionData
data.phase = info.PhaseName
If Not FSUIPC_Process(lngResult) Then
    FSError lngResult
    Exit Function
End If

'Calculate latitude/longitude
60 data.Latitude = (curLat * 10000#) * 90# / (10001750# * 65536# * 65536#)
data.Longitude = (curLon * 10000#) * 360# / (65536# * 65536# * 65536# * 65536#)

'Calculate speeds
data.AirSpeed = CInt(ASpeed / 128#)
data.GroundSpeed = CInt(GSpeed * 3600# / 65536# / 1852#)
61 data.VerticalSpeed = CInt(VSpeed * 60# * 3.28084 / 256#)
data.Mach = CDbl(Mach / 20480#)
If (isOnGround = 1) Then
    Dim tmpTSpeed As Long
62  tmpTSpeed = CLng(TSpeed * 60# * 3.28084 / 256#)
    If (Abs(tmpTSpeed) < 32000) Then data.TouchdownSpeed = tmpTSpeed
End If

'Calculate G-force/Angle of Attack
data.GForce = Gs / 625#
data.AngleOfAttack = 100 - (100# * AofA / 32767)

'Calculate Altitude AGL/MSL
65 terrElevation = (terrElevation * 3.28084) / 256
data.AltitudeMSL = altMSL * 3.28084
data.AltitudeAGL = data.AltitudeMSL - terrElevation

'Calculate heading
66 magVar = magVar * 360# / 65536#
data.Heading = CInt(hdg * (360# / (65536# * 65536#)))
data.Heading = data.Heading - magVar
If (data.Heading < 0) Then data.Heading = (360 + data.Heading)

'Get wind data
67 data.WindHeading = CInt(WindHeading * 360# / 65536#)
data.WindSpeed = WindSpeed

'Adjust surface wind heading if magnetic and not true
68 If (data.AltitudeMSL > sWindCeiling) Then
    data.WindHeading = data.WindHeading - magVar
    If (data.WindHeading < 0) Then data.WindHeading = (360 + data.WindHeading)
End If

'Calculate flaps setting
data.Flaps = (flapsL / 512) + (flapsR / 512)

'Calculate pitch/bank
data.Pitch = Round(Pitch * 360# / (-65536# * 65536#), 2)
data.Bank = Round(Bank * 360# / (-65536# * 65536#), 2)

'Calculate total fuel remaining in all tanks.
Dim FuelGallons As Double
Dim TankPct As Double
70 For x = 0 To UBound(TankCapacity)
    If TankCapacity(x) > 0 Then
        TankPct = TankLevel(x) / 128# / 65536#
        FuelGallons = FuelGallons + (TankPct * TankCapacity(x))
    End If
Next

'Get fuel/total weight
data.fuel = FuelGallons * (FuelWeight / 256)
data.weight = (ZeroFuelWeight / 256) + data.fuel

'Count how many engines are running and calculate average N1/N2
80 data.EngineCount = EngineCount
Dim EngineProp As Long
For x = 0 To 3
    If EngineRunning(x) Then
        data.EnginesStarted = True
        data.setFuelFlow x, CLng(FuelFlow(x))
        data.setN1 x, EngineN1(x)
        If ((EngineType = JET) Or (EngineType = TURBOPROP)) Then
            data.setN2 x, EngineN2(x)
            If aInfo.HasAfterburner Then data.AfterBurner = data.AfterBurner Or (EngineAB(x) = 1)
        Else
            data.setN2 x, 0
        End If
    Else
        data.setN1 x, 0
        data.setN2 x, 0
    End If
    
    data.setThrottle x, (EngineThrottle(x) / 163.84)
Next

'Load cockpit view
data.CockpitView = (isCockpitView < 3)

'Build flags
Dim isAP As Boolean
isAP = (apMode <> 0)
data.Paused = (isPaused = 1) Or (isReplay = 1) Or (isFrozen > 0)
data.Slewing = (isSlew = 1)
data.Parked = (isParkBrake = 32767)
data.onGround = (isOnGround = 1)
data.Spoilers = (Spoilers = 4800)
data.GearDown = (gearPos > 8192)
data.PushBack = (isPushBack <> 3)
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
ShowMessage err.Description + " at Line " + CStr(Erl) + " of FDR", ACARSERRORCOLOR
Resume ExitFunc

End Function

Private Sub ShowWeight(ByVal weight As Long, ByVal fuel As Long)
    ShowMessage "Total Weight: " & Format(weight, "#,##0") & " lbs, Total Fuel: " & _
        Format(fuel, "#,##0") & " lbs", ACARSTEXTCOLOR
End Sub

Public Function PhaseChanged(cPos As PositionData) As Boolean
    Static TakeoffCheckCount As Integer

    Select Case info.FlightPhase
        Case PREFLIGHT
            'If the parking brake is released, we enter the Pushback phase.
            If cPos.PushBack Then
                info.FlightPhase = PushBack
                TakeoffCheckCount = 0
                PhaseChanged = True
                ShowWeight cPos.weight, cPos.fuel
                ShowFSMessage "Pushing Back", True, 5
            ElseIf (cPos.GroundSpeed > 3) Then
                info.FlightPhase = TAXI_OUT
                TakeoffCheckCount = 0
                info.TaxiOutTime.LocalTime = Now
                info.TaxiFuel = cPos.fuel
                info.TaxiWeight = cPos.weight
                PhaseChanged = True
                ShowFSMessage "Starting Taxi", True, 5
            End If
        
        Case PushBack
            'If we are moving forward, we enter the Taxi Out phase.
            If Not cPos.PushBack Then
                info.FlightPhase = TAXI_OUT
                info.TaxiOutTime.LocalTime = Now
                info.TaxiFuel = cPos.fuel
                info.TaxiWeight = cPos.weight
                PhaseChanged = True
                ShowFSMessage "Starting Taxi", True, 5
            End If

        Case TAXI_OUT
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
                info.FlightPhase = TAKEOFF
                info.TakeoffTime.LocalTime = Now
                If config.SB3Connected Then SB3Transponder True
                PhaseChanged = True
                ShowFSMessage "Takeoff Detected", False, 5
            ElseIf Not cPos.onGround And (cPos.AltitudeAGL > 15) And (cPos.AirSpeed > 45) Then
                info.FlightPhase = AIRBORNE
                info.TakeoffTime.LocalTime = Now
                If config.SB3Connected Then SB3Transponder True
                info.TakeoffSpeed = cPos.AirSpeed
                info.TakeoffFuel = cPos.fuel
                info.TakeoffWeight = cPos.weight
                info.TakeoffN1 = cPos.AverageN1
                ShowWeight cPos.weight, cPos.fuel
                ShowFSMessage "LIFTOFF at " & CStr(cPos.AirSpeed) & " knots", False, 8
                PhaseChanged = True
            End If

        Case TAKEOFF
            'If we're off the ground, then we enter the Airborne phase.
            If Not cPos.onGround Then
                info.FlightPhase = AIRBORNE
                info.TakeoffSpeed = cPos.AirSpeed
                info.TakeoffFuel = cPos.fuel
                info.TakeoffWeight = cPos.weight
                info.TakeoffN1 = cPos.AverageN1
                info.TakeoffTime.LocalTime = Now
                ShowWeight cPos.weight, cPos.fuel
                ShowFSMessage "LIFTOFF at " & CStr(cPos.AirSpeed) & " knots", False, 8
                PhaseChanged = True
            ElseIf ((cPos.AirSpeed < 60) And (cPos.AverageThrottle < 30)) Then
                info.FlightPhase = TAXI_OUT
                If config.SB3Connected Then SB3Transponder False
                PhaseChanged = True
            End If
            
        Case AIRBORNE
            'If we've been in the Airborne phase for more than 20 seconds, and now
            'we're back on the ground, then we enter the Landed phase. The 20 seconds
            'is for debouncing.
            If cPos.onGround Then
                info.FlightPhase = ROLLOUT
                info.LandingTime.LocalTime = Now
                info.LandingSpeed = cPos.AirSpeed
                info.LandingVSpeed = cPos.TouchdownSpeed
                info.LandingFuel = cPos.fuel
                info.LandingWeight = cPos.weight
                info.LandingN1 = cPos.AverageN1
                ShowWeight cPos.weight, cPos.fuel
                ShowMessage "Touchdown speed: " & Format(cPos.TouchdownSpeed, "##0.0") & _
                    " feet/minute", ACARSTEXTCOLOR
                ShowFSMessage "TOUCHDOWN at " & Format(cPos.TouchdownSpeed, "##0.0") & _
                    " feet/minute", False, 10
                PhaseChanged = True
            End If
            
        Case ROLLOUT
            'If ground speed falls below 30 knots, we enter the Taxi In phase.
            If (cPos.GroundSpeed < 30) Then
                info.FlightPhase = TAXI_IN
                info.TaxiInTime.LocalTime = Now
                If config.SB3Connected Then SB3Transponder False
                PhaseChanged = True
            ElseIf (Not cPos.onGround And (cPos.AltitudeAGL > 7)) Then
                info.FlightPhase = AIRBORNE
                If config.SB3Connected Then SB3Transponder True
                PhaseChanged = True
            End If

        Case TAXI_IN
            'If parking brake is set, we enter the "At Gate" phase.
            If cPos.Parked Then
                info.FlightPhase = ATGATE
                info.GateTime.LocalTime = Now
                info.GateFuel = cPos.fuel
                info.GateWeight = cPos.weight
                info.FlightData = True
                PhaseChanged = True
            End If

        Case ATGATE
            'If all engines are shut down, we enter the Shutdown phase.
             If Not cPos.EnginesStarted Then
                info.FlightPhase = SHUTDOWN
                info.ShutdownTime.LocalTime = Now
                info.ShutdownFuel = cPos.fuel
                info.ShutdownWeight = cPos.weight
                ShowWeight cPos.weight, cPos.fuel
                ShowFSMessage "Engines shutdown", True, 5
                PhaseChanged = True
            ElseIf Not cPos.Parked Then
                info.FlightPhase = TAXI_IN
                PhaseChanged = True
            End If
    End Select
End Function

Public Sub CheckSimRate(minRate As Integer, maxRate As Integer)
    Dim lngResult As Long
    Dim newSimRate As Long
    Dim x As Integer
    Dim doProcess As Boolean
    
    'If we are replaying or paused, then skip this
    If (pos.Paused Or pos.Slewing) Then Exit Sub

    'Check the simulator rate
    doProcess = False
    If (pos.simRate > (maxRate * 256)) Then
        ShowMessage "Reset Sim Rate to " + CStr(maxRate) + "x", ACARSERRORCOLOR
        newSimRate = maxRate * 256
        Call FSUIPC_Write(&HC1A, 2, VarPtr(newSimRate), lngResult)
        doProcess = True
    ElseIf (pos.simRate < (minRate * 256)) Then
        ShowMessage "Reset Sim Rate to " + CStr(minRate) + "x", ACARSERRORCOLOR
        newSimRate = minRate * 256
        Call FSUIPC_Write(&HC1A, 2, VarPtr(newSimRate), lngResult)
        doProcess = True
    End If
    
    'Check slew mode
    If (pos.Slewing And Not pos.onGround) Then
        Dim slewEnabled As Integer
        slewEnabled = 0
        ShowMessage "Disabling Slew Mode when Airborne", ACARSERRORCOLOR
        Call FSUIPC_Write(&H5DC, 2, VarPtr(slewEnabled), lngResult)
        doProcess = True
    End If
    
    'Update FS via FSUIPC
    If doProcess Then Call FSUIPC_Process(lngResult)
End Sub

Public Sub SaveFlight()
    Dim fName As String
    Dim dwResult As Long
    Dim x As Integer
    Dim ctlCodes As Variant, fsCtls() As FSCONTROL
    
    'Only save on FS9
    If Not config.FSUIPCConnected Or (info.FSVersion < 7) Then Exit Sub

    'Determine the file name
    fName = "ACARS Flight " + SavedFlightID(info) + Chr(0)
    
    'Open the panels if using the Payne panel - if not in cockpit view, abort
    If acInfo.Payne7X7 Then
        If Not pos.CockpitView Then
            If config.ShowDebug Then ShowMessage "Not in Cockpit view - aborting LP panel save", DEBUGTEXTCOLOR
            Exit Sub
        End If
    
        ctlCodes = Array(10, 27, 41, 50, 111, 123, 182)
        ReDim fsCtls(UBound(ctlCodes))
        For x = 0 To UBound(ctlCodes)
            fsCtls(x).ID = KEY_PANEL_ID_OPEN
            fsCtls(x).value = ctlCodes(x)
            If config.ShowDebug Then ShowMessage "Opening Panel " + CStr(ctlCodes(x)), DEBUGTEXTCOLOR
            Call FSUIPC_Write(&H3110, 8, VarPtr(fsCtls(x)), dwResult)
        Next
    
        'Call FSUIPC
        If Not FSUIPC_Process(dwResult) Then ShowMessage "Error restoring Lonny Payne panel", _
            ACARSERRORCOLOR
    End If
    
    'Save the flight
    Call FSUIPC_WriteS(&H3F04, Len(fName), fName, dwResult)
    Call FSUIPC_Write(&H3F00, 2, VarPtr(FLIGHT_SAVE), dwResult)
    If Not FSUIPC_Process(dwResult) Then ShowMessage "Error Saving Flight", ACARSERRORCOLOR
    If config.ShowDebug Then ShowMessage "Saved Flight via FSUIPC", DEBUGTEXTCOLOR
    
    'Close the panels again
    If acInfo.Payne7X7 Then
        ctlCodes = Array(10, 41, 50, 111, 123, 182)
        ReDim fsCtls(UBound(ctlCodes))
        For x = 0 To UBound(ctlCodes)
            fsCtls(x).ID = KEY_PANEL_ID_CLOSE
            fsCtls(x).value = ctlCodes(x)
            If config.ShowDebug Then ShowMessage "Hiding Panel " + CStr(ctlCodes(x)), DEBUGTEXTCOLOR
            Call FSUIPC_Write(&H3110, 8, VarPtr(fsCtls(x)), dwResult)
        Next
        
        'Call FSUIPC
        If Not FSUIPC_Process(dwResult) Then ShowMessage "Error closing Lonny Payne panel", ACARSERRORCOLOR
    End If
End Sub

Public Sub RestoreFlight()
    Dim fName As String
    Dim dwResult As Long
    
    'Only restore on FS9
    If (info.FSVersion < 7) Then Exit Sub
    
    'Determine the file name
    If (info.FlightID > 0) Then
        fName = "ACARS Flight" + Format(info.FlightID, "000000") + Chr(0)
    Else
        fName = "ACARS Flight" + Format(info.StartTime.UTCTime, "yyyymmddhh") + Chr(0)
    End If
        
    'Restore the flight
    Call FSUIPC_WriteS(&H3F04, Len(fName), fName, dwResult)
    Call FSUIPC_Write(&H3F00, 2, VarPtr(FLIGHT_LOAD), dwResult)
    If Not FSUIPC_Process(dwResult) Then ShowMessage "Error Loading Flight", ACARSERRORCOLOR
    If config.ShowDebug Then ShowMessage "Loaded Flight via FSUIPC", DEBUGTEXTCOLOR
End Sub

Private Function UpdateJetPosInterval() As Integer
    If ((info.FlightPhase = PREFLIGHT) Or (info.FlightPhase = ATGATE) Or (info.FlightPhase = SHUTDOWN)) Then
        UpdateJetPosInterval = 90
    ElseIf (info.FlightPhase = PushBack) Then
        UpdateJetPosInterval = 20
    ElseIf ((info.FlightPhase = TAXI_OUT) Or (info.FlightPhase = TAXI_IN)) Then
        UpdateJetPosInterval = 10
    ElseIf ((info.FlightPhase = TAKEOFF) Or (info.FlightPhase = ROLLOUT)) Then
        UpdateJetPosInterval = 6
    ElseIf (pos.GroundSpeed < 175) Then
        UpdateJetPosInterval = 9
    ElseIf (pos.GroundSpeed < 235) Then
        UpdateJetPosInterval = 10
    ElseIf (pos.GroundSpeed < 255) Then
        UpdateJetPosInterval = 14
    ElseIf (pos.GroundSpeed < 295) Then
        UpdateJetPosInterval = 20
    ElseIf (pos.GroundSpeed < 560) Then
        UpdateJetPosInterval = 60
    Else
        UpdateJetPosInterval = 50
    End If
End Function

Private Function UpdateTurbopropPosInterval() As Integer
    If ((info.FlightPhase = PREFLIGHT) Or (info.FlightPhase = ATGATE) Or (info.FlightPhase = SHUTDOWN)) Then
        UpdateTurbopropPosInterval = 90
    ElseIf (info.FlightPhase = PushBack) Then
        UpdateTurbopropPosInterval = 20
    ElseIf ((info.FlightPhase = TAXI_OUT) Or (info.FlightPhase = TAXI_IN)) Then
        UpdateTurbopropPosInterval = 10
    ElseIf ((info.FlightPhase = TAKEOFF) Or (info.FlightPhase = ROLLOUT)) Then
        UpdateTurbopropPosInterval = 6
    ElseIf (pos.GroundSpeed < 175) Then
        UpdateTurbopropPosInterval = 10
    ElseIf (pos.GroundSpeed < 235) Then
        UpdateTurbopropPosInterval = 20
    ElseIf (pos.GroundSpeed < 255) Then
        UpdateTurbopropPosInterval = 50
    Else
        UpdateTurbopropPosInterval = 75
    End If
End Function

Public Function UpdatePositionInterval() As Integer
    If ((EngineType = PISTON) Or (EngineType = TURBOPROP)) Then
        UpdatePositionInterval = UpdateTurbopropPosInterval
    Else
        UpdatePositionInterval = UpdateJetPosInterval
    End If
End Function

Public Function GetAircraftInfo() As AircraftInfo
    Dim airInfo As New AircraftInfo
    Dim AIRBytes(255) As Byte, FSBytes(255) As Byte
    Dim pAlias As String
    Dim dwResult As Long, x As Integer
    
    Dim PMDGAirNames As Variant
    PMDGAirNames = Array("b737-600.air", "b737-700.air", "b737-800.air", "b737-900.air", _
        "b737-700wl.air", "b737-800wl.air")
    
    'Read the air file/fs path
    Call FSUIPC_Read(&H3E00, 256, VarPtr(FSBytes(0)), dwResult)
    Call FSUIPC_Read(&H3C00, 256, VarPtr(AIRBytes(0)), dwResult)
    If Not FSUIPC_Process(dwResult) Then
        ShowMessage "Error detecting Aircraft type", ACARSERRORCOLOR
        Set GetAircraftInfo = airInfo
        Exit Function
    End If
    
    'Parse the null-terminated strings
    airInfo.FSPath = BytesToStr(FSBytes)
    airInfo.AirPath = airInfo.FSPath + BytesToStr(AIRBytes)
    
    'Check for PMDG 737
    For x = 0 To UBound(PMDGAirNames)
        If (LCase(airInfo.AIRFile) = PMDGAirNames(x)) Then
            airInfo.PMDG737 = True
            Exit For
       End If
    Next x
    
    'Check for Lonny Payne panel
    pAlias = UCase(ReadINI("fltsim", "alias", "", airInfo.AircraftPath + "panel\panel.cfg"))
    airInfo.Payne7X7 = ((pAlias = "FSFSCONV\PANEL.JET.767") Or (pAlias = "FSFSCONV\PANEL.JET.757"))
    
    'Check for afterburner-equipped aircraft
    airInfo.HasAfterburner = (ReadINI("TurbineEngineData", "afterburner_available", "0", airInfo.CFGFile) = "1")
    If airInfo.HasAfterburner Then ShowMessage "Afterburner detected", ACARSTEXTCOLOR
    
    'Load the ICAO/IATA code
    airInfo.Code = UCase(ReadINI("General", "atc_model", "", airInfo.CFGFile))
    
    'Display conditions
    If config.ShowDebug Then
        ShowMessage "FDE = " + airInfo.AIRFile, DEBUGTEXTCOLOR
        If airInfo.Payne7X7 Then ShowMessage "Detected Lonny Payne 757/767 panel", ACARSTEXTCOLOR
        If airInfo.PMDG737 Then ShowMessage "Detected PMDG 737 package", ACARSTEXTCOLOR
        If airInfo.PMDG747 Then ShowMessage "Detected PMDG 747 package", ACARSTEXTCOLOR
        If airInfo.LDS767 Then ShowMessage "Detected Level D 767 package", ACARSTEXTCOLOR
    End If
    
    'Return the aircraft info
    Set GetAircraftInfo = airInfo
End Function

Public Sub InitFuelCheck()
    'Init the FSUIPC offset arrays
    tankOffsets = Array(&HB74, &HB7C, &HB84, &HB8C, &HB94, &HB9C, &HBA4, &H1244, &H124C, &H1254, &H125C)
    tankSizeOffsets = Array(&HB78, &HB80, &HB88, &HB90, &HB98, &HBA0, &HBA8, &H1248, &H1250, &H1258, &H1260)
End Sub

Private Function BytesToStr(ByRef chars() As Byte) As String
    Dim x As Integer
    Dim result As String
    
    For x = 0 To UBound(chars)
        If (chars(x) = 0) Then Exit For
        result = result + Chr(chars(x))
    Next
    
    BytesToStr = result
End Function

Public Function IsFSRunning() As Boolean
    IsFSRunning = (FindWindow("FS98MAIN", vbNullString) <> 0)
End Function

Public Function IsFSReady() As Boolean
    Dim isReady As Byte, inMenu As Byte
    Dim dwResult As Long
    
    'Check if we're connected
    If Not (IsFSRunning() And config.FSUIPCConnected) Then Exit Function
    
    'Check "Ready to Fly" and "Modal Dialog" offsets
    Call FSUIPC_Read(&H3364, 1, VarPtr(isReady), dwResult)
    Call FSUIPC_Read(&H3365, 1, VarPtr(inMenu), dwResult)
    If Not FSUIPC_Process(dwResult) Then
        FSUIPC_Close
        MsgBox "Error querying Microsoft Flight Simulator", vbError + vbOKOnly, "IsFSReady FSUIPC Error"
        Exit Function
    End If

    IsFSReady = (isReady = 0) And (inMenu = 0)
End Function
