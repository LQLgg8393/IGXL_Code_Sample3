Attribute VB_Name = "VBT_A_Support"
Option Explicit


Private Const m_SETTLEWAIT_dbl As Double = 0.001   ' new spec settling time in sec,
'1.0
Public Function htl_SetupDSSC( _
       ByVal in_DSSCPattern_PAT As Pattern, _
       ByVal in_DSSCPinGroup_PL As PinList, _
       ByVal in_SignalName_str As String, _
       ByVal in_SampleSize_lng As Long, _
       ByRef io_CapWave_PLD As PinListData) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBU_SpecUtil", "htl_SetupDSSC", in_DSSCPattern_PAT)

    '   Declarations
    Dim i_PinGroup_str() As String
    Dim i_Pin_str As Variant
    Dim i_Offline_DSP As New DSPWave

    '   Check
    If in_DSSCPinGroup_PL = "" Then Exit Function

    If TheExec.TesterMode = testModeOnline Then

        If in_DSSCPattern_PAT = "" Then Exit Function
        '   Setup DSSC, Signal, PinGroup and DSPWave
        With TheHdw.DSSC.Pins(in_DSSCPinGroup_PL) _
             .Pattern(in_DSSCPattern_PAT).Capture.Signals _
             .Add(in_SignalName_str)

            .SampleSize = in_SampleSize_lng
            .LoadSettings
            io_CapWave_PLD = .DSPWave
        End With
    Else
        '   simulate the DSPWave
        i_Offline_DSP.CreateConstant 1, in_SampleSize_lng, DspLong
        'io_CapWave_PLD.AddPin (in_DSSCPinGroup_PL)
        i_PinGroup_str = Split(CStr(in_DSSCPinGroup_PL), ",", , vbTextCompare)
        For Each i_Pin_str In i_PinGroup_str
            io_CapWave_PLD.AddPin (i_Pin_str)
            io_CapWave_PLD.Pins(i_Pin_str).Value = i_Offline_DSP
        Next
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


'2.0
Public Function MeasureFrequency( _
       ByVal in_PinList_PL As PinList, _
       ByRef io_Interval_dbl As Double, _
       ByRef out_Frequency_PLD As PinListData, _
       Optional ByVal in_source As Long = VOH, _
       Optional ByVal in_slop As Long = Positive) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBA_Common_NonVBT", "MeasureFrequency")

    With TheHdw.Digital.Pins(in_PinList_PL).FreqCtr
        .Enable = IntervalEnable
        .Clear
        .EventSource = in_source
        .EventSlope = in_slop
        ' set and recalculate the time window
        If io_Interval_dbl < 0.000001 Then
            'the interval setup in test instance indicate that _
             the frequency measurement accuracy will be > 1MHz, _
             are you sure?
            Stop
        End If
        .Interval = io_Interval_dbl
        io_Interval_dbl = .Interval

        .Start
        Call TheHdw.Wait(io_Interval_dbl * 1.1)
        out_Frequency_PLD = .Read.Math.Divide(io_Interval_dbl)
    End With
    If TheExec.TesterMode = testModeOffline Then
        out_Frequency_PLD = out_Frequency_PLD.Math.Add(100 * MHz)
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function PPMU_FIMV( _
       ByVal in_PinList_PL As PinList, _
       ByVal in_ForceI_dbl As Double, _
       ByRef out_MeasureV_PLD As PinListData, _
       Optional ByVal in_SettleWaitSec_dbl As Double = 0.001, _
       Optional ByVal in_SampleSize_lng As Long = 1, _
       Optional ByVal in_ConnectB4Meas_bool As Boolean = True, _
       Optional ByVal in_DisconnectAftMeas_bool As Boolean = True, _
       Optional ByVal in_Switch2PE_bool As Boolean = True, _
       Optional ByVal in_MeterRange_dbl As Double = 0.005) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBA_Common_NonVBT", "PPMU_FIMV")
    
    If TheExec.TesterMode = testModeOffline Then
        Dim i_pin As Variant
        Dim pin As Variant
        i_pin = Split(in_PinList_PL, ",")
        For Each pin In i_pin
            out_MeasureV_PLD.AddPin (pin)
            out_MeasureV_PLD.Pins(pin) = 0
        Next pin
    Else

        With TheHdw.PPMU.Pins(in_PinList_PL)
            If in_ConnectB4Meas_bool Then
                If in_Switch2PE_bool Then TheHdw.Digital.Pins(in_PinList_PL).Disconnect
                .ForceI 0
                .Gate = tlOn
                .Connect
            End If
            .ForceI in_ForceI_dbl, in_MeterRange_dbl
            TheHdw.Wait in_SettleWaitSec_dbl
            out_MeasureV_PLD = .Read(tlPPMUReadMeasurements, in_SampleSize_lng)
            If in_DisconnectAftMeas_bool Then
                .ForceI 0
                .Disconnect
                .Gate = tlOff
                If in_Switch2PE_bool Then TheHdw.Digital.Pins(in_PinList_PL).Connect
            End If
        End With
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function PPMU_FVMI( _
       ByVal in_PinList_PL As PinList, _
       ByVal in_ForceV_dbl As Double, _
       ByRef out_MeasureI_PLD As PinListData, _
       Optional ByVal in_SettleWaitSec_dbl As Double = 0.001, _
       Optional ByVal in_SampleSize_lng As Long = 1, _
       Optional ByVal in_ConnectB4Meas_bool As Boolean = True, _
       Optional ByVal in_DisconnectAftMeas_bool As Boolean = True, _
       Optional ByVal in_Switch2PE_bool As Boolean = True, _
       Optional ByVal in_MeterRange_dbl As Double = 0.005) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBA_Common_NonVBT", "PPMU_FVMI")

    If TheExec.TesterMode = testModeOffline Then
        out_MeasureI_PLD.AddPin (in_PinList_PL)
        out_MeasureI_PLD.Pins(in_PinList_PL) = 0
    Else
        With TheHdw.PPMU.Pins(in_PinList_PL)
            If in_ConnectB4Meas_bool Then
                If in_Switch2PE_bool Then TheHdw.Digital.Pins(in_PinList_PL).Disconnect
                .ForceI 0
                .Gate = tlOn
                .Connect
            End If
            .ForceV in_ForceV_dbl, in_MeterRange_dbl
            TheHdw.Wait in_SettleWaitSec_dbl
            out_MeasureI_PLD = .Read(tlPPMUReadMeasurements, in_SampleSize_lng)
            If in_DisconnectAftMeas_bool Then
                .ForceI 0
                .Disconnect
                .Gate = tlOff
                If in_Switch2PE_bool Then TheHdw.Digital.Pins(in_PinList_PL).Connect
            End If
        End With
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function PPMU_FV( _
       ByVal in_PinList_PL As PinList, _
       ByVal in_ForceV_dbl As Double, _
       Optional ByVal in_SettleWaitSec_dbl As Double = 0.001) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBA_Common_NonVBT", "PPMU_FV")

    With TheHdw.PPMU.Pins(in_PinList_PL)
        .ForceV in_ForceV_dbl
        '        If .IsConnected = False Then
        TheHdw.Digital.Pins(in_PinList_PL).Disconnect
        .Gate = tlOff
        .Connect
        .Gate = tlOn
        '        End If
        TheHdw.Wait in_SettleWaitSec_dbl
    End With

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PPMU_Disconnect( _
       ByVal in_PinList_PL As PinList, _
       Optional ByVal in_SwitchToPE_bool As Boolean = True) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBA_Common_NonVBT", "PPMU_Disconnect")

    With TheHdw.PPMU.Pins(in_PinList_PL)
        .ForceI 0
        TheHdw.Wait 0.000999
        .Disconnect
        If in_SwitchToPE_bool Then
            TheHdw.Digital.Pins(in_PinList_PL).InitState = chInitoff
            TheHdw.Digital.Pins(in_PinList_PL).Connect
        End If
    End With

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function U_MIPI_Switch_vbt( _
       in_Drv0Pins As PinList, _
       in_Drv1Pins As PinList, _
       Optional in_SettleWaitSec_dbl As Double = 0.001) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__Basic", "U_MIPI_Switch_vbt", TheExec.DataManager.InstanceName)
    
'''    Dim CurInstanceName As String
'''    CurInstanceName = TheExec.DataManager.InstanceName
    
'''    ' Run pattern while relay on
'''    If InStr(1, CurInstanceName, "_ON", vbTextCompare) > 1 Then
'''        TheHdw.PPMU.AllowPPMUFuncRelayConnection True, False
'''    End If
    

    If in_Drv0Pins.Value <> "" Then
        Call PPMU_FV(in_Drv0Pins, 0#, 0)
    End If
    If in_Drv1Pins.Value <> "" Then
        Call PPMU_FV(in_Drv1Pins, 3.3, 0)
    End If
    TheHdw.Wait in_SettleWaitSec_dbl
    
'''    If InStr(1, CurInstanceName, "_OFF", vbTextCompare) > 1 Then
'''        TheHdw.PPMU.Pins(in_Drv0Pins).Disconnect
'''        TheHdw.PPMU.Pins(in_Drv1Pins).Disconnect
'''        TheHdw.PPMU.AllowPPMUFuncRelayConnection False, False
'''    End If

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function U_UTILITY_Switch_vbt( _
       in_Util_OFF As PinList, _
       in_Util_ON As PinList, _
       Optional in_SettleWaitSec_dbl As Double = 0.001) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__Basic", "U_UTILITY_Switch_vbt", TheExec.DataManager.InstanceName)


    If in_Util_OFF.Value <> "" Then TheHdw.Utility.Pins(in_Util_OFF).State = tlUtilBitOff
    If in_Util_ON.Value <> "" Then TheHdw.Utility.Pins(in_Util_ON).State = tlUtilBitOn
    TheHdw.Wait in_SettleWaitSec_dbl

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

' below PA clock need to be updated!!!

'Public Function PA_CLK_Start_38M4_32K() As Long
'
'    On Error GoTo errHandler
'
'    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = True
'    TheHdw.Protocol.Ports("CLK_32K").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_32K").IdleWait
'
'    TheHdw.Protocol.Ports("CLK_38M4").Enabled = True
'    TheHdw.Protocol.Ports("CLK_38M4").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_38M4").IdleWait
'
'    TheHdw.Wait 0.001 '* 10
'    TheExec.Datalog.WriteComment "clk start"
'
'    Exit Function
'
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function PA_CLK_Stop_38M4_32K() As Long
'
'    On Error GoTo errHandler
'
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = False
'    TheHdw.Protocol.Ports("CLK_38M4").Enabled = False
'    TheHdw.Wait 0.001
'
'    Exit Function
'
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function PA_CLK_Start_32K_1PIN() As Long
'
'    On Error GoTo errHandler
'
'    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = True
'    TheHdw.Protocol.Ports("CLK_32K").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_32K").IdleWait
'
'    TheHdw.Wait 0.001
'
'    Exit Function
'
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'Public Function PA_CLK_STOP_32K_1PIN() As Long
'
'    On Error GoTo errHandler
'
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = False
'
'    TheHdw.Wait 0.001
'
'    Exit Function
'
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function PA_CLK_32K_Start_NoALT() 'argc As Long, argv() As String)
'
'    On Error GoTo errHandler
'
''''    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
''    TheHdw.Digital.Pins("gIO_ALL").InitState = chInitHi
'    TheHdw.Wait 0.001
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = True
'    TheHdw.Protocol.Ports("CLK_32K").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_32K").IdleWait
'
'    TheHdw.Wait 0.001 * 10
'
'    Exit Function
'
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function PA_CLK_32K_Stop() As Long
'
'    On Error GoTo errHandler
'
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = False
'    TheHdw.Wait 0.01
'
'    Exit Function
'
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function PA_CLK_Start_32K_1PIN_IPF(argc As Long, argv() As String) As Long
'
'    On Error GoTo errHandler
'
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = True
'    TheHdw.Protocol.Ports("CLK_32K").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_32K").IdleWait
'
'    TheHdw.Wait 0.001
'
'    Exit Function
'
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function PA_CLK_STOP_32K_1PIN_IPF(argc As Long, argv() As String) As Long
'
'    On Error GoTo errHandler
'
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = False
'    TheHdw.Wait 0.001
'
'    Exit Function
'
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function PA_CLK_Start_38M4_32K_IPF(argc As Long, argv() As String) As Long
'
'    On Error GoTo errHandler
'
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = True
'    TheHdw.Protocol.Ports("CLK_32K").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_32K").IdleWait
'
'    TheHdw.Protocol.Ports("CLK_38M4").Enabled = True
'    TheHdw.Protocol.Ports("CLK_38M4").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_38M4").IdleWait
'
'    TheHdw.Wait 0.001 '* 10
'    TheExec.Datalog.WriteComment "clk start"
'
'    Exit Function
'
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'Public Function PA_CLK_Stop_38M4_32K_IPF(argc As Long, argv() As String) As Long
'
'    On Error GoTo errHandler
'
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = False
'    TheHdw.Protocol.Ports("CLK_38M4").Enabled = False
'    TheHdw.Wait 0.001
'
'    Exit Function
'
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function



''''''''''''''

Public Function FRC_CLK_Start_32K_1PIN() As Long

    On Error GoTo errHandler

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    With TheHdw.Digital.Pins("CLK_32K").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 32000#
        .Start
    End With
    TheHdw.Wait 0.001

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function FRC_CLK_STOP_32K_1PIN() As Long

    On Error GoTo errHandler
    
    TheHdw.Digital.Pins("CLK_32K").FreeRunningClock.Stop

    TheHdw.Wait 0.001

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function FRC_CLK_Start_32K_1PIN_IPF(argc As Long, argv() As String) As Long

    On Error GoTo errHandler
    
    With TheHdw.Digital.Pins("CLK_32K").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 32000#
        .Start
    End With

    TheHdw.Wait 0.001

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function FRC_CLK_STOP_32K_1PIN_IPF(argc As Long, argv() As String) As Long

    On Error GoTo errHandler
    
    
    TheHdw.Digital.Pins("CLK_32K").FreeRunningClock.Stop

    TheHdw.Wait 0.001

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function FRC_CLK_Start_38M4_32K_IPF(argc As Long, argv() As String) As Long

    On Error GoTo errHandler

    With TheHdw.Digital.Pins("CLK_32K").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 32000#
        .Start
    End With
    
    With TheHdw.Digital.Pins("CLK_38M4").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 38400000#
        .Start
    End With

    TheHdw.Wait 0.001 '* 10
    TheExec.Datalog.WriteComment "clk start"

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function FRC_CLK_Stop_38M4_32K_IPF(argc As Long, argv() As String) As Long

    On Error GoTo errHandler
    
    TheHdw.Digital.Pins("CLK_32K").FreeRunningClock.Stop
    TheHdw.Digital.Pins("CLK_38M4").FreeRunningClock.Stop

    TheHdw.Wait 0.001

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function FRC_CLK_32K_Start_NoALT() 'argc As Long, argv() As String)

    On Error GoTo errHandler

'''    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'    TheHdw.Digital.Pins("gIO_ALL").InitState = chInitHi
'    TheHdw.Wait 0.001
    
    With TheHdw.Digital.Pins("CLK_32K").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 32000#
        .Start
    End With

    TheHdw.Wait 0.001 * 10

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
'

Public Function FRC_CLK_32K_Stop() As Long

    On Error GoTo errHandler

    TheHdw.Digital.Pins("CLK_32K").FreeRunningClock.Stop
    TheHdw.Wait 0.01

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

