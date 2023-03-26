Attribute VB_Name = "ZR_Core"
Option Explicit
'// ========================================================================
'// MIT License
'//
'// Copyright (c) 2023 John Rivett-Carnac +44 7887 570 669
'//
'// Permission is hereby granted, free of charge, to any person obtaining a copy
'// of this software and associated documentation files (the "Software"), to deal
'// in the Software without restriction, including without limitation the rights
'// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'// copies of the Software, and to permit persons to whom the Software is
'// furnished to do so, subject to the following conditions:
'//
'// The above copyright notice and this permission notice shall be included in all
'// copies or substantial portions of the Software.
'//
'// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'// SOFTWARE.
'// ========================================================================


    '// Singleton instance of the appliation timer object
    '// which has sole responsibility for setting
    '// the application timer, when needed.
    Public zr_appTimer As New ZR_clsAppTimer
    Public zr_bAppTimerCreated As Boolean
    
    
    '// Singleton instance of core object
    'Private xCore As zr_clsCore
    
    '// ---------------------------------------------------------------------
    '// The tick queue which holds all the waiting timers
    Public xTickQueue As zr_clsTickQueue
    
    
    '// The base date/timer from which all tick numbers are
    '// calculated.
    Private xBaseTime As Date
    Private xMaxTime As Date
    
    
    '// ============================================================================ Logger object
    Public zr_Log As New zr_clsLog

    
    Private xLastTimerID As Long '// Unique ID of the last timer that was instantiated
    
    
    '//------------------------------------------------ Stream monitor dictionary
    Private dStreamMonitors As New Dictionary
    
    
'// ================================================================================================================
'// Event stream monitor Management
'//

Public Function zr_StreamExists(streamLabel As String) As Boolean
'// Indicates if a stream monitor with the given label exists
    zr_StreamExists = dStreamMonitors.Exists(streamLabel)
End Function


Public Function zr_StreamMonitor(streamLabel As String) As ZR_clsStream
'// Returns the stream monitor with the given label. If it does not already
'// exist, then a new instance is created and added to the collection
R "zr_StreamMonitor " & streamLabel

    If Not dStreamMonitors.Exists(streamLabel) Then
        dStreamMonitors.Add streamLabel, New ZR_clsStream
        
        dStreamMonitors(streamLabel).streamLabel = streamLabel
    End If
    
    Set zr_StreamMonitor = dStreamMonitors(streamLabel)
L
End Function

Public Function zr_NextTimerID() As Long
'// Returns the next incremental timer ID. Used to assign a unique ID when a new timer object is created

    xLastTimerID = xLastTimerID + 1
    If xLastTimerID Mod 2 = 0 Then
        log "Even-numbered timer id"
    Else
        log "Odd-numbered timer id"
    End If
    zr_NextTimerID = xLastTimerID
    
End Function


'//===================================
Public Function zr_TickQueue() As zr_clsTickQueue

    If xTickQueue Is Nothing Then Set xTickQueue = New zr_clsTickQueue
    Set zr_TickQueue = xTickQueue
    
End Function

    
    
    
'// =================================================
'// This function returns the singleton core object through which
'// most of the interaction with the timers will be done. The core
'// object provides the means to pick up events from the core.
'Public Property Get zr_Core() As zr_clsCore
'    If xCore Is Nothing Then
'        zr_Start
'        xBaseTime = Now - FiveSeconds
'        MaxTime = zr_TimeFromTickNumber(MaxTickNumber)
'    End If
'    Set zr_Core = xCore
'End Property


'// -------------------------------------
Public Property Get zr_TicksPerSecond() As Long
    zr_TicksPerSecond = TicksPerSecond
End Property

'// ---------------------------------------------------------------------------
'// These functions control logging to the system log, which is by default the
'// immediate window. The user can stop logging to the immediate window and
'// implement their own logging function by intercepting the systemEvent
Public Function zr_StartLog(Optional logToDebug = False) As zr_clsLog
    
    If logToDebug Then zr_Log.sendToDebugStart
    zr_Log.startLog
    Set zr_StartLog = zr_Log
    
End Function

Public Sub zr_StopLog()
    zr_Log.stopLog
    
End Sub


'//========================================================
'// Is called initially to set some options. Need not be called by the user
Public Sub zr_Start()
    Static bStarted As Boolean
    If bStarted Then Exit Sub
    bStarted = True
    
R "zr_Start"

    '// Set the base time for calculating ticks
    zr_TickBaseTime
    
L
End Sub

'// Provides a means to force Continue of the timing loop, if it has failed, without resetting the project
'// You can call this from the debug window: zr_Core.zr_RestartAppTimer
Public Sub zr_RestartAppTimer()
    zr_appTimer.setAppTimer_forTick 0, True
End Sub

'// ==================================================================================
'// Create a new timer instance.
Public Function zr_NewTimer(Optional timeToExpiry As Date = 0, Optional expiryDateTime As Date = 0, Optional tickerInterval As Date = 0) As zr_clsTimer
'// Returns a newly-minted timer object. Normally invoked by one of the
'// clsTimerOptions functions .run or .zr_Start
R "zr_NewTimer"
    
    '// Create the new timer instance
    Dim oNewTimer As New zr_clsTimer
    
    '// if options have been supplied then configure the timer with the options
    oNewTimer.setConfig timeToExpiry, expiryDateTime, tickerInterval
    
    '// Return the new timer
    Set zr_NewTimer = oNewTimer
    
L
End Function



Public Function zr_SendLogToDebug(Optional TF As Boolean = True) As zr_clsLog
    If TF Then
        zr_Log.sendToDebugStart
    Else
        zr_Log.sendToDebugStop
    End If
    
    Set zr_SendLogToDebug = zr_Log
    
End Function

'// Calculates and sets the base time from which all tick numbers
'// are calculated.
Public Function zr_TickBaseTime() As Date
    
    If xBaseTime = 0 Then
        xBaseTime = Now - OneSecond
        
        
    End If
    
    zr_TickBaseTime = xBaseTime
    
End Function

Public Function zr_MaxTime() As Date
    If xMaxTime = 0 Then
        xMaxTime = zr_TimeFromTickNumber(MaxTickNumber)
    End If
    zr_MaxTime = xMaxTime
    
End Function

Public Function increment(sLabel As String) As String
    Dim s_n As String
    Dim i_n As Integer
    Dim s_l As String
    s_n = Right(sLabel, 1)
    If Len(sLabel) = 0 Then s_l = "" Else s_l = Left(sLabel, Len(sLabel) - 1)
    
    
    Select Case s_n
    Case 0 To 8
        i_n = CInt(s_n) + 1
        increment = s_l & i_n
    Case 9
        increment = increment(s_l) & "0"
    Case Else
        increment = sLabel & IIf(Right(sLabel, 1) = " ", "", " ") & "1"
    End Select
        
End Function


'// ---------------------------------------------------------------------
'//
'//     Functions for converting times to tick numbers and vice-versa
'//

Public Function zr_TickNumberFromTime(dTime As Date) As Long
    zr_TickNumberFromTime = (dTime - zr_TickBaseTime) * SecondsPerDay * TicksPerSecond
    
End Function


Public Function zr_TicksFromInterval(dIntervalTime As Date) As Long
    zr_TicksFromInterval = CLng(dIntervalTime * 86400) * TicksPerSecond
End Function
' / TicksPerSecond

Public Function zr_TimeFromTickNumber(iTickNumber As Long) As Date
    zr_TimeFromTickNumber = zr_TickBaseTime + (iTickNumber * OneSecond) / TicksPerSecond
End Function


'// This is the call-back function for application.timer, it is used to capture the timer expiry
'// and call the expiry function on the singleton instance of zr_zr_appTimer
Public Sub zr_NextTick()
    zr_appTimer.appTimerHasExpired
End Sub

Public Function zr_ts2string(nStatus As ts_TimerStatus) As String

    Const kTimerStatusCodes = "RDY STD RUN PSD EXP"
    On Error Resume Next
    zr_ts2string = Mid(kTimerStatusCodes, nStatus * 4 + 1, 3)
    If Err.Number <> 0 Or Len(zr_ts2string) <> 3 Then zr_ts2string = "N/A"
    
End Function


Public Function zr_cs2string(nStatus As cs_CoreStatus) As String

    Const kCoreStatusCodes = "IDL WTG STG"
    On Error Resume Next
    zr_cs2string = Mid(kCoreStatusCodes, nStatus * 4 + 1, 3)
    If Err.Number <> 0 Or Len(zr_cs2string) <> 3 Then zr_cs2string = "N/A"
    
End Function


Public Function zr_ev2string(nEventType As ev_EventType) As String
'
'    Public Enum ev_EventType
'         ev_Start = 1
'         ev_tick = 2
'         ev_Expire = 3
'         ev_Cancel = 4
'         ev_Pause = 5
'         ev_Continue = 6
'    End Enum

    Const kEventTypeCodes = "00000 START TICK* EXPIR CANCL PAUSE RESUM "
    On Error Resume Next
    zr_ev2string = Mid(kEventTypeCodes, nEventType * 6 + 1, 5)
    If Err.Number <> 0 Or Len(zr_ev2string) <> 5 Then zr_ev2string = "*N/A*"
    
End Function


'// Logging ==========================================================================
'// Private logging functions
    Private Sub log(msg As String)
        '// Ouputs a message to the log, if logging is turned on, at the current indent level
        zr_Log.log "CoreM] " & msg
    End Sub
    Private Sub R(label As String)
        '// Ouputs a message to the log, if logging is turned on, and indents (margin to the right) with the label text
        zr_Log.R "CoreM] " & label
    End Sub
    Private Sub L()
        '// If logging is turned on, outdents (margin to the left) and writes the label from the preceding indent
        zr_Log.L
    End Sub
    Private Sub logErr(msg As String)
        zr_Log.log "* ERROR CoreM]" & msg
    End Sub


