VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zr_clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private myID As Long
Public label As String

Private xCurrentEventTickNumber As Long '// when event is raised, this is the tick number that triggered the event
                                        '// or it is the current tick number calculated from the current time

'// !!!!!!!!! TEMPORARILY SET TO PUBLIC
Public xNextEventTickNumber As Long    '// the tick number at which the next event should be actioned, based on config
Private xNextEventType As ev_EventType '// the type of event that will be raised on the next tick

Public iQueuedTickNumber As Long        '// the tick number at which this timer has already been queued
                                        '// set to MaxTickNumber if not queued

'// Start event info
Private xTimerHasBeenStarted As Boolean '// indicates that the user has called Start
Private xStartEventRaised As Boolean    '// indicates that a Start event has been raised

Private xIsRunning As Boolean      '// indicates that the Start event has been issued

Private xTimerHasBeenPaused As Boolean  '// indicates that the user has called Pause
Private xPauseEventRaised As Boolean    '// indicates that a Paused event has been raised after the user called Pause
Private xTimerHasBeenContinued As Boolean '// indicates that the user has called Resume
Private xContinueEventRaised As Boolean   '// indicates that the Continued event has been raised after the call to Resume


Private xIntervalTickerIsActive As Boolean  '// indicates that the interval ticker is active
Private xTickerInterval As Date             '// the interval in time that applies to this active interval ticker
Private xTickerTicks As Long                '// The number of ticks in the ticker interval

Private xExpirySet As Boolean           '// indicates that the timer has expiry time set
Private xNoExpirySet As Boolean         '// indicates that the timer has no expiry time set
Private xHasNotExpired As Boolean       '// indicates that the timer has not expired
Private xHasExpired As Boolean          '// indicates that the timer has expired
Private xExpiryEventRaised As Boolean   '// indicates that the expiry event has been raised

Private xLifeTime As Date               '// The time from start to expiry,
Private xLifeTicks As Long              '// ticks from Start to expiry, expiry depends on when start is called
Private xExpiresAtTime As Date          '// user-specified date/time for expiry, expiry is independent of when start is called

Private xExpiryTickNumber As Long       '// the tick number at which the expiry event will be / has been raised

Private xTimerHasBeenCancelled As Boolean      '// indicates that the timer has been removed by the user - no more events will be raised

Private xUserData As Object '// User Data to be sent with all events

'// The tick queue holding this event
Private xTickQueue As zr_clsTickQueue



'// user-assigned stream label. All timers with the label can be
'// monitored through the corresponding stream monitor object
Public streamLabel As String

Private xStreamMonitor As ZR_clsStream


'// Links for timer 2xlinked list
Public oNextTimer As zr_clsTimer
Public oPrevTimer As zr_clsTimer


'// EVENT DECLARATIONS
Public Event Started(eventData As zr_clsEventData)
Public Event Paused(eventData As zr_clsEventData)
Public Event Continued(eventData As zr_clsEventData)
Public Event Tick(eventData As zr_clsEventData)
Public Event Expired(eventData As zr_clsEventData)
Public Event Cancelled(eventData As zr_clsEventData)

'// Logging module label
Private xLogLabel As String
Public Property Get nextEventTickNumber() As Long
    nextEventTickNumber = xNextEventTickNumber
End Property

Public Property Get expiryTickNumber() As Long
    expiryTickNumber = xExpiryTickNumber
End Property

Public Property Get isRunning() As Boolean
    isRunning = xIsRunning
End Property


'//===================================================================================
Public Function fireEvent(Optional piTickNumber As Long = 0) As Long
Dim i As Integer: i = R("fireEvent #" & piTickNumber)

'// This is called from the core when a tick event occurs. The purpose is
'// to check if any events need to be fired on this timer and if so
'// fire them. The function returns the tick number of the next
'// event on this timer
    
    '// Check if the next event on this timer is in the future - no action required
    If piTickNumber < xNextEventTickNumber Then
        log "Tick number too low, next tick is " & xNextEventTickNumber
        L
        Exit Function
    End If
    
    '// The next event on this timer is due
    xCurrentEventTickNumber = piTickNumber
    
    '// Construct the event data to be passed in the event
    Dim evtData As zr_clsEventData
    
    Set evtData = getNextEventData
    
    With evtData
        
        log "Firing event ******"
        .report
        
    End With
    
    '// Now we are ready to fire the event
    
    '// Check if we have to send event to a stream
    If Len(streamLabel) > 0 Then
        xStreamMonitor.Fire evtData
    End If
    
    
    
'Enum zr_Core.evtEventType
'    evStart = 0
'    evTick = 1
'    evExpire = 2
'    evRemove = 3
'    evPause = 4
'    evResume = 5


    Select Case xNextEventType
    Case ev_EventType.ev_start
        RE_Started evtData
    Case ev_EventType.ev_tick
        RE_Tick evtData
    Case ev_EventType.ev_Expire
        RE_Expired evtData
    Case ev_EventType.ev_Cancel
        RE_Cancelled evtData
    Case ev_EventType.ev_Pause
        RE_Paused evtData
    Case ev_EventType.ev_Continue
        RE_Continued evtData
    Case Else
        xError "Unrecognised next event type"
        Exit Function
    End Select
    
    Set evtData = Nothing
    
    
    fireEvent = CalcNextEventTickNumber
    
    
L i
End Function

'// Loads a new event data object with the relevant data for the next event
Private Function getNextEventData() As zr_clsEventData
    Dim oEV As New zr_clsEventData
    With oEV
        .timerID = Me.ID
        .eventTickNumber = xNextEventTickNumber
        
        .eventTime = ZR_Core.zr_TimeFromTickNumber(xNextEventTickNumber)
        If xExpirySet Then
        
            .expiryTickNumber = xExpiryTickNumber
            .expiresAtTime = xExpiresAtTime
            .ticksToExpiry = xExpiryTickNumber - xNextEventTickNumber
            .timeToExpiry = xExpiresAtTime - .eventTime
            
        End If
        
        
        .eventType = xNextEventType
        .eventTypeName = ZR_Core.zr_ev2string(xNextEventType)
        .timerLabel = label
        .streamLabel = streamLabel
        
        'Set .timer = Me
        Set .userData = xUserData
        
    End With
    Set getNextEventData = oEV
End Function


Private Function CalcNextEventTickNumber() As Long
'// This function examines the current state of the timer and works out when
'// the next event will be due, and the type of that event, then it informs the
'// core so that if required, the core can adjust the core timer.

    

    '// Assume there is no event due, so next event is at the limit
    xNextEventTickNumber = 999999999
    
    '// If the user has not started the timer, then we don't need to do anything
    If Not xTimerHasBeenStarted Then Exit Function
    If xHasExpired And xExpiryEventRaised Then
        xTickQueue.RemoveTimer Me
        Exit Function
    End If
    
    
    '// set time to next event to non-value
    Dim iTicksToNextEvent As Long
    iTicksToNextEvent = -99
    
    '//
    '// ---------------- This first group are all immediate events, which are raised after user action
    '//
    
        '// After starting, the first event is always the zr_Start event and it will be fired at the next tick
        If Not xStartEventRaised Then
                '// Calculate the expiry tick number, if expiry details have been set
                If xLifeTicks > 0 Then
                    xExpiryTickNumber = ZR_Core.zr_TickNumberFromTime(Now) + xLifeTicks
                End If
                '// set the zr_Start tick to the next available tick number
                iTicksToNextEvent = 0
                xNextEventType = ev_EventType.ev_start
                
        '// Timer might have been paused, in which case check for the pause event having been raised
        ElseIf xTimerHasBeenPaused And Not xPauseEventRaised Then
                
                '// Need to raise the pause event to report that the timer has been paused
                iTicksToNextEvent = 0
                xNextEventType = ev_EventType.ev_Pause
                
        '// If the timer was paused and now has been Continued, then we need to raise the
        '// Continued event
        ElseIf xTimerHasBeenPaused And xTimerHasBeenContinued And Not xContinueEventRaised Then
        
                '// Raise the resume event to report that the timer has been Continued
                iTicksToNextEvent = 0
                xNextEventType = ev_EventType.ev_Continue
                       
        ElseIf xTimerHasBeenPaused Then
            '// Do not do anything, waiting for the resume action
        Else
    
    '// --------------------- The next group are the automatic events, such as tick and expiry
                '// If the timer is generating regular ticks, get the time to the next tick
                If xIntervalTickerIsActive Then
                    
                    iTicksToNextEvent = xTickerTicks
                    xNextEventType = ev_EventType.ev_tick
                End If
                
                '// If the timer has an expiry date, then work out how long to
                '// expiry, and check if that's earlier than the next ticker tick
                If xExpirySet And xHasNotExpired Then
                
                    Dim iTicksToExpiry As Long
                    If xExpiryTickNumber = MaxTickNumber Then
                        
                    End If
                    iTicksToExpiry = xExpiryTickNumber - xCurrentEventTickNumber
                    If iTicksToExpiry <= iTicksToNextEvent Or iTicksToNextEvent = -99 Then
                        iTicksToNextEvent = iTicksToExpiry
                        xNextEventType = ev_EventType.ev_Expire
                    End If
                    
                End If
        
        End If
    '// ===============================================================
    '// Now check the resulting next tick and set the timer if needed
    
    
    '// Check if the next tick has been set, and if so set the timer
    If iTicksToNextEvent > -99 Then
    
        '// We have established which event type is next, and when it should be fired
        xNextEventTickNumber = ZR_Core.zr_TickNumberFromTime(Now) + iTicksToNextEvent
        
        '// Make sure that this timer is added to the event queue
        If xNextEventTickNumber <> Me.iQueuedTickNumber Then xTickQueue.insertTimer Me
        
    End If
    
    
    '// set return value
    CalcNextEventTickNumber = xNextEventTickNumber
    
    '//xTickQueue.setQueueTimer
    '//>>>>>>>>>>>
    
End Function

'// USER FUNCTIONS ---------------------------------------------------------
'// These functions are called by the user to initiate actions on the timer

Public Function Start() As zr_clsTimer

'// Causes the timer to become active. The events will be [started], then [tick] if relevant, then [expired] if relevant
    Set Start = Me
    
    Select Case True
    Case xTimerHasBeenStarted
        xError "Attempt to zr_Start timer twice"
        Exit Function
        '/-------------->>>>>>
    Case xTimerHasBeenCancelled
        xError "Attempt to start timer that was removed"
        Exit Function
        '/-------------->>>>>>
    End Select
    
    xTimerHasBeenStarted = True
    xStartEventRaised = False
    
    If Len(streamLabel) <> 0 Then
        Set xStreamMonitor = zr_StreamMonitor(streamLabel)
    End If
    
    CalcNextEventTickNumber
    
    
    Me.report "After Start"

    xTickQueue.setQueueTimer
End Function


Public Sub Pause()
    xTimerHasBeenPaused = True
    xTimerHasBeenContinued = False
    
    Dim evd As zr_clsEventData
    Set evd = getNextEventData
    
    With evd
        .eventTickNumber = 0
        .eventType = ev_Pause
        .eventTypeName = zr_ev2string(ev_Pause)
        .report
    End With
    
    RaiseEvent Paused(evd)
    
    '// Calculate the next tick number and insert the timer into
    '// the queue at the corresponding tick position
    CalcNextEventTickNumber
    
    '// Nudge the tick queue to check if an adjustment to the app timer is needed
    xTickQueue.setQueueTimer
    

End Sub

'// Called by user to Continue timer after previously calling pause
Public Sub Continue()
    If Not xTimerHasBeenPaused Then Exit Sub
    If xTimerHasBeenCancelled Then Exit Sub
    
    Dim evd As zr_clsEventData
    Set evd = getNextEventData
    
    With evd
        .eventTickNumber = 0
        .eventType = ev_Continue
        .eventTypeName = zr_ev2string(ev_Continue)
        
        .report
    End With
    
    xTimerHasBeenPaused = False
    xTimerHasBeenContinued = True
    xContinueEventRaised = True
    
    RaiseEvent Continued(evd)
    
    '// We can reset these as all actions relating
    '// to pausing and resuming haave been completed
    xTimerHasBeenContinued = False
    xContinueEventRaised = False
    
    
    CalcNextEventTickNumber
    xTickQueue.setQueueTimer
    
End Sub

'// Cancel timer disconnects the timer from the timer
'// queue and effectively disables it. Once removed
'// the timer will not generate any new events, so
'// references to the timer should be discarded.
Public Sub Cancel()
    
    '// remove the timer from the queue
    xTickQueue.RemoveTimer Me
    
    '// set flags
    xTimerHasBeenCancelled = True
    xIsRunning = False
    
    '// Raise the Cancelled event
    Dim evd As zr_clsEventData
    Set evd = getNextEventData
    
        With evd
        
            .eventType = ev_EventType.ev_Cancel
            .eventTypeName = ZR_Core.zr_ev2string(ev_EventType.ev_Cancel)
            .eventTickNumber = ZR_Core.zr_TickNumberFromTime(Now)
            
        End With
    
    RaiseEvent Cancelled(evd)
    
    If Len(streamLabel) > 0 Then
        zr_StreamMonitor(streamLabel).Fire evd
    End If
    
    
End Sub


'// =======================================
'// Timer events raised from the core
'Enum zr_Core.evtEventType
'    evStart = 0
'    evTick = 1
'    evExpire = 2
'    evCancel = 3
'    evPause = 4
'    evResume = 5

Private Sub RE_Started(evd As zr_clsEventData)
log "+RE_Started"
    RaiseEvent Started(evd)
    xStartEventRaised = True
    '// IS ANYTHING MORE NEEDED HERE?
    CalcNextEventTickNumber
End Sub
Private Sub RE_Tick(evd As zr_clsEventData)
log "+RE_Tick"
    '// Tick event does not change status of the timer
    RaiseEvent Tick(evd)
    CalcNextEventTickNumber
End Sub
Private Sub RE_Expired(evd As zr_clsEventData)
log "*+RE_Expired"
    RaiseEvent Expired(evd)
    xHasExpired = True
    xHasNotExpired = False
    xExpiryEventRaised = True
    xExpiryTickNumber = xCurrentEventTickNumber
    
End Sub
Public Property Get isActive() As Boolean
    isActive = xTimerHasBeenStarted And Not xExpiryEventRaised And Not xTimerHasBeenPaused
    
End Property
Private Sub RE_Cancelled(evd As zr_clsEventData)
log ">>RE_Cancelled"

    xTickQueue.RemoveTimer Me
    xTimerHasBeenCancelled = True
    RaiseEvent Cancelled(evd)
    
End Sub
Private Sub RE_Paused(evd As zr_clsEventData)
log ">>RE_Paused"
    xTimerHasBeenPaused = True
    xPauseEventRaised = True
    
    xTimerHasBeenContinued = False
    xContinueEventRaised = False
    RaiseEvent Paused(evd)

End Sub
Private Sub RE_Continued(evd As zr_clsEventData)
log ">>RE_Continued"

    xTimerHasBeenPaused = True
    xTimerHasBeenContinued = True
    xPauseEventRaised = True
    
    RaiseEvent Continued(evd)
    
    xTimerHasBeenPaused = False
    xTimerHasBeenContinued = False
    xContinueEventRaised = False
    xPauseEventRaised = False
    
End Sub

Public Sub setConfig(Optional lifeTime As Date = 0, Optional expiryDateTime As Date = 0, Optional tickerInterval As Date = 0)
    Select Case True
    Case lifeTime > 0
       xLifeTime = lifeTime
        
    Case expiryDateTime > 0
        Me.expiresAt expiryDateTime
        
    Case tickerInterval <= 0
        Exit Sub
    End Select
    
    If tickerInterval > 0 Then Me.tickerInterval tickerInterval
    
End Sub
Private Sub resetExpiry()
    xExpirySet = False
    xNoExpirySet = True
    xExpiresAtTime = MaxTime
    xExpiryTickNumber = MaxTickNumber
    xLifeTime = 0
    xLifeTicks = 0
    xHasExpired = False
    xExpiryEventRaised = False
    xHasNotExpired = True
    
End Sub

Public Function expiresAt(pExpiryDateTime As Date) As zr_clsTimer
    xExpiresAtTime = pExpiryDateTime
    xExpiryTickNumber = zr_TickNumberFromTime(pExpiryDateTime)
    xExpirySet = True
    xNoExpirySet = False

    Set expiresAt = Me
    
End Function
Public Function lifeTime(pLifeTime As Date) As zr_clsTimer
    xLifeTime = pLifeTime
    xLifeTicks = zr_TicksFromInterval(pLifeTime)
    xExpiresAtTime = 0
    
    
    Set lifeTime = Me
    
End Function
Public Function tickerInterval(pTickerInterval As Date) As zr_clsTimer
    xTickerInterval = pTickerInterval
    xTickerTicks = ZR_Core.zr_TicksFromInterval(pTickerInterval)
    
    xIntervalTickerIsActive = True
    
    
    Set tickerInterval = Me
    
End Function


Public Property Set userData(ud As Object)
    xUserData = ud
End Property

Public Property Get userData() As Object
    userData = xUserData
    
End Property
Public Property Get ID() As Long
    ID = myID
End Property


'// ==========================================================================
'// Initializes all timer properties to default values
Private Sub Class_Initialize()
    myID = ZR_Core.zr_NextTimerID
    xLogLabel = "Timer #" & myID
    Debug.Print "Timer #" & myID & " ********* Initialize"
    
    Reset
    
End Sub


Public Function clone() As zr_clsTimer
    Dim oClone As zr_clsTimer
    Set oClone = ZR_Core.getNewTimer()
    oClone.setConfig xLifeTime, xExpiresAtTime, xTickerInterval, xStreamLabel
    oClone.label = increment(Me.label)
    oClone.streamLabel = Me.streamLabel
    
    Set clone = oClone
    
End Function


Private Sub Reset()
    'Set zr_Core = zr_Core.zr_Core
R "Reset"

    iQueuedTickNumber = 0
    
    Set xTickQueue = ZR_Core.zr_TickQueue
    xTickQueue.RemoveTimer Me
    
    
    xCurrentEventTickNumber = zr_Constants.NullTickNumber
    xNextEventTickNumber = zr_Constants.NullTickNumber
    xNextEventType = ev_EventType.ev_start
    
    xTimerHasBeenStarted = False
    xStartEventRaised = False
    
    xIsRunning = False
    
    xTimerHasBeenPaused = False
    xPauseEventRaised = False
    
    xTimerHasBeenContinued = False
    xContinueEventRaised = False
    
    xIntervalTickerIsActive = False
    xTickerInterval = 0
    
    label = "T#" & Me.ID
    streamLabel = ""
    
    xExpirySet = False
    xNoExpirySet = True
    xHasNotExpired = True
    xExpiryTickNumber = zr_Constants.NullTickNumber
L
End Sub

Public Sub report(Optional label As String = "")
'// Reports the timer information to immediate console
    
    log "+*** " & label & " ---------------------------ID=" & Me.ID
    log IIf(xIntervalTickerIsActive, "| Is ticking interval=" & xTickerInterval, "| Is not ticking")
    log "| Started " & xTimerHasBeenStarted & ", event raised " & xStartEventRaised
    
    If xHasNotExpired Then log "| Not expired, expiry at tick " & xExpiryTickNumber
    log "| Expiry tick number: " & xExpiryTickNumber
    log "| Next event " & ZR_Core.zr_ev2string(xNextEventType) & " at tick #" & xNextEventTickNumber
    
    
    log "+*** ------------------------------------------"
    
End Sub

'// Used for reporting only
Public Function sStatusFlags() As String
    Dim sinfo As String
    sinfo = YN(xTimerHasBeenStarted) & _
            YN(xStartEventRaised) & _
            YN(xTimerHasBeenPaused) & _
            YN(xPauseEventRaised) & _
            YN(xTimerHasBeenContinued) & _
            YN(xContinueEventRaised) & _
            YN(xIntervalTickerIsActive) & _
            YN(xIsRunning) & _
            YN(xHasExpired) & _
            YN(xExpiryEventRaised)
            
    sStatusFlags = sinfo
    
End Function
Private Function YN(b As Boolean)
    YN = IIf(b, "Y", "N")
End Function
Public Sub oneLineReport()
    log " NE#" & xNextEventTickNumber & " EVT=" & ZR_Core.zr_ev2string(xNextEventType) & " FLAGS=" & sStatusFlags
End Sub

Private Sub xError(msg As String)
    log "***ERROR*** " & msg
End Sub



'// ---------------------------------------------- Logging functions


    Private Sub log(msg As String)
        zr_Log.log xLogLabel & " " & msg
    End Sub
    Private Sub L(Optional pIndent As Integer = -1)
        zr_Log.L pIndent
    End Sub
    Private Function R(label As String) As Integer
        R = zr_Log.R(xLogLabel & " " & label)
    End Function


Private Sub Class_Terminate()
    Debug.Print "xxxxxxxxxxxxxx TERMINATE"
End Sub

