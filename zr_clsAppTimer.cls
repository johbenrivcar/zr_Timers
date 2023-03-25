VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZR_clsAppTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'// ========================================================================
'// Implements an intelligent wrapper around the application.timer method, to ensure
'// that there is only ever a single application.timer that is active.
'// The application.timer is always set to expire at the earliest required time
'// to fire the next timer event across all zr_clsTimer objects created by the
'// application.

'// The timer event queue manager (zr_clsTickQueue) always requests a timer to be
'// set for the time of the next waiting event in the queue. The AppTimer checks
'// to see if a timer has already been set and clears that timer if it is for a
'// different time than has been requested and cancels the old timer if so.

'// A new timer set to expire at the requested time is created if required.

'// The application timer always invokes the public function zr_Core.zr_NextTick
'// which calls back to the method .appTimerHasExpired on this AppTimer object.
'// This in turn raises the Expired event to propagate the expiry notification
'// back to the event queue manager.


Private xTickNumber As Long
Private xTimerTime As Date
Private bTimerIsSet As Boolean

Private xbTimerExpiryNotificationInProcess As Boolean


Public Event Expired(iTickNumber As Long)
Public Event Finished(iTickNumber As Long)
Public Event Cancelled(iTickNumber As Long)
Public Event Started(iTickNumber As Long)


'// Initialization ===================================================================
Private Sub Class_Initialize()
   
    R "Initialize"
    
    If zr_bAppTimerCreated Then
        log "***ERROR*** zr_appTimer is a singleton, cannot be instantiated more than once"
        
        Err.Raise 8001#, , "Attempt to create more than one application timer instance"
        
    End If
    zr_bAppTimerCreated = True

       
    L
    
End Sub

'// ================================================================================
'// Method called to set the timer to a point in time at the given time interval
'// from now.
Public Function setAppTimer_forInterval(dInterval As Date) As Long

    '// Check if we are still in the middle of processing an earlier expiry. Ignore
    '// the request if so. We can safely do this because the request for a new
    '// timer will be repeated later, if it is still required.
    If xbTimerExpiryNotificationInProcess Then Exit Function
    
    '// Now work out the tick number that corresponds to the required expiry
    Dim iTickNumber As Long
    iTickNumber = ZR_Core.zr_TickNumberFromTime(Now + dInterval)
    
    '// Set the timer for the requested tick number
    Me.setAppTimer_forTick iTickNumber
    
    '// Return the tick number to the caller (for info)
    setAppTimer_forInterval = iTickNumber
    
End Function

'// ==================================================================================
'// Sets the application.timer to the time corresponding to a specific tick number
'// The optional force parameter will always create a new timer even if one already
'// exists for the same requested time (normally we would leave the old timer running)
Public Function setAppTimer_forTick(piTickNumber As Long, Optional bForce As Boolean = False) As Long

    '// Do not do anything if we are in the middle of processing an early timer expiry
    If xbTimerExpiryNotificationInProcess Then Exit Function
    
    '// Check if this is a force for the same tick number as was last requested
    If bForce And piTickNumber = 0 Then piTickNumber = xTickNumber
     
    '// Get the tick number of the current real time
    Dim iNowTick As Long: iNowTick = ZR_Core.zr_TickNumberFromTime(Now)
        
    '// Check if there is an application.timer already running
    If bTimerIsSet Then
    
        '// Exit if this matches the tick number that has been requested (unless force has been set)
        If piTickNumber = xTickNumber Then If Not bForce Then Exit Function
        
        '// We need to remove the old timer that has been set
        On Error Resume Next
        '// Avoid trapping a system error if the timer is not found - we don't care if that's the case
        Application.OnTime xTimerTime, "zr_NextTick", , False
        On Error GoTo 0
        bTimerIsSet = False
        
        log "##TICK #" & xTickNumber
        
        '// Notify any listeners that application.timer was cancelled
        '// (not normally significant, used only for testing)
        RaiseEvent Cancelled(xTickNumber)
        
    End If
    
    '// If the tick number is set to max, this means there are no timers waiting
    '// and we do not need to set any new application.timer
    If piTickNumber >= zr_Constants.MaxTickNumber Then
        bTimerIsSet = False
        xTickNumber = zr_Constants.MaxTickNumber
        xTimerTime = 0
        Exit Function
    End If
    
    '// So now we actually set the application.timer
    xTimerTime = ZR_Core.zr_TimeFromTickNumber(piTickNumber)
    xTickNumber = piTickNumber
    Application.OnTime xTimerTime, "zr_NextTick", , True
    
    '// Flag that the timer has been set
    bTimerIsSet = True
    
    If (zr_Log.isLogging) Then log "**TICK #" & xTickNumber & " (waiting " & (xTimerTime - Now) * CDbl(SecondsPerDay) & " seconds)"
    
    '// Notify listeners that application.timer has been started
    '// (not normally significant, used only for testing)
    RaiseEvent Started(xTickNumber)
    
End Function

'// ==========================================================
'// This is called from the core module when the application timer expires
Public Sub appTimerHasExpired()

    '// Set state flags to reflect timer status
    bTimerIsSet = False
    
    '// Set the inprocess flag
    xbTimerExpiryNotificationInProcess = True
    
    '// Raise the timerExpired event to notify the tick queue handler
    '// or other process, of timer expiry
    log ">>TICK #" & xTickNumber
    
    Dim saveTN As Long: saveTN = xTickNumber
    
    RaiseEvent Expired(xTickNumber)
    
    '// Reset inprocess flag
    xbTimerExpiryNotificationInProcess = False
    
    '// Notify listeners that application.timer full cycle has been completed
    '// (not normally significant, used only for testing)
    log "<<EXPIRED #" & saveTN
    RaiseEvent Finished(saveTN)
    
End Sub



'// Logging ==========================================================================
'// Private logging functions
    Private Sub log(msg As String)
        '// Ouputs a message to the log, if logging is turned on, at the current indent level
        zr_Log.log "appTimer] " & msg
    End Sub
    Private Sub R(label As String)
        '// Ouputs a message to the log, if logging is turned on, and indents (margin to the right) with the label text
        zr_Log.R "appTimer] " & label
    End Sub
    Private Sub L()
        '// If logging is turned on, outdents (margin to the left) and writes the label from the preceding indent
        zr_Log.L
    End Sub
    Private Sub logErr(msg As String)
        zr_Log.log "** ERR appTimer] " & msg
    End Sub

