VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zr_clsTickQueue"
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
'// Implements the manager for the central queue of timer events waiting for
'// the application.timer to trigger event nofications to the user application.
'//
'// The queue contains a chain of clsTickEntry objects, one for each tick on which
'// at least one timer event is waiting.
'//
'// The clsAppTimer is used to control setting and expiry of a timer for the first entry
'// (i.e. with the lowest tick number) in the queue.
'//
'// If any new events are added to the queue which have a tick number earlier than the
'// current earliest entry, then the new events are added at the head of the queue, and
'// the application.timer is replaced with one that will fire at the new, earlier time.
'// There is only ever a single application.timer waiting, corresponding always to the
'// first entry in the queue.
'//
'// If the queue is empty then no application.timer is set.
'//
'// The schematic below represents the queue with the first two entries for tick #324 and #492.
'// In this case the appTimer will be set to fire on tick #324, the lowest tick number in the
'// queue.
'  ______________             ______________          ______________                ______________
' | Head         | Next      | Entry        | Next   | Entry        | Next         | Tail         |
' |  tick #-1    |---------->|  tick #324   |------->|  tick #492   |---/..../---->|  tick #99999 |
' |______________|           |______________|        |______________|              |______________|
'                                    |                       |
'                                    |                       |
'                             _______V______           ______V_______
'                            | Timers       |_        | Timers       |_
'                            |  tick #324   | |       |  tick #492   | |
'                            |______________| |       |______________| |
'                              |______________|         |______________|
'
'


    '// Static head element. The empty queue has head and tail elements, any
    '// new entries are inserted between these, in order of
    '// tick number. The first element in the queue (lowest tick number) is
    '// found at Head.oNext
    Public Head As New zr_clsTickEntry
    
    '// Dictionary of ticks in the queue, by tick number. Used when adding a new
    '// timer to find if a tick entry already exists for the required tick number
    Public dTickEntriesInQueue As New Dictionary
    
    '// Reference to the singleton AppTimer which wraps the actual system.timer
    '// that will be set to fire at the first tick number
    Private WithEvents appTimer As ZR_clsAppTimer
Attribute appTimer.VB_VarHelpID = -1
    

'// ===================================================================
'// Initialization. Sets all defaults
Private Sub Class_Initialize()

    '// Get system object from core
    Set appTimer = ZR_Core.zr_appTimer
    'Set sysEvents = zr_Core.zr_SysEvents
    
    '// Create a new entry to act as tail of the queue
    Dim oTail As New zr_clsTickEntry
    
    '// Tick number of the head
    Head.tickNumber = -1
    
    '// Tick number of the tail
    oTail.tickNumber = zr_Constants.MaxTickNumber
    
    '// Add the tail entry to the head reference to form an
    '// empty queue.
    Set Head.oNext = oTail
    
End Sub

'// ----------------------------------------------
'// Returns the first tick number in the queue. This is
'// always the tick number of the first element, pointed to
'// by Head.oNext
Public Property Get firstQueuedTickNumber() As Long
    firstQueuedTickNumber = Head.oNext.tickNumber
End Property


'// This is run whenever any new timer has been added to the queue.
'// The purpose is to tell the appTimer object when the next tick
'// needs to be fired.
Public Sub setQueueTimer()

    '// get the first tick number, which is the tick number of the
    '// first tick entry in the queue. This should be by definition
    '// the tick that will be fired next by the application.timer
    
    appTimer.setAppTimer_forTick firstQueuedTickNumber
    
End Sub

'// Event handler to pick up notification from the appTimer that
'// the previous timer has expired, all actions relating to that
'// timer have been completed. So, now there is no active timer,
'// and we have to check if a new timer is needed.
Private Sub appTimer_Finished(iTickNumber As Long)
    setQueueTimer
End Sub

'// Event handler to pick up notification from appTimer that
'// the application.timer has expired, so any timer events
'// due on or before this tick number must be fired.
Private Sub appTimer_Expired(iTickNumber As Long)

Dim indent As Integer: indent = R("")


    '// Get the tick number that corresponds to the real time, now.
    '// We ignore the tick number provided in the parameters, because
    '// we just need to know which events must be fired relative to
    '// real time.
    Dim iNowTickNumber As Long
    iNowTickNumber = ZR_Core.zr_TickNumberFromTime(Now)
    
    log "=============================================================="
    log "Loop start >#" & iNowTickNumber
    
    '// We have to process all the timers that have a tick number less
    '// than or equal to the Now tick
    
    '// Get the first entry in the queue
    Dim oEntry As zr_clsTickEntry
    Set oEntry = Head.oNext
    
    '// Loop to process entries from the tick queue
    '// that have a tick number later than now.
    Do Until oEntry.tickNumber > iNowTickNumber
        
        '// Remove the entry from the queue - any timers in this entry will
        '// need to be fired, and the entry dropped.
        log "Removing entry #" & oEntry.tickNumber & " from tick queue"
        Me.removeTick oEntry
        
        '// Call FireTimers on the entry to raise all the timer events
        log "Firing timers on #" & oEntry.tickNumber
        oEntry.FireTimers iNowTickNumber
        
        '// Get the next entry from the queue
        Set oEntry = Head.oNext
        log "Next entry is #" & oEntry.tickNumber
        
    Loop
    
log "------Loop end <#" & iNowTickNumber
L indent

End Sub



'// =================================================================
'// Inserts the given timer into the tick queue
'// The timer is added to the tick entry that corresponds to the timer's next event tick number.
'// If the timer was already in the queue, it is removed from the tick entry that it was previously attached to.
'// If there is no corresponding tick entry a new one is created and inserted into the tick queue.
Public Sub insertTimer(oTimer As zr_clsTimer)
Dim i As Integer: i = R("insertTimer id=" & oTimer.ID)
    
    '// Get the tick number of the next event for this timer
    Dim iTimerNextTick As Long: iTimerNextTick = oTimer.nextEventTickNumber
    Dim itick As Long
    Dim oTickEntry As zr_clsTickEntry
    
    '// Check if this timer is already registered under a tick number
    '// and, if so, is it the same tick number?
    itick = oTimer.iQueuedTickNumber
    If itick > 0 Then
    
        '// If this timer is already registered under the correct tick, do nothing
        If itick = iTimerNextTick Then
            log "Timer already queued at tick #" & itick
            L i
            Exit Sub
        End If
        
        '// Remove from that tick entry in the tick queue
        On Error Resume Next
        If dTickEntriesInQueue.Exists(itick) Then
            Set oTickEntry = dTickEntriesInQueue.Item(itick)
            log "Removing timer from entry #" & itick
            oTickEntry.RemoveTimer oTimer
        End If
        On Error GoTo 0
        
        
    End If
    
    
    
    '// Now get the entry for the next tick number for this timer
    '// If the entry did not previously not exist, then it is created by this function
    Set oTickEntry = getTickEntry(iTimerNextTick)
    
    '// Insert the timer into that entry
    log "adding timer to entry #" & iTimerNextTick
    
    oTickEntry.AddTimer oTimer

    
L i
End Sub

'// Removes a timer from the queue. The tick number is stored in the timer
'// under iQueuedTickNumber.
Public Sub RemoveTimer(oTimer As zr_clsTimer)

    Dim itick As Long
    itick = oTimer.iQueuedTickNumber
    
    '// Get hold of the entry corresponding to the tick number
    If dTickEntriesInQueue.Exists(itick) Then
        Dim oTick As zr_clsTickEntry
        Set oTick = dTickEntriesInQueue(itick)
        
        '// Remove the timer from the tick entry
        oTick.RemoveTimer oTimer
    End If
    
End Sub

'// Gets the tick entry from the queue corresponding to the given tick number
'// If no entry exists, then a new entry is created and inserted into the queue
Private Function getTickEntry(piTick As Long) As zr_clsTickEntry
Dim i As Integer: i = R("getTickEntry" & piTick)

    '// Does this tick number already exist?
    If dTickEntriesInQueue.Exists(piTick) Then
        Set getTickEntry = dTickEntriesInQueue(piTick)
        log "Tick entry already in queue, returning"
        L i
        Exit Function
        '||=================================> EXIT -- tick entry already exists >>
    End If
    
    
    '// ========================================= creating new tick entry
    '// We need a new entry in the tick queue
    log "creating new tick entry for #" & piTick
    
    Dim oNewTick As New zr_clsTickEntry
    oNewTick.tickNumber = piTick
    
    
    '// Add the tick to the tick dictionary
    dTickEntriesInQueue.Add piTick, oNewTick
    
    
    '// ========================================= Insert into the queue
    Dim oPrev As zr_clsTickEntry: Set oPrev = Head
    
    Do
        '// Trap for the end of the queue (this should not happen
        '// because the end of the queue is always created at maxTickNumber
        If oPrev.oNext Is Nothing Then Exit Do
        
        '// looking for the first entry with a ticknumber > than the new one
        If oPrev.oNext.tickNumber > piTick Then Exit Do
        
        '// Move to the next entry
        Set oPrev = oPrev.oNext
        
    Loop
    
    '// Now we have to insert the new tick entry into the queue between prev and next
    Set oNewTick.oNext = oPrev.oNext
    Set oPrev.oNext = oNewTick
    log "Inserted tick #" & piTick & " between #" & oPrev.tickNumber & " and #" & oNewTick.oNext.tickNumber
    
    Set getTickEntry = oNewTick
L i
End Function

'// Removes a tick entry from the queue
Public Sub removeTick(oTickEntry As zr_clsTickEntry)
Dim i As Integer: i = R("removeTick #" & oTickEntry.tickNumber)

    '// You can only remove the first entry in the queue
    If oTickEntry.tickNumber <> Head.oNext.tickNumber Then
        log "*** ERROR *** Attempt to remove tick entry that is not first in the queue"
        L i
        Exit Sub
    End If
    
    '// Change the next pointer at the head of the queue
    Set Head.oNext = Head.oNext.oNext
    log "Head of the queue is now #" & Head.oNext.tickNumber
    
    '// remove from the dictionary of queued ticks
    If dTickEntriesInQueue.Exists(oTickEntry.tickNumber) Then
        log "Removing #" & oTickEntry.tickNumber & " from queue dictionary"
        dTickEntriesInQueue.Remove oTickEntry.tickNumber
    End If
    
L i
End Sub

'// REPORTING FUNCTION - reports all entries in the queue
Public Sub report()
    Dim oEntry As zr_clsTickEntry
    Set oEntry = Head
    log "-------------------------- Ticker queue"
    Dim i As Integer: i = R("")
    
    Do
        oEntry.report
        Set oEntry = oEntry.oNext
        If oEntry Is Nothing Then Exit Do
        
    Loop
    L i
    log "========================== End of ticker queue"
End Sub

'// Logging ==========================================================================
'// Private logging functions
    Private Sub log(msg As String)
        '// Ouputs a message to the log, if logging is turned on, at the current indent level
        zr_Log.log "TickQueue] " & msg
    End Sub
    Private Function R(label As String) As Integer
        '// Ouputs a message to the log, if logging is turned on, and indents (margin to the right) with the label text
        R = zr_Log.R("TickQueue] " & label)
    End Function
    Private Sub L(Optional pIndent As Integer = -1)
        '// If logging is turned on, outdents (margin to the left) and writes the label from the preceding indent
        zr_Log.L pIndent
    End Sub
    Private Sub logErr(msg As String)
        zr_Log.log "* ERROR TickQueue]" & msg
    End Sub
