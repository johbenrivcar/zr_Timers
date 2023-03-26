VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zr_clsTickEntry"
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

'// Implements the class of objects that form the tick queue under the control
'// of clsTickQueue. The queue is a forward-linked list of entries of this class
'// each of which represents a future tick number. There is one entry per required
'// tick number. A tick number is required if there is a timer whose next event
'// will fire on that tick.
'// When a new timer is started, the next event tick is calculated and the timer
'// is attached to the corresponding tick entry in the tick queue. When the time
'// comes to activate the timer event, this tick entry will be notified via a
'// call to FireTimers. After all the timers have fired this entry is discarded.

'// The current, i.e. last processed, tick number of the queue
Public tickNumber As Long

'// Dictionary of all timers that have been inserted into the queue, by timer ID
Public dTimers As New Dictionary

'// The anchor entry, which never changes. The first queue entry is reference by oNext.oNext
Public oNext As zr_clsTickEntry

'// ====================================================
'// Initialization
'//
Private Sub Class_Initialize()
    
    tickNumber = -1
    dTimers.RemoveAll
    Set oNext = Nothing
    
End Sub

Public Property Get count() As Integer
    count = dTimers.count
End Property
Public Sub AddTimer(oTimer As zr_clsTimer)
'// adds the given timer to this tick number

    dTimers.Add oTimer.ID, oTimer
    oTimer.iQueuedTickNumber = Me.tickNumber
    
log "Added T#" & oTimer.ID & " to Tick " & Me.tickNumber & ", count=" & Me.count


End Sub
Public Sub RemoveTimer(oTimer As zr_clsTimer)
    If dTimers.Exists(oTimer.ID) Then dTimers.Remove oTimer.ID
    oTimer.iQueuedTickNumber = 0
    
log "Removed T#" & oTimer.ID & " from Tick " & Me.tickNumber & ", count=" & Me.count

End Sub

Public Sub FireTimers(iTickNumber As Long)
Dim i As Integer: i = R("FireTimers " & iTickNumber & " in TickQueue #" & Me.tickNumber)

    If iTickNumber < Me.tickNumber Then
        log "!!! ERROR !!! - attempt to fire timers with out-of line tick number"
        L i
        Exit Sub
        '//----------------------> EXIT
    End If
    
        
        Dim oT As zr_clsTimer
        Dim keys As Variant
        Dim key As Variant
        keys = dTimers.keys
        
        For Each key In keys
            Set oT = dTimers(key)
            dTimers.Remove key
            
            oT.iQueuedTickNumber = zr_Constants.MaxTickNumber
            
            oT.fireEvent iTickNumber
            DoEvents
        Next key
    
    
L i
End Sub

    Public Sub report()
        '// lists all the timers attached to this tick entry. They all share the
        '// same tick number for their next event, and will be fired when the timer reaches it.
        
        On Error Resume Next
        Dim snext As String
        
        If oNext Is Nothing Then snext = " (END)" Else snext = " -> #" & oNext.tickNumber
        
        log snext
        If dTimers.count > 0 Then
            R ">----------- " & dTimers.count
            Dim key As Variant
            For Each key In dTimers.keys
                dTimers(key).oneLineReport
            Next key
            L
        End If
    End Sub


'// Private logging functions
    Private Sub log(msg As String)
        zr_Log.log "Tick #" & Me.tickNumber & " " & msg
    End Sub
    Private Sub L(Optional pIndent As Integer = -1)
        zr_Log.L (pIndent)
    End Sub
    Private Function R(Optional label As String = "") As Integer
        R = zr_Log.R("Tick#" & Me.tickNumber & " " & label)
    End Function

Private Sub Class_Terminate()
    Debug.Print "TickEntry #" & Me.tickNumber & " xxxxxxx Being terminated"
    
    Dim keys: keys = dTimers.keys
    Dim key
    For Each key In keys
        dTimers(key).Cancel
    Next key
    
End Sub
