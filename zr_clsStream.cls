VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZR_clsStream"
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

'// ========================================================================
'// Implements an event Stream object that merges events from multiple timers
'// The user application can assign a StreamLabel to any timer, and this will
'// cause that timer to publish events through the matching Stream object.
'//
'// The user application can obtain a reference to the stream object and
'// then monitor any/all of the events that are published by all timers
'// that have the same label.
'//

Public streamLabel As String


'// EVENT DECLARATIONS - events per event type
Public Event Started(evtData As zr_clsEventData)
Public Event Paused(evtData As zr_clsEventData)
Public Event Continued(evtData As zr_clsEventData)
Public Event Tick(evtData As zr_clsEventData)
Public Event Expired(evtData As zr_clsEventData)
Public Event Cancelled(evtData As zr_clsEventData)

'// Broadcast for all event types from all timers
'// Process by checking evtData for event information
Public Event Fire(evtData As zr_clsEventData)

'// =====================================================
'// This function is called by all timers that have a
'// matching stream label, whenever they publish an
'// event.
'// Note this function may also be called by the user
'// application to interpose their own events.
Public Sub Fire(eventData As zr_clsEventData)


'Standard timer event types:
'Public Enum ev_EventType
'    ev_Start = 1
'    ev_Tick = 2
'    ev_Expire = 3
'    ev_Cancel = 4
'    ev_Pause = 5
'    ev_Continue = 6
'End Enum

    '// Always raise the fire event for all incoming events
    RaiseEvent Fire(eventData)
    
    '// Raise other events depending on the type of the incoming event
    Select Case eventData.eventType
        Case ev_EventType.ev_start
            RaiseEvent Started(eventData)
        Case ev_EventType.ev_tick
            RaiseEvent Tick(eventData)
        Case ev_EventType.ev_Expire
            RaiseEvent Expired(eventData)
        Case ev_EventType.ev_Cancel
            RaiseEvent Cancelled(eventData)
        Case ev_EventType.ev_Pause
            RaiseEvent Paused(eventData)
        Case ev_EventType.ev_Continue
            RaiseEvent Continued(eventData)
        Case Else
            '// Do nothing, this corresponds (possibly) to a user-defined event type
    End Select
    
End Sub
'

Private Sub logErr(msg As String)
        ZR_Core.zr_Log.log "* ERROR StreamMonitor]" & msg
End Sub
