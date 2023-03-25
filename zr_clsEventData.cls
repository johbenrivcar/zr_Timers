VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zr_clsEventData"
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



'// Carries data relating to a specific timer event, which is passed as the single
'// parameter for all timer events.

'// Note that for some event types some data may not be filled in.

    Public timerID As Long
    Public userData As Object
    Public eventTickNumber As Long
    Public eventType As Long
    Public timerLabel As String
    Public streamLabel As String
    Public eventTypeName As String
    Public eventTime As Date
    Public expiresAtTime As Date
    Public expiryTickNumber As Long
    Public ticksToExpiry As Long
    Public timeToExpiry As Long
    
'// Report function (for debugging only)
Public Sub report()
    With zr_Log
        If Not .isLogging Then Exit Sub
        
        .log "--- Event from [" & IIf(Len(timerLabel) = 0, "Timer", timerLabel) & "]"
        .log "* tick #" & eventTickNumber
        .log "* type: " & eventTypeName & " (" & eventType & ")"
        .log "* stream: " & streamLabel
        .log "* ticks to expy: " & ticksToExpiry
        .log "* has user data? " & Not (userData Is Nothing)
    End With
End Sub
