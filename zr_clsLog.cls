VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zr_clsLog"
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
'// Implements a singleton log object that is used to pipe all log messages
'// to a client process via the log(msg) event. The client process and write
'// log messages to any suitable output stream (e.g. listbox or text file)
'//
'// Can send log messages to debug (immediate) window if required
'//
Public Event log(msg As String)

Private xLogToDebug As Boolean
Private xCurrentIndent As String
Private xLRIndent As String

Private xIndentLevel As Integer
Private xLabelstack(2000) As String
Private xIsLogging As Boolean

'// ----------------------------------------
Public Property Get isLogging() As Boolean
    isLogging = xIsLogging
    
End Property


Private Sub xLog(Optional msg As String = "")
    If Not xIsLogging Then Exit Sub
    Dim ss As String: ss = xCurrentIndent & msg
    
    '// Send to debug if requested
    If xLogToDebug Then Debug.Print ss
    
    '// Inform any listeners of the logged message
    RaiseEvent log(ss)
    
End Sub

'// Public function used to write a text message to the log
Public Sub log(Optional msg As String = "")
    If Not xIsLogging Then Exit Sub
    
    xLog "|" & msg
    
End Sub

'// Called to suspend logging at any time
Public Sub stopLog()
    If xIsLogging Then log "---------------- Stopped logging at " & Now & " -------------------"
    xIsLogging = False
    
End Sub

'// Calle to start/Continue logging at any time
Public Sub startLog()
    If xIsLogging Then Exit Sub
    
    xIsLogging = True
    
    log "----------Started logging at " & Now & "-------------------------"
    log ""
    
End Sub

'// =====================================================================
'// R and L functions provide tracked indentation of log messages so that
'// the log stream can track call stack indentation. Note Rs and Ls should
'// always be invoked in pairs. e.g. on entry to and exit from a function
'// The indentation level is returned from the R call so that it can be
'// passed to the L call. The indentation is then checked to make sure
'// there are no missing L calls.
Public Function R(Optional label As String = "") As Integer
    '// Return the indentation level (starting at 0)
    R = xIndentLevel
    
    '// Store the label for this level (e.g. like the function name)
    xLabelstack(xIndentLevel) = label
    
    '// Write the indentation start message to the log if required
    If xIsLogging Then
        xLog "\\>>" & label
    End If
    
    '// Increment the indentation by one level
    xCurrentIndent = ". " & xCurrentIndent
    xIndentLevel = xIndentLevel + 1
    

End Function

'// L function handles the unindent, should be matched with an earlier call to R.
'// The indent parameter should be the indent level provided from the R call.
Public Sub L(Optional pIndent As Integer = -1)
    
    '// Decrement the current indent level (brings it down to the matching R level)
    If xIndentLevel > 0 Then xIndentLevel = xIndentLevel - 1
    
    '// If indent was supplied, check that it matches the expected level
    If pIndent >= 0 Then
        If pIndent <> xIndentLevel Then
            If xIsLogging Then log "**Error** Indent level mismatch - expected=" & xIndentLevel & " got=" & pIndent & " (" & xLabelstack(pIndent) & ")"
            
            '// Force indentation to the expected level
            xIndentLevel = pIndent
        End If
    End If
    
    '// reduce the indentation
    If xIndentLevel = 0 Then xCurrentIndent = "" Else xCurrentIndent = Right(xCurrentIndent, xIndentLevel * 2)
    
    '// Write the unindent message to the log file if required
    If xIsLogging Then xLog "//<<" & xLabelstack(xIndentLevel)
     
End Sub


'// These members control whether the system log is sent to the debug (immediate) window
Public Sub sendToDebugStart()
    xLogToDebug = True
End Sub

Public Sub sendToDebugStop()
    xLogToDebug = False
End Sub

'// Initialisation
Private Sub Class_Initialize()
    xIsLogging = False
    xCurrentIndent = "|"
    xIndentLevel = 0
    xLogToDebug = False
    
End Sub
