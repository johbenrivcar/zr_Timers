Attribute VB_Name = "zr_Constants"
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

'// Timer constants ==========================================

    Public Const MaxTickNumber = 999999999#
    Public Const NullTickNumber = -9999#
    
    Public Const SecondsPerDay = 86400#
    Public Const OneSecond = 1# / SecondsPerDay
    Public Const Asap = OneSecond / 10#
    Public Const FiveSeconds = 5# / SecondsPerDay
    Public Const TenSeconds = 10# / SecondsPerDay
    Public Const OneMinute = 60# / SecondsPerDay
    Public Const FiveMinutes = 300# / SecondsPerDay
    Public Const OneHour = 1# / 24#
    
    
    Public Const TicksPerSecond = 1 ' Always set to 1 - testing on other values has not been done
    
'// Public enumerators for status and type codes ========================
'// Status code for timers
    Public Enum ts_TimerStatus
         ts_Ready = 0
         ts_Started = 1
         ts_Running = 2
         ts_Paused = 3
         ts_Expired = 4
    End Enum

    Public Enum cs_CoreStatus
      cs_Unstarted = 0
      cs_IdleNoTimers = 10
      cs_WaitingForTick = 20
      cs_SettingNextTick = 30
    End Enum

    Public Enum ev_EventType
         ev_start = 1
         ev_tick = 2
         ev_Expire = 3
         ev_Cancel = 4
         ev_Pause = 5
         ev_Continue = 6
    End Enum
