# ZR Timers

This project provides and implementation of timer objects for use in MS-Excel macro-enabled spreadsheets.

All implementation is in vba, all source code is included

## Overview

* ZR source modules (.bas and .cls) must be included in the application Excel spreadsheet, or in a separate Excel spreadsheet whose project is referenced from the application sheet (a suitable spreadsheet is included in the project with latest versions of the source modules)
* Timers are vba objects which are configured at run time to fire events at intervals defined by the user, or as a result of user actions. Events are processed by application event handlers which attach to the timer through **WithEvents** syntax.

#### Summary of features

* Timers are created as vba objects. Events are handled by using **WithEvents** syntax.
* Timers may fire **Tick events** at a regular inteval >= 1 second.
* Timer **Expiry** is set as:
  * a period of time from .Start call or
  * a fixed point in time, regardless of when .Start is called.
* Timers may be Paused and Continued. During Pause, Tick events are suppressed. Pausing does not affect expiry.
* Timers emit events on change of state. Possible events are: **Started**, **Tick**, **Expired**, **Paused**, **Continued**, **Cancelled**.

## How to install

**Method 1 - Incorporate source modules into your Excel application**

This has the advantage that all code is included in your application, so it does not require reference to the zr_Timers project

1. Download all source modules (.cls [7 files] and .bas [2 files]) from GitHub to a local folder
2. Open your Excel application spreadsheet. Open the vba editor. In the project explorer, import all nine source modules into the application.
3. Use the zr_timer methods in your application to create and run timers.

**Method 2 - Refer to the zr_Timers project**

1. Download the Excel spreadsheet zr_Timers.xlsm from GitHub to a local folder
2. Open zr_Timers.xlsm and enable macros
3. Open your application Excel spreadsheet. Open the vba code window, click on your application in the project navigation pane, then
   * Choose Tools>References.
   * In the list of Available References find zr_Timers and set the check box
4. Now use the zr_Timers methods in your application. You may need to qualify references by prefixing calls with **zr_Timers**.

##### Example code

---

Option Explicit

Dim WithEvents tt As zr_clsTimer

Public Sub t1()
Set tt = zr_newTimer(tenSeconds)
tt.Start
End Sub

Private Sub tt_Expired(eventData As zr_Timers.zr_clsEventData)
Debug.Print "Timer has Expired"
End Sub

---

# Programming guide

## Core methods

These methods and properties are members of the zr_Core module


| Method      | Parameters                                                              | Notes                                                                                                                                                                   |
| :------------ | :------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| zr_start    | None                                                                    | Initialises the timer system. Is called implicitly by most other core methods.                                                                                          |
| zr_newTimer | timeToExpiry [o date], expiryDateTime [o date], tickerInterval [o date] | All arguments are optional. There are methods that correspond to these arguments which can be set separately.<br /> Returns a new Timer object with the given settings. |
