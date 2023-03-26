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

![](assets/20230326_170051_CodeExample1.png)

# Programming guide

## Core methods

These methods and properties are members of the zr_Core module. They are globally available


| Method      | Parameters                                                              | Notes                                                                                                                                                                   |
| :------------ | :------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| zr_start    | None                                                                    | Initialises the timer system. Is called implicitly by most other core methods.                                                                                          |
| zr_newTimer | timeToExpiry [o date], expiryDateTime [o date], tickerInterval [o date] | All arguments are optional. There are methods that correspond to these arguments which can be set separately.<br /> Returns a new Timer object with the given settings. |

## Timer methods

These methods and properties are members of the class zr_clsTimer. They are invoked on an instance of the class.


| Method              | Parameters         | Notes                                                                                                                                                                                                                                                                                                                                                     |
| :-------------------- | -------------------- | :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Start               | None               | Starts the timer. The expiry time is calculated on start. Tick events will be issued at intervals specified using**expiresAt** method.                                                                                                                                                                                                                    |
| Pause               | None               | Pauses a timer if it is running. During pause no Tick or Expired events are emitted. The Expiry date/time are not affected by pause. If expiry time is passed during pause, the Expired event will be emitted as soon as Continue method is called.                                                                                                       |
| Continue            | None               | Continues the timer after a pause. Tick and Expired events will be issued after Continue has been called                                                                                                                                                                                                                                                  |
| Cancel              | None               | Stops the timer and removes it from the timer system. The Cancelled event is emitted.                                                                                                                                                                                                                                                                     |
| lifeTime            | pLifeTime [m date] | The parameter sets the period of time that is to elapse between Start and Expiry. The period starts when Start method is called. If the timer is already running, it will run for the life time starting from now.                                                                                                                                        |
| expiresAt           | pExpiryDateTime    | The parameter sets the exact date and time when the timer will expire. This is fixed and will not depend on when Start was called.                                                                                                                                                                                                                        |
| tickerInterval      | pTickerInterval    | The parameter specifies the time period to elapse between ticks. Minimum time is 1 second.                                                                                                                                                                                                                                                                |
| userData [property] | ud [m object]      | Specifies any object that contains data to be passed in events emitted by this timer. The object is referenced dynamically so subsequent changes made to any properties of the object will be reflected on later events emitted by the timer. May be set to Nothing. May be set while the timer is running. Syntax: Set oTimer.userData = userDataObject. |
| ID                  |                    | Returns a long integer that is a unique identifier of this Timer instance.                                                                                                                                                                                                                                                                                |

w

wertwertwe
