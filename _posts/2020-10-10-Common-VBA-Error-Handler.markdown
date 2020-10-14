---
layout: post
title: A common VBA Error Handler
subtitle: An Error Handler assembled from the best which can be found in foruns
date: 2020-10-02 16:00 +0200
categories: vba common
---


In this post<br>
[Function](#function)<br>
[Syntax](#syntax)<br>
[Installation of the Error Handler](#installation-of-the-error-handler)<br>
[Installation of the Alternative VBA MsgBox](#installation-of-the-alternative-vba-msgbox)<br>
[Basic usage](#basic-usage)<br>
[Usage which provides a "path to the error"](#usage-which-provides-a-path-to-the-error)<br>
[Usage which supports debugging](#usage-which-supports-debugging)<br>
[Usage which supports regression testing](#usage-which-supports-regression-testing)<br>
[Development, test, maintenance](#development-test-maintenance)


### Function
Only 4 additional code lines in a procedure make the difference (see [Basic usage](#basic-usage)).<br>
The _ErrHndlr_ functions appearance and behaviour is pretty similar to the VBA MsgBox as it by default displays an OK button only for example. Of course it  provides a lot more than just displaying a message and returning a clicked button's value. 

#### 1. Path to the error
A major advantage of the _ErrHndlr_ function: When there is no choice for the user, i.e. only one - usually the default OK - button is displayed, and [the _Entry Procedure_](#the-entry-procedure) is known the error is passed on back up to the _Entry Procedure_ by which the path to the error is assembled and finally displayed.

#### 2. Error type discrimination
The error message discriminates between _VB Runtime Error_, _Application Error_, and _Database-Error_

#### 3. Clear indication of the Error source
The source of the error is displayed in the form <module>.<procedure> (see [_Entry Procedure_](#entry-procedure))

#### 4. Display of an [Optional execution time trace](#optional-execution-time-trace)

Whenever an [_Entry Procedure_](#entry-procedure) is reached during execution, optionally an execution time trace is displayed in the VBE immediate window

#### 5. Free buttons specification
When the _Alternative VBA MsgBox_ (UserForm _fMsg_) is used the error message may be displayed with (nearly) any number of _buttons_ with a desired caption string, even in combination with the _VBA MsgBox_ buttons value (vbYesNo, etc.). This offers a very elegant way for a  [usage which supports debugging](#a-usage-which-supports-debugging).

#### 6. Optional error log
Yet not implemented

### Syntax
```vbs
ErrHndlr error-number, error-source, error-description, error-line[, buttons]
```
The procedure has these named arguments:

|  Argument  | Description |
| ---------- | ----------- |
| errnumber  |             |
| errsource  |             |
| errdscrptn |             |
| errline    |             |
| buttons    | Optional. Variant. Defaults to vbOkOnly when omitted.<br>May be  [value for the VBA MsgBox buttons argument](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) or - in case the Alternative VBA MsgBox (_fMsg_) is used - descriptive button caption strings, including line breaks, delimited by a comma. |

### Installation of the Error Handler
- Download and import [_mErrHndlr_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mErrHndlr.bas)
- Download and import [_clsCallStack_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/clsCallStack.cls)
- Download and import [_clsCallStackItem_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/clsCallStackItem.cls)

### Installation of the Alternative VBA MsgBox
See the [Debugging](#debugging) for one of the benefits of it.
- Download [fMsg.frm](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm)
- Download  [fMsg.frx](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsf.frx)
- Import _fMsg.frm_ 
- in the module _mErrHndlr_ set the local Conditional Compile Argument:<br>`#Const AlternateMsgBox = 1`

### Basic usage
 ```vbscript
 Public/Private Sub/Function Any()
   Const PROC = "the name of the procedure" ' for the identification of the error source
   On Error Goto on_error ' obligatory anyway
   
   .... any code

exit_proc:
   Exit Sub/Function
   
on_error:
   ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub/Function
```
displays:

**without** the use of the **Alternative  MsgBox**

when the **Alternative  MsgBox** is used

### Usage which provides a "path to the error"

### Usage which supports **debugging** 
The combination _mErrHndlr_ module plus _fMsg_ UserForm offers an elegant way to identify the code line which caused the error. When the Conditional Compile Argument `Debugging = 1` the error message is displayed with an additional Resume button
![image](../Assets/ErrrorMessageWithResumeButton.png)
of which the xaption is returned when clicked. This return can be used as follows:
```vbs
on_error:
   If ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeButton _
   Then Stop: Resume ' F8 leads to the error line
Exit Sub/Function
```
With the advantages of the **Alternative VBA MsgBox** provided by the _fMsg_ UserForm) there may be other additionally specified buttons for a  user's choice.

### Usage which supports **regression testing**

When a number of test procedures are to be executed one after the other an expected behaviour or result may be asserted by means of `Debug.Assert <expression returning true or false>`.<br> An error condition however would stop the execution. Using a specific "Resume Next" and/or "Continue with next test" allows to perform a `Resume Next` or `Goto exit_proc`. The _ErrHndlr_ has these extra buttons built in which are displayed only when the Conditional Compile Argument `Regression = 1`. The following is an example of the regression test for the _ErrHndlr_ function:


```vbs
' still to be included here!
```

#### Difference in display of the error message
##### Using the VB MsgBox
![](../Assets/ErrorMsgMsgBox.png)
##### Using the Alternative VB MsgBox
![](../Assets/ErrMsgAlternativeMsgBox.png)

### Development, test, maintenance
- The dedicated _Common Component Workbook_ ErrHndlr.xlsm is the development, test, and maintenance environment (see the Guthub repo [Common-VBA-Errror-Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler).
- The module _mTest_ contains all test procedures

### Optional execution time trace
When the Conditional Compile Argument `ExecTrace =1` and the [_Entry Procedure_](#entry-procedure) is reached the below kind of execution trace is displayed in the VBE immediate window
![image](../Assets/ExectionTrace.png)

### The _Entry Procedure_
In a call hierarchy the topmost procedure with a BoP/EoP code line (see below) is called the _Entry Procedure_. Usually it is the procedure which is directly or indirectly initiated by a user's  action or an event like Workbook_Open or Workbook_Change.<br>
The indication of the _Entry Procedure_ is essential for the display of the path to the error and the optional display of the execution trace.
```vbs
Private/Public Sub/Function
   Const PROC = "procedure-name"
   ...
   BoP ErrSrc(PROC) ' Begin of Procedure
   ....
   EoP ErrSrc(PROC)
   Exit Sub)Function
   
on_error:
   .....
End Sub/Function
```
and the function (obligatory in each module):
```vbs
Private Function ErrSrcByVal s As String) As String
   ErrSrc = "modile-name." & s
End Function
```