---
layout: post
title: A comprehensive common VBA Error Handler inspired by the best of the web
subtitle: An Error Handler inspired by the best of the web
date: 2020-11-07
categories: vba common
comments: true
---

**This is not a tutorial about error handling** but the description of a full featured ready to use error handler module with an optional execution trace module.

In this post<br>
[Error Handler](#error-handler)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Services](#error-handler-services)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Syntax](#error-handler-syntax)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Installation](#error-handler-installation)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Usage](#error-handler-usage)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Basic usage](#basic-usage)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Usage providing a "path to the error" with the error message](#usage-providing-a-path-to-the-error-with-the-error-message)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Debug supporting usage](#debug-supporting-usage)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Usage supporting test](#usage-supporting-test)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Usage details](#usage-details)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[The _Entry Procedure_](#the-entry-procedure)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Making use of the free buttons](#making-use-of-the-free-buttons)<br><br>
[Execution Trace](#execution-trace)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Service](#execution-trace-service)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Installation](#installation)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Usage](#execution-trace-usage)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[_Compact_ (default) or _Detailed_ trace result](#compact-default-or-detailed-trace-result)<br><br>
[Contribution, development, test, maintenance](#contribution-development-test-maintenance)

## Error Handler
### Error Handler Services
Only a few additional code lines in a procedure unfold the provided services:
- **Path to the error**<br>One advantage of this error handler is the display of the path to the error built/assembled when the error is passed on from the error source procedure back up to [the Entry Procedure](#the-entry-procedure) (provided it is known)
- **Free buttons specification**<br>[Free specified buttons](#free-specified-buttons) displayed with the error message allow an eeror processing based on a user's choice.<br>The [usage which supports debugging](#a-usage-which-supports-debugging) is one already built-in example, another one is the [Usage supporting test](#usage-supportingtest)
- **Error type distinction**<br>The error message distincts between _VB Runtime Error_, _Application Error_, and _Database-Error_
- **Error source and error line**<br>The error message displays the source of the error plus the error line when available
- **Execution time trace (optional module)**<br>Each time when the processing has returned to an [_Entry Procedure_](#the-entry-procedure) an [optional execution time trace](#optional-execution-time-trace) with the precise execution time of each [traced procedure](#) and/or [traced number of code lines](#traced-number-of-code-lines) is displayed in the VBE immediate window
- **Error log**<br>The implementation of an optional error log is a still pending issue

### Error Handler Syntax
```vbs
   mErH.ErrMsg error-number, error-source, error-description, error-line[, buttons]
```
The procedure has these named arguments:

|  Argument  |   Description   |
| ---------- | --------------- |
| errnumber  | Err.Number      |
| errsource  | ErrSrc(PROC).   |
| errdscrptn | Err.Description |
| errline    | Erl             |
| buttons    | Optional. Variant. Defaults to "Terminate execution" button when omitted.<br>May be a [value for the VBA MsgBox buttons argument](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) and/or any descriptive button caption string (including line breaks for a multi-line caption. The buttons may be provided as a comma delimited string, a collection or a dictionary. vbLf items display the following buttons in a new row. |

### Error Handler Installation
- Download and import the module  [_mErH_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mErH.bas)
- Download the UserForm  [fMsg.frm](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm) and   [fMsg.frx](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frx) and import _fMsg.frm_

### Error Handler Usage
#### Basic usage
The below code works but does not provide a path to the error.

```vbs
 Public/Private Sub/Function Any()
   Const PROC = "Any" ' identification of the error source and (if used) the execution trace
   On Error Goto eh ' obligatory anyway
   
   .... any code

xt: Exit Sub/Function
   
eh: mErH.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub/Function
```

displays:
![](../Assets/ErrMsgAlternativeMsgBox.png)
![](/Assets/ErrMsgAlternativeMsgBox.png)


#### Usage providing a "path to the error" with the error message
When the user has no choice because just the default button is displayed with the error message an error is passed on to [the _Entry Procedure](#the-entry-procedure) and thereby the path to the error is assembled.
 

#### Debug supporting usage
One of the most common problems in identifying the code line which caused an error. Without line numbers, the mir lines a procedure has the more difficult. Those familiar with the "trick"

```vbs
eh:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErH.ErrMsg ....
End Sub/Function
```

may appreciate that this is integrated in the _mErH_ module. When the Conditional Compile Argument<br>
`Debugging = 1` an additional button is displayed with the error message:
![](../Assets/ErrrorMessageWithResumeButton.png)
![](/Assets/ErrrorMessageWithResumeButton.png)

and when the button is clicked ...

```vbs

eh: If ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeButton _
    Then Stop: Resume ' F8 leads to the error line
Exit Sub/Function
```

does the job. In production the Conditional Compile Argument `Debuggin = 0` prevents the display of this button.

#### Usage supporting test
For _Common Components_ like this error handler I regard a regression test obligatory before a code modification is published. However, any test of an error condition stops the test process when there is only the default button displayed.

With the Conditional Compile Argument `Test = 1` two additional buttons are displayed: **Continue with next code line** and **Continue with next procedure**

image still missing

and the following can be for a test continuation

```vbs

```
#### Usage details
##### The _Entry Procedure_
The _Entry Procedure_ is the one the error handler recognizes as the begin of the execution of VBA code. I.e. it is the first procedure in a call hierarchy with a pair of BoP/EoP statements. Provided at least one such a procedure has been passed, an error is passed on back up to this _Entry Procedure_ while the _Path to the error_ is assembled. The following is an example code for an _Entry Procedure_ or any procedure which, in case of an error is contained in the [path to the error](#path-to-the-error)

```vbs
Private/Public Sub/Function Any
    On Error Goto eh
    Const PROC = "Any"
    mErH.BoP ErrSrc(PROC) ' Begin of procedure
   ....
   If ... Then Goto xt ' never use Exit! It will bypass the EoP execution
   ....
xt: mErH.EoP ErrSrc(PROC) ' End of procedure
    Exit Sub)Function
   
eh: mErH.ErrMsg .....
End Sub/Function
```
and the function (obligatory in each module):
```vbs
Private Function ErrSrcByVal s As String) As String
   ErrSrc = "modile-name." & s
End Function
```

##### The _Path to the error_

  
##### Making use of the free buttons
The use of the _fMsg_ UserForm in general provides an enormous flexibility regarding the display of buttons. This can be used with the display of an error message to provide the user with any number of choices. Because the error message is fixed it is an advantage that the displayed buttons may have any free multi-line caption text, returned when the button is clicked. Example: The ErrHndlr statement:<br>
```vbs
Private Sub Any()
    Const PROC = "Any"
    On Error Goto eh
    ' .... any code, declarations etc.
    mErH.BoP ErrSrc(PROC)
    ' .... any code   
    Err.Raise AppErr(1), ErrSrc(PROC), "Display of a free defined button in addition to the usual Ok button (resumes the error when clicked)"
    Goto xt
    ' .... any code    

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(Err.Number, ErrSrc(PROC), Err.Description, Erl, buttons:=vbOKOnly & "," & vbLf & ",My button")
        Case "My button": Resume
    End Select
End Sub
```
displays:

![](../Assets/FreeButtonSpecification.png)
![](/Assets/FreeButtonSpecification.png)<br>
<small>Note that the additional button is displayed in a second row due to the vbLf in the buttons argument.</small>

See also the [Alternative VBA MsgBox](https://github.com/warbe-maker/VBA-MsgBox-Alternative) for more details on how to use it and its advantages (not yet available as post).

## Execution Trace
### Execution Trace Service
When the optional module _mTrc_ is installed it provides an execution trace whenever the processing reaches an [_Entry Procedure_](#the-entry-procedure).

### Installation
Download and import the module  [_mTrc_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mTrc.bas) 

### Execution Trace Usage
The execution trace module _mTrc_ may be used without the error handler in which case the usage differs slightly (see []())
Set the Conditional Compile Argument `ExecTrace = 1` and make sure all BoP/EoP are qualified with `mErH.` That's it. Any executed procedure with an<br> `mErH.BoP ErrSrc(PROC)`<br>at the beginning and an<br> `mErH.EoP ErrSrc(PROC)` <br>code lines at the end of a procedure will be included in the displayed trace result.

#### _Compact_ (default) or _Detailed_ trace result
The default is a trace display like the following:
![](../Assets/ExecutionTrace.png)
![](/Assets/ExecutionTrace.png)<br>

However, for those who do not believe in the displayed figures a detailed view may be of interest. With `mTrc.DisplayedInfo = Detailed` (yes, standard modules may have properties but they are just not auto-sensed) the following kind of trace information is displayed:
![](../Assets/ExecutionTraceDetailed.png)
![](/Assets/ExecutionTraceDetailed.png)<br>

#### The Seconds Precision Property
still to be completed

### Contribution, development, test, maintenance
It had become a habit: A dedicated _Common Component Workbook_ **ErH.xlsm** is used for development, test, and maintenance. This Workbook is kept in a dedicated folder which is the local equivalent (in github terminology the clone of the public [GitHub repo Common-VBA-Errror-Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler). The module **_mTest_** contains all obligatory test procedures when the code is modified, the module **_mDemo_** all procedures for the images in this post. The modules **_mErH_** and **_fMsg_** are downloaded from this source. Thus, it is wise not to make any changes without specifying a branch which is merged to the master once a code change has finished and successfully tested.

Those interested not only in using the Error Handler but also modify or even contribute in improving it may fork or clone it to their own computer which is very well supported by the [GitHub Desktop for Windows](https://desktop.github.com). That's my environment for a continuous improvement process.