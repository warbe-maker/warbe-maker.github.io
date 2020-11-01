---
layout: post
title: A comprehensive common VBA Error Handler inspired by the best of the web
subtitle: An Error Handler inspired by the best of the web
date: 2020-10-02 16:00 +0200
categories: vba common
---

**This is not a tutorial about error handling** but the description of a full featured ready to use error handler module.

In this post<br>
[Error Handler](#error-handler)<br>
&nbsp;&nbsp;&nbsp;[Services](#error-handler-services)<br>
&nbsp;&nbsp;&nbsp;[Syntax](#error-handler-syntax)<br>
&nbsp;&nbsp;&nbsp;[Installation](#error-handler-installation)<br>
&nbsp;&nbsp;&nbsp;[Usage](#error-handler-usage)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Basic usage](#basic-usage)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Usage providing a "path to the error" with the error message](#usage-providing-a-path-to-the-error-with-the-error-message)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Debug supporting usage](#debug-supporting-usage)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Usage supporting test](#usage-supporting-test)<br>
[Usage details](#usage-details)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Tracing procedure and code execution](#tracing-procedure-and-code-execution)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[The _Entry Procedure_](#the-entry-procedure)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Making use of the free buttons](#making-use-of-the-free-buttons)<br>
[Execution Trace](#execution-trace)<br>
&nbsp;&nbsp;&nbsp;[Service](#execution-trace-service)<br>
&nbsp;&nbsp;&nbsp;[Installation](#execution-trace-installation)<br>
&nbsp;&nbsp;&nbsp;[Usage](#execution-trace-usage)<br>
[Contribution, development, test, maintenance](#contribution-development-test-maintenance)

## Error Handler
### Error Handler Services
Only a few additional code lines in a procedure unfold the provided services:
- **Path to the error**<br>One advantage of the _ErrHndlr_ is the display of the path to the error built/assembled when the error is passed on from the error source procedure back up to [the Entry Procedure](#the-entry-procedure) (provided it is known)
- **Free buttons specification**<br>[Free specified buttons](#free-specified-buttons) displayed with the error message allow an eeror processing based on a user's choice.<br>The [usage which supports debugging](#a-usage-which-supports-debugging) is one already built-in example, another one is the [Usage supporting test](#usage-supportingtest)
- **Error type distinction**<br>The error message distincts between _VB Runtime Error_, _Application Error_, and _Database-Error_
- **Error source and error line**<br>The error message displays the source of the error plus the error line when available
- **Execution time trace**<br>Each time when the processing has returned to an [_Entry Procedure_](#the-entry-procedure) an [optional execution time trace](#optional-execution-time-trace) with the precise execution time of each [traced procedure](#) and/or [traced number of code lines](#traced-number-of-code-lines) is displayed in the VBE immediate window
- **Error log**<br>The implementation of an optional error log is a still pending issue

### Error Handler Syntax
```vbs
ErrHndlr error-number, error-source, error-description, error-line[, buttons]
```
The procedure has these named arguments:

|  Argument  |   Description   |
| ---------- | --------------- |
| errnumber  | Err.Number      |
| errsource  | ErrSrc(PROC).   |
| errdscrptn | Err.Description |
| errline    | Erl             |
| buttons    | Optional. Variant. Defaults to "Terminate execution" button when omitted.<br>May be a [value for the VBA MsgBox buttons argument](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) and/or any descriptive button caption string (including line breaks for a multi-line caption. The buttons may be provided as a comma delimited string, a collection or a dictionary. vbLf items display the following buttons in a new row. |

<<<<<<< HEAD
### Error Handler Installation
- Download and import [_mErrHndlr_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mErrHndlr.bas)
- Download [fMsg.frm](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm)
- Download  [fMsg.frx](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frx)
- Import _fMsg.frm_ 

=======
### Installation of the Error Handler
>>>>>>> branch 'master' of https://github.com/warbe-maker/warbe-maker.github.io
Note: This error handler only unfolds all its advantages with the _Alternative VBA MsgBox_. Effort spent in allowing a usage merely based on the VBA MsgBox has been stopped because of it's constraints.
<<<<<<< HEAD
 
### Error Handler Usage
=======
- Download and import the module  [_mErrHndlr_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mErrHndlr.bas)
- Download the UserForm  [fMsg.frm](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm) and   [fMsg.frx](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frx) and import _fMsg.frm_

Optionally, when an execution trace is desired (will only become active when the Conditional Compile Argument `ExecTrace = 1`):
- Download and import the module  [_mTrace_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mTrace.bas)

### Usage
>>>>>>> branch 'master' of https://github.com/warbe-maker/warbe-maker.github.io
#### Basic usage
The below code works but does not provide a path to the error.

```vbs
 Public/Private Sub/Function Any()
   Const PROC = "the name of the procedure" ' for the identification of the error source
   On Error Goto on_error ' obligatory anyway
   
   .... any code

exit_proc:
   Exit Sub/Function
   
on_error:
   mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub/Function
```

displays:
![](../Assets/ErrMsgAlternativeMsgBox.png)
![](/Assets/ErrMsgAlternativeMsgBox.png)
![](../Assets/ErrMsgAlternativeMsgBox.png)

#### Usage providing a "path to the error" with the error message
When the user has no choice because just the default button is displayed with the error message an error is passed on to [the _Entry Procedure](#the-entry-procedure) and thereby the path to the error is assembled.
 

#### Debug supporting usage
One of the most common problems in identifying the code line which caused an error. Without line numbers, the mir lines a procedure has the more difficult. Those familiar with the "trick"

```vbs
on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErrHndlr.ErrHndlr ....
End Sub/Function
```

may appreciate that this is integrated in the _mErrHndlr_ module. When the Conditional Compile Argument<br>
`Debugging = 1` an additional button is displayed with the error message:
![](../Assets/ErrrorMessageWithResumeButton.png)
![](/Assets/ErrrorMessageWithResumeButton.png)

and when the button is clicked ...

```vbs
on_error:
   If ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl) = ResumeButton _
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

### The _Entry Procedure_
In a call hierarchy the topmost procedure with a BoP/EoP code line (see code sample below) is called the _Entry Procedure_. Usually it is the procedure which is directly or indirectly initiated by a user's  action or an event. The indication of an _Entry Procedure_ is essential for the display of the **path to the error** as for the optional display of the [execution trace](#tracing-procedure-and-or-code-execution) .
```vbs
Private/Public Sub/Function
   Const PROC = "procedure-name"
   ...
   BoP ErrSrc(PROC) ' Begin of procedure
   ....
   EoP ErrSrc(PROC) ' End of procedure
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
### Tracing procedure and code execution
Provided the Conditional Compile Argument `ExecTrace =1`, whenever the processing reaches an [_Entry Procedure_](#entry-procedure) the execution trace is displayed in the VBE immediate window. Performance issues may require a more detailed execution tracing than just complete procedures. A pair of BoT/EoT statements may surround any number of code lines within a procedure as follows:
```vbs
    BoT "my code lines" ' Begin of trace
    .... ' any code lines
    EoT "my code lines" ' End of trace (string must mathc with BoT statement)
```
Example:
![](../Assets/ExecTraceOfCodeLines.png)
![](/Assets/ExecTraceOfCodeLines.png)
<<<<<<< HEAD

=======
>>>>>>> branch 'master' of https://github.com/warbe-maker/warbe-maker.github.io
### Making use of the free buttons
The use of the _fMsg_ UserForm in general provides an enormous flexibility regarding the display of buttons. This can be used with the display of an error message to provide the user with any number of choices. Because the error message is fixed it is an advantage that the displayed buttons may have any free multi-line caption text, returned when the button is clicked. Example: The ErrHndlr statement:<br>
```vbs
Private Sub Demo_7_Free_Button_Display()

    On Error GoTo on_error
    Const PROC = "Demo_7_Free_Button_Display"

    Err.Raise AppErr(1), ErrSrc(PROC), "Display of a free defined button in addition to the usual Ok button (resumes the error when clicked)"
    Exit Sub

on_error:
    Select Case mErrHndlr.ErrHndlr(Err.Number, ErrSrc(PROC), Err.Description, Erl, buttons:=vbOKOnly & "," & vbLf & ",My button")
        Case "My button": Resume
    End Select
End Sub
```
displays:

![](../Assets/FreeButtonSpecification.png)
![](/Assets/FreeButtonSpecification.png)<br>
<small>Note that the additional button is displayed in a second row due to the vbLf in the buttons argument.</small>

See also the [Alternative VBA MsgBox](https://github.com/warbe-maker/VBA-MsgBox-Alternative) for more details on how to use it and its advantages (not yet available as post).

### Contribution, development, test, maintenance
It had become a habit: A dedicated _Common Component Workbook_ **ErrHndlr.xlsm** is used for development, test, and maintenance. This Workbook is kept in a dedicated folder which is the local equivalent (in github terminology the clone of the public [GitHub repo Common-VBA-Errror-Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler). The module **_mTest_** contains all obligatory test procedures when the code is modified, the module **_mDemo_** all procedures for the images in this post. The modules **_mErrHndlr_** and **_fMsg_** are downloaded from this source. Thus, it is wise not to make any changes without specifying a branch which is merged to the master once a code change has finished and successfully tested.

Those interested not only in using the Error Handler but also modify or even contribute in improving it may fork or clone it to their own computer which is very well supported by the [GitHub Desktop for Windows](https://desktop.github.com). That's my environment for a continuous improvement process.

## Execution Trace
### Execution Trace Service
The execution trace is an optional component of the error handler.
### Execution Trace Installation
### Execution Trace Usage
### Execution Trace Properties
### Precision Property
### Display Property
#### Compact DisplaY (default)
#### Detailed Display
