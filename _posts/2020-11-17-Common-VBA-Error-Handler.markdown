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
[Services](#services)<br>
[Syntax of the _ErrMsg_ function](#syntax-of-the-errmsg-function)<br>
[Installation](#installation)<br>
[Usage](#usage)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Basic usage](#basic-usage)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Usage providing a "path to the error" with the error message](#usage-providing-a-path-to-the-error-with-the-error-message)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Debugging and  test support](#debugging-and-test-support)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Test supporting](#test-support)<br>
[Usage/services details](#usage-services-details)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[The Path to the error](#the-path-to-the-error)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[The _Entry Procedure_](#the-entry-procedure)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Free buttons specification](#free-buttons-specification)<br>
[Optional Execution Trace Service](#optional-execution-trace-service)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Installation of the Execution Trace](#installation-of-the-execution-trace)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[Usage of the Execution Trace](#usage-of-the-execution-trace)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[_Compact_ (default) versus _Detailed_ trace result](#compact-default-versus-detailed-trace-result)<br>
&nbsp;&nbsp;&nbsp;&nbsp;[The Seconds Precision Property](#the-seconds-precision-property)<br><br>
[Contribution, development, test, maintenance](#contribution-development-test-maintenance)

## Services
### Common service
The main services are provided by the _ErrMsg_ function of the mErH_ module which
- displays an structured error message with
  - **[the type of the error](#the-type-of-the-error)**, 
  - the description of the error,
  - the error source,
  - the **[path to the error](#the-path-to-the-error)** provided **[the _Entry Procedure_](#the-entry-procedure)** is known, 
  - an optional **[additional information about the error](#additional-information-about-the-error)**,
  - (almost) any number of **[Free specified buttons](#free-buttons-specification)**
- waits for the user's button clicked and provides/returns [the reply buttons value](#processing-the-returned-reply) to the caller.

### Debugging and test support
- When the Conditional Compile Argument `Test = 1`:
  - two additional [test option buttons](#the-test-option-buttons) are displayed<br>![image](../Assets/ErrMsgWithTestOption.png)
  - When _asserted error numbers are specified for a test procedure the corresponding error messages are not displayed but processing continues which perfectly supports testing of error conditions within a regression test.
- When the Conditional Compile Argument `Debuggig = 1` two additional buttons support identifying the error line or continue<br>
![image](../Assets/ErrMsgWithDebuggingOption.png)


## Syntax of the _ErrMsg_ function
```vbs
   mErH.ErrMsg error-number, error-source, error-description, error-line[, buttons]
```
The procedure has these named arguments:

|  Argument   |   Description   |
| ----------- | --------------- |
| err_number  | Optional, defaults to err.Number when omitted      |
| err_source  | Obligatory, string expression providing <module>.<procedure>, e.g. ErrSrc(PROC).   |
| err_dscrptn | Optional, defaults to err.Description when omitted |
| err_line    | Optional, defaults to  Erl when omitted            |
| err_buttons | Optional. Variant. Defaults to "Terminate execution" button when omitted.<br>May be a [value for the VBA MsgBox buttons argument](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) and/or any descriptive button caption string (including line breaks for a multi-line caption. The buttons may be provided as a comma delimited string, a collection or a dictionary. vbLf items display the following buttons in a new row. |

## Installation
- Download and import the module  [_mErH_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mErH.bas)
- Download the UserForm  [fMsg.frm](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm) and   [fMsg.frx](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frx) and import _fMsg.frm_

## Usage
### Basic usage
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
### Usage/services details
#### The type of the error
The error handler distinguishes between
- Application error
- VB Runtime error
- Database error

#### The _Entry Procedure_
The procedure which the error handler has recognized as the top level procedure of a call hierarchy by means of a pair of BoP/EoP statements. Provided at least one such procedure has been passed and the user has no reply choices since only one button is


, an error is passed on back up to this _Entry Procedure_ while the _Path to the error_ is assembled. The following is an example code for an _Entry Procedure_ or any procedure which, in case of an error is contained in the [path to the error](#path-to-the-error)

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

##### The test option buttons
When the Conditional Compile Argument `Test = 1` the error message looks as follows:
![](/Assets/ErrMsgWithTestOption.png)
![](../Assets/ErrMsgWithTestOption.png)

##### The debugging option buttons
When the Conditional Compile Argument `Debugging = 1` the error message looks as follows:
![](/Assets/ErrMsgWithDebuggingOption.png)
![](../Assets/ErrMsgWithDebuggingOption.png)
The additional button have an advantage over the equivalent:
```
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
```
because this cannot be altered which means it loops until the reason for the error has been eliminated which may lead to an emergency code change just to continue without an error.

With both Conditional Compile Arguments `Test = 1` and `Debugging = 1` four additional buttons are displayed.

##### Free buttons specification
Buttons can be provided as a comma delimited string, an array, a Collection or a Dictionary whereby the items are a VBA MsgBox value, a button's caption string, or a vbLf indicating the following buttons are displayed in a new row. This free buttons specification is a service provided by the used fMsg UserForm, a Common VBA Message Form.
The below example of an _ErrMsg_:
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

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC) err_buttons:=vbOKOnly & "," & vbLf & ",Resume Error")
        Case "Resume Error": Stop: Resume
    End Select
End Sub
```
displays:

![](../Assets/FreeButtonSpecification.png)
![](/Assets/FreeButtonSpecification.png)<br>
<small>Note that the additional button is displayed in a second row due to the vbLf in the buttons argument.</small>

See also the [Alternative VBA MsgBox](https://github.com/warbe-maker/VBA-MsgBox-Alternative) for more details on how to use it and its advantages (not yet available as post).

## Optional Execution Trace Service
When the optional module _mTrc_ is installed it provides an execution trace whenever the processing reaches an [_Entry Procedure_](#the-entry-procedure).

## Installation of the Execution Trace
Download and import the module  [_mTrc_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mTrc.bas) 

## Usage of the Execution Trace
The execution trace module _mTrc_ may be used without the error handler in which case the usage differs slightly (see []())
Set the Conditional Compile Argument `ExecTrace = 1` and make sure all BoP/EoP are qualified with `mErH.` That's it. Any executed procedure with an<br> `mErH.BoP ErrSrc(PROC)`<br>at the beginning and an<br> `mErH.EoP ErrSrc(PROC)` <br>code lines at the end of a procedure will be included in the displayed trace result.

### _Compact_ (default) versus _Detailed_ trace result
The default is a trace display like the following:
![](../Assets/ExecutionTrace.png)
![](/Assets/ExecutionTrace.png)<br>

However, for those who do not believe in the displayed figures a detailed view may be of interest. With `mTrc.DisplayedInfo = Detailed` (yes, standard modules may have properties but they are just not auto-sensed) the following kind of trace information is displayed:
![](../Assets/ExecutionTraceDetailed.png)
![](/Assets/ExecutionTraceDetailed.png)<br>


## Contribution, development, test, maintenance
The dedicated _Common Component Workbook_ **ErH.xlsm** is used for development, test, and maintenance. This Workbook is kept in a dedicated folder which is the local equivalent (in github terminology the clone of the public [GitHub repo Common-VBA-Errror-Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler). The module **_mTest_** contains all obligatory test procedures when the code is modified, the module **_mDemo_** all procedures for the images in this post. The modules **_mErH_** and **_fMsg_** are downloaded from this source. Thus, it is wise not to make any changes without specifying a branch which is merged to the master once a code change has finished and successfully tested.

Those interested not only in using the Error Handler but also modify or even contribute in improving it may fork or clone it to their own computer which is very well supported by the [GitHub Desktop for Windows](https://desktop.github.com). That's my environment for a continuous improvement process.