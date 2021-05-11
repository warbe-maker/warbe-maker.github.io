---
layout: post
title: Common VBA Error Services
date:          2021-01-16
categories: vba common error handling
modified_date: 2021-04-29
---
Error services inspired by the best of the web. Not a tutorial but a complete description of the services and how to use them.
<!--more-->

## Services
### The _ErrMsg_ service
 displays a well structured error message with
  - the **[type of the error](#error-types)** by distinguishing [Application error](#the-apperr-service), VB Runtime error, and Database error 
  - the description of the error (_err.Description_),
  - the **[error source](#the-error-source)**,
  - the **[path to the error](#the-bopeop-service-for-the-path-to-the-error)** provided the **[_Entry Procedure_](#the-entry-procedure-for-the-path-to-the-error)** is known, 
  - an optional **[additional information about an error](#error-description)** Services
  - (almost) any number of **[free specified buttons](#free-buttons-specification)**
  - the error line when available
- waits for the user's button clicked and provides/returns [the reply button's value](#processing-reply-buttons) to the caller.

The _ErrMsg_ service has the following syntax (error description and error line are obtained from the err object)
```VB
    On Error Goto eh
    ' .....
eh: mErH.ErrMsg error-source[, buttons]
```
The _ErrMsg_ service has these named arguments:

|  Argument   |   Description   |
| ----------- | --------------- |
| err_source  | Obligatory, string expression providing \<module>.\<procedure>, see [ErrSrc(PROC)](#the-error-source).   |
| err_buttons | Optional. Variant. Defaults to "Terminate execution" button when omitted.<br>May be a value for the VBA MsgBox [_Buttons_][7] argument and/or any descriptive button caption string (including line breaks for a multi-line caption. The buttons may be provided as a comma delimited string, a collection or a dictionary. vbLf items display the following buttons in a new row. |

### The _AppErr_ service
In order to not confuse errors raised with `err.Raise ...` the _AppErr_ service adds  the [_vbObjectError_][10] constant to a given positive number to turn it into a negative. An advantage by the way: Each procedure can have it's own positive error numbers ranging from 1 to n with `err.Raise mErH.AppErr(n)`. The _ErrMsg_ service, when detecting a negative error number uses the _AppErr_ service to turn it back into it's original positive error number.

### The _BoP/EoP_ service for the path-to-the-error
The _ErrMsg_ service only displays a path to the error when an [_Entry Procedure_](#the-entry-procedure-for-the-path-to-the-error) has been indicated. The path to the error is assembled when the error passed on from the error source back up to the _Entry Procedure_ where the error is displayed when reached.

The _BoP/EoP_ services have the following syntax:<br>
`mErH.BoP procedure-id[, arguments]`<br>
`mErH.EoP procedure-id`<br>
with the following named arguments:

| Service |   Argument   |   Description   |
| ------- | ------------ | --------------- |
| BoP     | bop_id       | Obligatory, String expression, unique identification of the procedure in the module (see [ErrSrc(PROC)](#the-error-source)) |
| BoP     | bop_arguments| Optional, ParamArray, a list of the procedures argument, optionally paired as name, value |
| EoP     | eop_id       | Obligatory, String expression, unique identification of the procedure name in the module (see [ErrSrc(PROC)](#the-error-source)) |

Note: When the error message not just allows one reply but provides several buttons (e.g. when the [debugging service](#the-debugging-service) is active), the error message is displayed immediately with the procedure which caused the error. In this case the path to the error is composed from a stack which is maintained along with each BoP/EoP statement. I.e. the path to the error contains only procedures which do use BoP/EoP statements.

### The _Debugging_ service
With the _Conditional Compile Argument_ `Debugging = 1` the error message is displayed with 3 additional buttons which allow:
- `Stop: Resume`,
- `Resume Next`,
- or `Goto xt`(clean exit and continue)

See also [Using the debugging service](#using-the-debugging-service)

### The _BoTP_ (begin of Test Procedure) service for automating regression tests
An - preferably automated - regression test will execute a series of test procedures. Any interruption other than one caused by a failed assertion n assertion should thus be avoided. The _BoTP_ allows the specification of **asserted error numbers** for procedures testing error conditions. For an error number indicated 'asserted' the _ErrMsg_ service bypassed the display of the error message.

The BoTP service has the following Syntax:<br>
`BoTP procedure-id, err-number[, err-number] ...`
with the following named arguments:

|      Argument     |   Description   |
| ----------------- | --------------- |
| botp_id           | Obligatory, Expression providing a unique name of the procedure, e.g.<br>[ErrSrc(PROC)](#the-error-source) |
| botp_errs_asserted| Obligatory, ParamArray with positive numbers |


## Installation
- Download and import the module  [_mErH_][1]
- Download the UserForm [fMsg.frm][2] and [fMsg.frx][3] and import _fMsg.frm_
- Download and import [mMsg.bas][4]
- Since the extra effort is very little, by the way installing the _Common VBA Execution Trace Service_ is worth being considered:<br> Download [mTrc.bas][5] and import it.<br>
When the _mTrc_ is installed and the _Conditional Compile Argument_ 'ExecTrace = 1' an execution trace is displayed whenever the processing reaches an [_Entry Procedure_](#the-entry-procedure). The trace includes all procedures executed which do have a [BoP/EoP statement](#the-bopeop-service-for-the-path-to-the-error).

### Installation of the Execution Trace
Download and import the module  [_mTrc_][5]. 

## Usage
### Usage of the _ErrMsg_ service
The below code works but does not provide a path to the error.

```
Public/Private Sub/Function Any()
   Const PROC = "Any" ' identification of the error source and (if used) the execution trace
   On Error Goto eh ' obligatory anyway
   
   .... any code

xt: Exit Sub/Function
   
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub/Function
```

To get the path to the error displayed with the error message required to additional code lines - at least in the entry procedure:

```VB
Public/Private Sub/Function Any()
   Const PROC = "Any" ' identification of the error source and (if used) the execution trace
   On Error Goto eh ' obligatory anyway
   mErH.BoP ErrSrc(PROC) ' indicates the beginning of the procedure
   .... any code

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub/Function
   
eh: mErH.ErrMsg ErrSrc(PROC) ' indicates the end of the procedure
End Sub/Function
```
displays for example:<br>
![](../Assets/ErrMsgAlternativeMsgBox.png)
![](/Assets/ErrMsgAlternativeMsgBox.png)

### Using the _Debugging_ service
When code lines are not numbered it appears that there is no way for identifying the error line - potentially extending unproductive time for error chasing. The below 'trick' provides a true godsend in case:

```VBS
eh:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErH.ErrMsg ....
End Sub/Function
```

The _ErrMsg_ service has this 'trick' already built-in:<br>When the _Conditional Compile Argument_ `Debugging = 1` the error message is displayed with 3 additional buttons:<br>
![](../Assets/ErrMsgWithDebuggingOption.png)
![](/Assets/ErrMsgWithDebuggingOption.png)<br>

and the extra reply buttons can be used as follows:
```
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt1ResumeNext: Resume Next
    End Select
```

or alternatively:

```vbs
eh: If ErrMsg(ErrSrc(PROC)) = mErH.DebugOpt1ResumeError _
    Then Stop: Resume ' F8 leads to the error line
Exit Sub/Function
```
Note that the additional reply buttons are provided as public Properties.<br>In production the _Conditional Compile Argument_ `Debuggin = 0` prevents the display of the debugging buttons.

### Regression test support
I am convinced of proper regression testing. It should be obligatory specially for _Common Components_ potentially used in any number of VB-Projects (like this _mErH_ module for instance). However, any test of an error condition would interrupt an - otherwise perfectly automated - regression test with the display of the tested error message. With the _BoTP_ service instead of the _BoP_ service ***asserted error numbers*** may be specified in order to have the display of the error message bypassed. Clever used, this bypassing can be used for regression only, which displays the error message when the test procedure is executed individually. Below is an example of an execution trace which documents the performed tests:<br>
![](/Assets/ExecTraceRegressionTest.png)
![](../Assets/ExecTraceRegressionTest.png)
Note: The fully automated regression test can be found in the the corresponding Github repo ["Common VBA Error Services"][8] or by just downloading the [mTest.bas][15] code module of in [ErH.xlsm][16] Workbook.

### Execution trace support
When the _mTrc_ module is imported and the _Conditional Compile Argument_ `ExecTrace = 1` any arguments provided with the _BoP_  service appear in the execution trace displayed when the entry procedure is reached.
```
Private Sub Any(ByVal arg1 As String, ByVal arg2 as Single)

    On Error Goto eh
    Const PROC = "Any"
    BoP ErrSrc(PROC), "arg1=", arg1, "arg2=", arg2

    ' .... ' any code

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg ErrSrc(PROC)
End Sub
```

## Other details

### Error types
The error handler distinguishes between
- [Application error](#the-application-error-service)<br>provided the error had been raised by `err.Raise mErH.AppErr(n) ...` with n = 1 to 2147221503 
- VB Runtime error
- Database error

### The _error source_
The following procedure will be copied to any in component the _mErH.ErrMsg_ service is used in order to identify any procedure:

```VB
Private Function ErrSrc(ByVal s As String) As String
   ErrSrc = "\<component-name>." & s
End Function
```

### The _Entry Procedure_ for the path-to-the-error
The top level procedure, i.e. the first one with a _BoP/EoP_ service statement is considered the _Entry Procedures_. Usually those procedures are event procedures like _Workbook\_Open_ or click events in a UserForm.

### Error description with additional information
The _ErrMsg_ service displayes the information provided by  _err.Description_. For an _Application Error_ the error description is what is provided with the `err.Raise` statement. When the error description looks like "This is a serious error.||This error may be avoided by ...." the string concatenated with || is regarded an additional information and will be displayed in the error message as such.

### Processing reply buttons
The _ErrMsg_ service returns the 'value' of the clicked button. This value may be _vbYesNo_, _vbIgnoreRetryCancel_, etc. or the caption string of the displayed button (including any vbLf!). This returned value may be processed as follows:
```VB
eh: Select Case mErH.ErrMsg(ErrSrc(PROC)
        Case ....
        Case ...
    End Select
```

### Free buttons specification
Because the _ErrMsg_ service uses the _[Common VBA Message Services][12]_ UserForm _fMsg_ for the display of the error message specifying any other but the default button is extremely flexible. Buttons can be provided as a comma delimited string, an array, a Collection or a Dictionary whereby each of the items may be a VBA MsgBox value, a button's caption string, or a vbLf indicating the following buttons are displayed in a new row. Specifying the buttons is supported by  module. Below is an example which displays and error message with an **Ok** button and a **My button** button by using the the _[Buttons][14]_ service of the _mMsg_ module:
```VB
Private Sub Any()
    Const PROC = "Any"
    On Error Goto eh

    
    mErH.BoP ErrSrc(PROC)
    ' .... any code   
    Err.Raise AppErr(1), ErrSrc(PROC), "Display of a free defined button in addition to the usual Ok button (resumes the error when clicked)"

    ' .... any code    

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(err_source:=ErrSrc(PROC) _
                          , err_buttons:=mMsg.Buttons(vbOKOnly, vbLf, "My button" _
                           )
        Case "My button": ' any action
    End Select
End Sub
```
displays:<br>
![](../Assets/FreeButtonSpecification.png)
![](/Assets/FreeButtonSpecification.png)<br>

See also the [Common VBA Message Services][6] post for more details on how to use the buttons argument, specifically by means of the _Buttons_ service.

### Making use of the optional Execution Trace
When the execution trace module _mTrc_ is used together with the _mErH_ services there is little to effort required. Any executed procedure with an<br> `mErH.BoP ErrSrc(PROC)`<br>at the beginning and an<br> `mErH.EoP ErrSrc(PROC)` <br>statement at the end of a procedure will be included in the displayed trace result.<br>
Note: When the _[Common VBA Execution Trace Service][11]_ had already been used before the _mErH_ module had been installed all _mTrc.BoP/mTrc.EoP_ have to be changed to _mErH.BoP/mErH.EoP_ . Any _mTrc.BoC/mTrc.EoC_ are ok.

### _Compact_ (default) versus _Detailed_ trace result
The default is a trace display like the following:<br>
![](../Assets/ExecutionTrace.png)
![](/Assets/ExecutionTrace.png)<br>

However, for those who do not believe in the displayed figures a detailed view may be of interest. With `mTrc.DisplayedInfo = Detailed` (yes, standard modules may have properties but they are just not auto-sensed) the following kind of trace information is displayed:<br>
![](../Assets/ExecutionTraceDetailed.png)
![](/Assets/ExecutionTraceDetailed.png)<br>


## Contribution, development, test, maintenance
The dedicated _Common Component Workbook_ [ErH.xlsm][16] is used for development, test, and maintenance. This Workbook is kept in a dedicated folder which is the local equivalent (in github terminology the clone of the public GitHub repo [Common-VBA-Error-Services][8]. The module [_mTest_][15] contains all obligatory test procedures when the code is modified. Code modifying contributions will be handled by means of a branch which is merged to the master once a code change has successfully passed the regression test.

A code modifying contribution is very well supported by the [GitHub Desktop for Windows][9] which is the environment I use.

[1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/mErH.bas
[2]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/source/fMsg.frm
[3]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/source/fMsg.frx
[4]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/source/mMsg.bas
[5]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/source/mTrc.bas
[6]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
[7]:https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings
[8]:https://github.com/warbe-maker/Common-VBA-Error-Services
[9]:https://desktop.github.com
[10]:https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1
[11]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/14/Common-VBA-Execution-Trace-Service.html
[12]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
[14]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html#the-buttons-service
[15]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/mTest.bas
[16]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/ErH.xlsm