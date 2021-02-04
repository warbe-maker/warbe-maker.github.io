---
layout: post
title: Common VBA Error Services (inspired by the best of the web)
subtitle: Comprehensive error handling inspired by the best of the web
date: 2021-01-16
categories: vba common error handling
---

**This is not a tutorial about error handling** but the description of  comprehensive, full featured, ready to use error services module.

## Services

### The _ErrMsg_ service
- displays a well structured error message with
  - the **[type of the error](#error-types)** (Application error, VB Runtime error, and Database error) 
  - the description of the error (_err.Description_),
  - the **[error source](#the-error-souce)**,
  - the **[path to the error](#the-bopeop-path-to-the-error) service** provided the **[_Entry Procedure_](#the-entry-procedure)** is known, 
  - an optional **[additional information about an error](#additional-information-about-an-error)**,
  - (almost) any number of **[free specified buttons](#free-buttons-specification)**
  - the error line when available
- waits for the user's button clicked and provides/returns [the reply button's value](#processing-the-returned-reply) to the caller.

The _ErrMsg_ service has the following syntax (error description and error line are obtained from the err object)
```VB
    On Error Goto eh
    ` .....
eh: mErH.ErrMsg error-source[, buttons]
```
The _ErrMsg_ service has these named arguments:

|  Argument   |   Description   |
| ----------- | --------------- |
| err_source  | Obligatory, string expression providing \<module>.\<procedure>, see [ErrSrc(PROC)](#the-error-source).   |
| err_buttons | Optional. Variant. Defaults to "Terminate execution" button when omitted.<br>May be a value for the VBA MsgBoc [_Buttons_][7] argument and/or any descriptive button caption string (including line breaks for a multi-line caption. The buttons may be provided as a comma delimited string, a collection or a dictionary. vbLf items display the following buttons in a new row. |

### The _AppErr_ service
The _ErrMsg_ service recognizes an _Application Error_ i.e. an error explicitly raised via `err.Raise` through a negative error number as suggested by VBA because the _AppErr_ service adds the [_vbObjectError_][10] constant to a given positive number to turn it into a negative, thereby preventing any confusion with VB Runtime errors. An advantage by the way: Each procedure can have it's own positive error numbers ranging from 1 to n with `err.Raise mErH.AppErr(n)`. The _ErrMsg_ service, when detecting a negative error number uses the _AppErr_ service to turn it back into it's original positive error number.

### The _BoP/EoP_ (path to the error) service
The _ErrMsg_ service only displays a path to the error when the [_Entry Procedure_](#the-entry-procedure) has been indicated. The path to the error is assembled when the error passed on from the error source back up to the _Entry Procedure_ where the error is displayed when reached.

The _BoP_ / _EoP_ services have the following syntax:<br>
`BoP procedure-id[, arguments]`<br>
`EoP procedure-id`<br>
with the following named arguments:

| Service |   Argument   |   Description   |
| ------- | ------------ | --------------- |
| BoP     | bop_id       | Obligatory, Expression providing a unique name of the procedure, e.g.<br>[ErrSrc(PROC)](#the-error-source) |
| BoP     | bop_arguments| Optional, ParamArray, a list of the procedures argument, optionally paired as name, value |
| EoP     | eop_id       | Obligatory, Expression providing a unique name of the procedure, e.g.<br>[ErrSrc(PROC)](#the-error-source) |

Note: When the user not only has one reply button but several reply choices (see the debugging service for instance), the error message is displayed immediately with the procedure which caused the error. In this case the path to the error is composed from a stack which is maintained along with each BoP/EoP statement. I.e. the path to the error contains only procedures which do use BoP/EoP statements.

### The debugging service for identifying an error line
With the _Conditional Compile Argument_ `Debuggig = 1` the error message is displayed with two additional buttons which allow a `Stop: Resume` reaction which leads to the code line the error occurred (see  [Usage of the debugging service](#usage-of-the-debugging-service-identifying-the-error-line-when-code-lines-are-not-numbered))

### The _BoTP_ service for automating regression tests
An - preferably automated - regression test will execute a series of test procedures. Any interruption other than one caused by a failed assertion for an assertion should thus be avoided. The _BoTP_ allows the specification of **asserted error numbers** for procedures testing error conditions. For an error number indicated 'asserted' the _ErrMsg_ service bypassed the display of the error message.

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
- Since the extra effort is very little, by the way installing the _Common VBA Execution Trace Service_ is worth being concidered:<br> Download [mTrc.bas][5] and import it.

## Usage
### Usage of the _ErrMsg_ service
The below code works but does not provide a path to the error.

```vbs
Public/Private Sub/Function Any()
   Const PROC = "Any" ' identification of the error source and (if used) the execution trace
   On Error Goto eh ' obligatory anyway
   
   .... any code

xt: Exit Sub/Function
   
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub/Function
```

To get the path to the error displayed with the error message required to additional code lines - at least in the entry procedure:

```vbs
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

### Usage of the debugging service (identifying the error line when code lines are not numbered)
It appears that there is no way for identifying the error line when the lines ar not numbered - what they usually aren't - what extends  unproductive error chasing time. The below 'trick' provides a true godsend in case:

```vbs
eh:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    mErH.ErrMsg ....
End Sub/Function
```

The _ErrMsg_ service has this 'trick' already built-in. When the _Conditional Compile Argument_ `Debugging = 1` the error message is displayed with two extra buttons:

![](../Assets/ErrMsgWithDebuggingOption.png)
![](/Assets/ErrMsgWithDebuggingOption.png)

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
I am a great fan of regression testing, regarded obligatory specifically for _Common Components_ like the _mErH_ module for instance (see the [full development and test environment][8] of example). However, any test of an error condition would interrupt a preferably automated regression test with the display of the tested error message. When a test procedure uses the _BoTP_ service instead of the _BoP_ service the ***asserted error numbers*** may be specified as follows which bypassed the display of the error message.

Example:

### Test support
With the Conditional Compile Argument `Test = 1` the _ErrMsg_ service displays the two additional buttons:
![](/Assets/ErrMsgWithTestOption.png)
![](../Assets/ErrMsgWithTestOption.png)
which may be considered when clicked as follows:


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

## Usage/services details

### Error types
The error handler distinguishes between
- [Application error](#the-application-error-service)<br>provided the error had been raised by `err.Raise mErH.AppErr(n) ...` with n = 1 to 2147221503 
- VB Runtime error
- Database error


### The _error source_
The following procedure will be copied to any in component the _mErH.ErrMsg_ service is used in order to identify any procedure:

```
Private Function ErrSrc(ByVal s As String) As String
   ErrSrc = "\<component-name>." & s
End Function
```

### The _Entry Procedure_
The procedure which the error handler has recognized as the top level procedure of a call hierarchy by means of the _BoP/EoP_ service statements is considered the _Entry Procedures_. Common entry procedures are any event procedures like _Workbook\_Open_ or click events in a UserForm.

### _err.Description_ with additional information
The _ErrMsg_ service displayes the information provided by  _err.Description_. For an _Application Error_ the error description is what is provided with the `err.Raise` statement. When the error description looks like "This is a serious error.||This error may be avoided by ...." the string concatenated with || is regarded an additional information and will be displayed in the error message as such.

### Processing reply buttons
The _ErrMsg_ service returns the 'value' of the clicked button. This value may be vbYesNo, vbIgnoreRetryCancel, etc. or the caption string of the displayed button (including any vbLf!). It may be processed in either of the following ways:
```
eh: Select Case mErH.ErrMsg(ErrSrc(PROC)
        Case ....
        Case ...
    End Select
```

### Free buttons specification
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

See also the [Common VBA Message Services][6] post for more details on how to use it and its advantages.

## Optional Execution Trace
### Service
When the optional module _mTrc_ is installed it provides an execution trace whenever the processing reaches an [_Entry Procedure_](#the-entry-procedure).

### Installation of the Execution Trace
Download and import the module  [_mTrc_][5]] 

### Using the Execution Trace
When the execution trace module _mTrc_ is used with the error handler it requires only the Conditional Compile Argument `ExecTrace = 1` to activate the trace. That's it. Any executed procedure with an<br> `mErH.BoP ErrSrc(PROC)`<br>at the beginning and an<br> `mErH.EoP ErrSrc(PROC)` <br>code lines at the end of a procedure will be included in the displayed trace result.<br>
Note: When the Common VBA Execution Trace had already been used before the mErH module had been installed all mTrc.BoP/mTrc.EoP have to be changed to mErH.BoP/mErH.EoP. Any mTrc.BoC/mTrc.EoC are ok.

### _Compact_ (default) versus _Detailed_ trace result
The default is a trace display like the following:
![](../Assets/ExecutionTrace.png)
![](/Assets/ExecutionTrace.png)<br>

However, for those who do not believe in the displayed figures a detailed view may be of interest. With `mTrc.DisplayedInfo = Detailed` (yes, standard modules may have properties but they are just not auto-sensed) the following kind of trace information is displayed:
![](../Assets/ExecutionTraceDetailed.png)
![](/Assets/ExecutionTraceDetailed.png)<br>


## Contribution, development, test, maintenance
The dedicated _Common Component Workbook_ **ErH.xlsm** is used for development, test, and maintenance. This Workbook is kept in a dedicated folder which is the local equivalent (in github terminology the clone of the public [GitHub repo Common-VBA-Errror-Handler][8]. The module **_mTest_** contains all obligatory test procedures when the code is modified, the module **_mDemo_** all procedures for the images in this post. The modules **_mErH_** and **_fMsg_** are downloaded from this source. Thus, it is wise not to make any changes without specifying a branch which is merged to the master once a code change has finished and successfully tested.

Those interested not only in using the Error Handler but also modify or even contribute in improving it may fork or clone it to their own computer which is very well supported by the [GitHub Desktop for Windows][9]. That's my environment for a continuous improvement process.

[1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/mErH.bas
[2]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/fMsg.frm
[3]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/fMsg.frx
[4]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/mMsg.bas
[5]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/mTrc.bas
[6]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
[7]:https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings
[8]:https://github.com/warbe-maker/Common-VBA-Error-Services
[9]:https://desktop.github.com
[10]:https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1