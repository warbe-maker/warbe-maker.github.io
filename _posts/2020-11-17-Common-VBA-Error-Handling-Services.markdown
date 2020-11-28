---
layout: post
title: Common VBA Error Handling Services (inspired by the best of the web)
subtitle: An Error Handler inspired by the best of the web
date: 2020-11-21
categories: vba common
---

**This is not a tutorial about error handling** but the description of a full featured ready to use error handler module with an optional execution trace module.

## Services of the Error Handler
### The _ErrMsg_ service
- displays a well structured error message with
  - the **[type of the error](#the-type-of-the-error)** (Application error, VB Runtime error, and Database error) 
  - the description of the error (Err.Description,
  - the **[error source](#the-error-source)**,
  - the **[path to the error](#the-path-to-the-error)** provided the **[_Entry Procedure_](#the-entry-procedure)** is known, 
  - an optional **[additional information about an error](#additional-information-about-an-error)**,
  - (almost) any number of **[free specified buttons](#free-buttons-specification)**
  - displays the error line when available
  - displays specific buttons supporting test, provided the  Conditional Compile Argument `Test = 1` (see image below)
  - displays specific buttons supporting debug provided the Conditional Compile Argument `Debuggig = 1`
  - considers provided **asserted error numbers** to bypass the display the corresponding error message (regression test support)
- waits for the user's button clicked and provides/returns [the reply button's value](#processing-the-returned-reply) to the caller.

![image](../Assets/ErrMsgWithDebuggingOption.png)

The _ErrMsg_ service has the following syntax
```vbs
   mErH.ErrMsg error-number, error-source, error-description, error-line[, buttons]
```
The _ErrMsg_ service has these named arguments:

|  Argument   |   Description   |
| ----------- | --------------- |
| err_number  | Optional, defaults to err.Number when omitted      |
| err_source  | Obligatory, string expression providing \<module>.\<procedure>, see [ErrSrc(PROC)](#the-error-source).   |
| err_dscrptn | Optional, defaults to err.Description when omitted |
| err_line    | Optional, defaults to  Erl when omitted            |
| err_buttons | Optional. Variant. Defaults to "Terminate execution" button when omitted.<br>May be a value for the VBA MsgBoc [_Buttons_](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) argument and/or any descriptive button caption string (including line breaks for a multi-line caption. The buttons may be provided as a comma delimited string, a collection or a dictionary. vbLf items display the following buttons in a new row. |

### The Begin/End of Procedure services _BoP_, _EoP_
Indicate the begin/end of a procedure. The information is used by the error handler to identify the [_Entry Procedure_](#the-entry-procedure) and to compose the path to the error.

The _BoP_ services has the following syntax: `BoP procedure-id[, arguments]`, the _EoP_ service has the syntax `EoP procedure-id`. The _BoP_\/_EoP_ services have the following named arguments:

| Service |   Argument   |   Description   |
| ------- | ------------ | --------------- |
| BoP     | bop_id       | Obligatory, Expression providing a unique name of the procedure, e.g.<br>[ErrSrc(PROC)](#the-error-source) |
| BoP     | bop_arguments| Optional, ParamArray, a list of the procedures argument, optionally paired as name, value |
| EoP     | eop_id       | Obligatory, Expression providing a unique name of the procedure, e.g.<br>[ErrSrc(PROC)](#the-error-source) |

Example:
```
Private Sub Any(ByVal arg1 As String, ByVal arg2 as Single)
    On Error Goto eh
    Const PROC = "Any"
    BoP ErrSrc(PROC), "arg1=", arg1, "arg2=", arg2
    .... ' any code

xt: EoP ErrSrc(PROC) ' never use Exit Sub but Goto xt instead
    Exit Sub         ' in order not to bypass the EoP service

eh: ErrMsg ....
End Sub
```

### The Begin of Test-procedure service _BoTP_
Indicates the begin of a test procedure.
The _BoTP_ services has the following syntax: `BoTP procedure-id, asserted-errors`. The _BoTP_ service has the following named arguments:

|      Argument     |   Description   |
| ----------------- | --------------- |
| botp_id           | Obligatory, Expression providing a unique name of the procedure, e.g.<br>[ErrSrc(PROC)](#the-error-source) |
| botp_errs_asserted| Obligatory, ParamArray, type Long, list of error numbers regarded asserted by the _ErrMsg_ service for the executed procedure including any sub-procedures.<br>Note 1: In case there are no asserted errors because the test-procdure does not test nay error conditions the _BoP_ service will be used instead.<br>Note 2: The _ErrMsg_ service will bypass the display of the error message for any of the corresponding error numbers. By this the test procedure will run without any user interaction as long as no other error but the asserted ones occours. |


## Installation
- Download and import the module  [_mErH_][1]
- Download the UserForm [fMsg.frm][2] and [fMsg.frx][3] and import _fMsg.frm_
- Download and import [mMsg.bas][4]
- Because at least some effort is the same intalling the Common VBA Execution Trace Service is worth being concidered: Download [mTrc.bas][5] and import it.

## Basic usage
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


### Usage providing a "path to the error" with the error message
When the user has no choice because just the default button is displayed with the error message an error is passed on to [the _Entry Procedure](#the-entry-procedure) and thereby the path to the error is assembled.
 

### Debug supporting usage
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

### Usage supporting test
For _Common Components_ like this error handler I regard a regression test obligatory before a code modification is published. However, any test of an error condition stops the test process when there is only the default button displayed.

With the Conditional Compile Argument `Test = 1` two additional buttons are displayed: **Continue with next code line** and **Continue with next procedure**

image still missing

and the following can be for a test continuation

```vbs

```
## Usage/services details
### The type of the error
The error handler distinguishes between
- [Application error](#the-application-error-service)<br>provided the error had been raised by `err.Raise mErH.AppErr(n) ...` with n = 1 to 2147221503 
- VB Runtime error
- Database error


### The _error source_
Since the err.Source only provides the Application name we have to care ourselves for this information:<br>
Copy the following in any module the error handler (mErH.ErrMsg) is used
```
Private Function ErrSrc(ByVal s As String) As String
   ErrSrc = "module-name." & s
End Function
```
### The _Application Error_ service
The error Handler provides the function _AppErr_ which turns a positive number into a negative by adding the constant [_vbObjectError_](<https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1>) when the error is raised with `err.Raise mErH.AppErr(n)`. The error handler (the _ErrMsg_ function) recognizes the negative number as an _Application Error_ and turns it back into the original postive number for display.

### The _Entry Procedure_
The procedure which the error handler has recognized as the top level procedure of a call hierarchy by means of a pair of BoP/EoP statements is considered the _Entry Procedures_.

```
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
Private Function ErrSrc(ByVal s As String) As String
   ErrSrc = "module-name." & s
End Function
```

### The _Path to the error_
For the display of the path to the error at least one procedure must have been regognized as an/the  [_Entry Procedure_](#the-entry-procedure).<br>
When the user has no reply choices since only one button is displayed with the error message, the path to the error is composed when the error passed on to the _Entry Procedure_ where the error is displayed when reached. This is the reason why in this particular case there is no need to have BoP/EoP statements in every procedure.

When the user has choices because more than one button is displayed with the error message the error is displayed immediately with the procedure which caused the error. In this case there is only one source for the path to the error which is the stack maintained by the error handler with each BoP/EoP statement. I.e. the path to the error depends on procedures which provide a BoP/EoP information.

### Additional information about an error
The displayed error description is what is provided by the err.Description property. However, in case of an _Application Error_ the description is provided with the err.Raise command. When the error description looks like "This is a serious error.||This error may be avoided by ...." the string concatenated with || is regarded an additional information and will be displayed in the error message as such.

### The test option buttons
With the Conditional Compile Argument `Test = 1` the error message will be displayed with two additional buttons which may be used when the [return value of the _ErrMsg_ is processed](#processing-the-return-value) further.
![](/Assets/ErrMsgWithTestOption.png)
![](../Assets/ErrMsgWithTestOption.png)

### Processing the return value
The return value is the value of the button when provided by vbYesNo, vbIgnoreRetryCancel, etc. or the caption string of the displayed button (including any vbLf). It may be processed in either of the following ways:
```
eh: Select Case mErH.ErrMsg(ErrSrc(PROC)
        Case ....
        Case ...
    End Select
```

or
```
eh: mErH.ErrMsg ErrSrc(PROC)
    Select Case mErH.ErrReply
        Case ....
        Case ...
    End Select
```
or
```
    Dim vReply As Variant
    
eh: mErH.ErrMsg ErrSrc(PROC), err_reply:=vReply
    Select Case vReply
        Case ....
        Case ...
    End Select
```

### The debugging option buttons
When the Conditional Compile Argument `Debugging = 1` the error message looks as follows:
![](/Assets/ErrMsgWithDebuggingOption.png)
![](../Assets/ErrMsgWithDebuggingOption.png)
The additional button have an advantage over the equivalent:
```
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
```
because this cannot be altered which means it loops until the reason for the error has been eliminated which may result in an unwanted code change just to continue without an error. The two buttons may be processed as return values as follows:
```
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt1ResumeNext: Resume Next
    End Select
```

With both Conditional Compile Arguments `Test = 1` and `Debugging = 1` four additional buttons are displayed and may be processed accordingly.

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

See also the [Alternative VBA MsgBox](https://github.com/warbe-maker/VBA-MsgBox-Alternative) for more details on how to use it and its advantages (not yet available as post).

## Optional Execution Trace
### Service
When the optional module _mTrc_ is installed it provides an execution trace whenever the processing reaches an [_Entry Procedure_](#the-entry-procedure).

### Installation of the Execution Trace
Download and import the module  [_mTrc_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mTrc.bas) 

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
The dedicated _Common Component Workbook_ **ErH.xlsm** is used for development, test, and maintenance. This Workbook is kept in a dedicated folder which is the local equivalent (in github terminology the clone of the public [GitHub repo Common-VBA-Errror-Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler). The module **_mTest_** contains all obligatory test procedures when the code is modified, the module **_mDemo_** all procedures for the images in this post. The modules **_mErH_** and **_fMsg_** are downloaded from this source. Thus, it is wise not to make any changes without specifying a branch which is merged to the master once a code change has finished and successfully tested.

Those interested not only in using the Error Handler but also modify or even contribute in improving it may fork or clone it to their own computer which is very well supported by the [GitHub Desktop for Windows](https://desktop.github.com). That's my environment for a continuous improvement process.

[1]: (https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mErH.bas)
[2]: (https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm)
[3]: (https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frx)
[4]: (https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mMsg.bas)
[5]: (https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mTrc.bas)