---
layout: post
title: Common VBA Error Services
date:          2021-01-16
categories: vba common error handling
modified_date: 2021-04-29
---
Error services inspired by the best of the web. Not an error handling tutorial but a complete description of the services and how to use them.
<!--more-->

## Introduction
The error message service described in this post might appear overdone and/or too complicated at the first glance. Having used it for a long time thereby continuously  improving it the result has become worth being established in each VBA component as a standard. I can promise that specifically error debugging will become as easy and fast as possible last but not least due to the built-in [_Debugging&nbsp;Option_](#the-debugging-option). In the future it will take just seconds to detect the error causing code line.

## Disambiguation
| Term            | Meaning                                         |
| --------------- | ----------------------------------------------- |
|_Application&nbsp;Error_| An error which had been raised by an `err.Raise` statement. An _Application Error_ is distinguished from any VB Run-time or Database error by adding the [_vbObjectError_][10] to an integer number. See [The _AppErr_ service](#the-apperr-service).  |
|_Entry&nbsp;Procedure_| The key for the display of a 'path to the error'. The top level procedure in a call hierarchy with a _BoP/EoP_ statement. Typical _Entry Procedures_ are event procedures like _Workbook\_Open_ , click events in a UserForm etc..|
|_Error&nbsp;Source_   | The error message can only display the source of the error when it had been provided as argument |
|_Common&nbsp;Components, Component| Term of the VB-Project object model for a UserForm, Data Module, Class Module, or Standard Module). I use the term for any component regarded common in the sense that it is used by more than one VB-Project.See also my [Common VBA Components][18] |

## The _ErrMsg_ service (by the mErH component)
 displays a well structured error message with
  - the **[type of the error](#error-types)** by distinguishing
    - Application Error (see [The _AppErr_ service](#the-apperr-service))
    - VB Run-time error
    - Database error 
  - the description of the error (_err.Description_),
  - the source of the error (see [The _ErrSrc_ service](#the-errsrc-service)),
  - the path to the error (see [The BoP/EoP service for the path-to-the-error](#the-bopeop-service-for-the-path-to-the-error)), 
  - an optional **[additional information about an error](#error-description)** Services
  - (almost) any number of **[free specified buttons](#free-buttons-specification)**
  - the error line when available
- waits for the user's button clicked and provides/returns [the reply button's value](#processing-reply-buttons) to the caller.

The _ErrMsg_ service has the following syntax (error description and error line are obtained from the err object)
```VB
    On Error Goto eh
    ' .....
eh: mErH.ErrMsg error-source
```
The _ErrMsg_ service has these named arguments:

|  Argument   |   Description   |
| ----------- | --------------- |
| err_source  | Obligatory, string expression providing \<module>.\<procedure>, see [ErrSrc(PROC)](#the-error-source).   |
| err_buttons | Optional. Variant. Defaults to "Terminate execution" button when omitted.<br>May be a value for the VBA MsgBox [_Buttons_][7] argument and/or any descriptive button caption string (including line breaks for a multi-line caption. The buttons may be provided as a comma delimited string, a collection or a dictionary. vbLf items display the following buttons in a new row. |

### The _BoTP_ service (by the mErH component)
A regression test - preferably automated, un-attended and self-asserted - may include tests of specific error conditions. The _BoTP_ service is the means to specify these anticipated error numbers. Any error number indicated 'asserted' the _ErrMsg_ service by-passess the display of the error message.

The BoTP service has the following Syntax:<br>
`BoTP procedure-id, err-number[, err-number] ...`
with the following named arguments:

|      Argument     |   Description   |
| ----------------- | --------------- |
| botp_id           | Obligatory, Expression providing a unique name of the procedure, e.g.<br>[ErrSrc(PROC)](#the-error-source) |
| botp_errs_asserted| Obligatory, ParamArray with positive numbers |

## Basic services (AppErr, BoP, EoP, ErrMsg)
The below basic services may either be copied into each component or used from the [mBasic][17] component when installed.

### The _AppErr_ service
The _AppErr_ is provided by a code snippet copied into each component which uses the _ErrMsg_ service. In order to not confuse errors raised with `err.Raise ...` the service adds the [_vbObjectError_][10] constant to a given positive number to turn it into a negative. Because an error is always displayed with the _Error&nbsp;Source_ each procedure can have it's own positive error numbers ranging from 1 to n. The _ErrMsg_ service, when detecting a negative error number uses the _AppErr_ service to turn it back into it's original positive error number.
#### Using the _AppErr_ service
Code snippet to be copied (alternatively to the installation of the [mBasic][17] component):
```VB
Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB run-time error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function
```
and called it with:
```VB
err.Raise AppErr(n), ErrSrc(PROC), "error description"
```
### The _BoP/EoP_ service for the path-to-the-error
The service is provided by the following two procedures either copied into each module or used with the mBasic component when installed.
```VB
Private Sub BoP(ByVal b_proc As String, _
          ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' service. When neither the Common Execution Trace
' Component (mTrc) nor the Common Error Handling Component (mErH) is installed
' (indicated by the Conditional Compile Arguments 'ExecTrace = 1' and/or the
' Conditional Compile Argument 'ErHComp = 1') this procedure does nothing.
' Else the service is handed over to the corresponding procedures.
' May be copied as Private Sub into any module or directly used when mBasic is
' installed.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case only the mTrc is installed but not the merH.
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Public Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' service. When neither the Common Execution Trace
' Component (mTrc) nor the Common Error Handling Component (mErH) is installed
' (indicated by the Conditional Compile Arguments 'ExecTrace = 1' and/or the
' Conditional Compile Argument 'ErHComp = 1') this procedure does nothing.
' Else the service is handed over to the corresponding procedures.
' May be copied as Private Sub into any module or directly used when mBasic is
' installed.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case the mTrc is installed but the merH is not.
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub
```
The _ErrMsg_ service of the _mErH_ module has to approaches for the display of a 'path to the error' if possible.
- When the error message is displayed immediately within the procedure the error had been raised the displayed 'path to the error' is the content of a call stack maintained by _BoP_/_EoP_ statements in procedures. By nature the 'path to the error' will be as complete as these statements in called procedures are. See [The _Debugging_ service](#the-debugging-option) for a situation when there is an extra button available to choose.
- When the user has no choice to choose a certain reply because only one button is displayed with the error message the _ErrMsg_ service passes on the error up to the 'Entry Procedure' thereby assembling the 'path to the error' on the way up. This pass-on method however is only used when the 'Entry Procedure is known which is only the case when at least the 'Entry Procedure' has _BoP/EoP_ statements.

The _BoP/EoP_ services have the following syntax:<br>
`BoP error_source[, arguments]`<br>
`EoP error_source`<br>
with the following named arguments:

| Service |   Argument   |   Description   |
| ------- | ------------ | --------------- |
| BoP     | bop_id       | Obligatory, String expression, unique identification of the procedure in the module (see [The _ErrSrc_ service](#the-errsrc-service)) |
| BoP     | bop_arguments| Optional, ParamArray, a list of the procedures argument, optionally paired as name, value |
| EoP     | eop_id       | Obligatory, String expression, unique identification of the procedure name in the module (see [The _ErrSrc_ service](#the-errsrc-service)) |

### The _Debugging_ option
With the _Conditional Compile Argument_ `Debugging = 1` the error message is displayed with 1 additional button which allows:
```VB
Private Sub My()
    Const PROC = "My"
    
    On error Goto eh
    ' ...
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      Goto xt
    End Select
End Sub ' Function, Property
```
As a matter of fact this option is also available with the _ErrMsg_ service provided with the [mMsg][1] component. When neither the mErH nor the mMsg/fMsg components are installed the service displays the error message by means of the VBA.MsgBox. When installed the service is handed over to the mErH component or when this one is not installed to the mMsg/fMsg component.

## Installation
- Download and import the [_mErH.bas_][1] module
- Download the UserForm [fMsg.frm][2] and [fMsg.frx][3] and import _fMsg.frm_
- Download and import the [mMsg.bas][4] module
- Recommendable also is the download in import of the [_mBasic.bas_][17] module. It will save copying code snippets to each module.
- By the way installing the _Common VBA Execution Trace Service_ is worth being considered. Very little effort with a great service when it comes to performance measures:<br> Download and import the[mTrc.bas][5] module.<br>
When the _mTrc_ module is installed and the _Conditional Compile Argument_ 'ExecTrace = 1' an execution trace is displayed whenever the processing reaches an [_Entry Procedure_](#the-entry-procedure). The trace includes all procedures executed which do have a [BoP/EoP statement](#the-bopeop-service-for-the-path-to-the-error).

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


### Regression test support
Convinced of the enormous benefit of a proper regression test on the long run this kind of test has become obligatory for all my _Common Components_ potentially used in any number of VB-Projects (like this _mErH_ module for instance). However, any test of an error condition would interrupt an - otherwise perfectly automated - regression test with the display of the tested error message. With the _BoTP_ service instead of the _BoP_ service ***asserted error numbers*** may be specified in order to have the display of the error message bypassed. Clever used, this bypassing can be used for regression only, which displays the error message when the test procedure is executed individually. Below is an example of an execution trace which documents the performed tests:<br>
![](/Assets/ExecTraceRegressionTest.png)
![](../Assets/ExecTraceRegressionTest.png)
Note: The fully automated regression test can be found in the the corresponding Github repo ["Common VBA Error Services"][8] or by just downloading the [mTest.bas][15] code module of in [ErH.xlsm][16] Workbook.

### Execution trace support
When the [_mTrc_][5] module is downloaded and imported and the _Conditional Compile Argument_ `ExecTrace = 1` any arguments provided with the _BoP_ service appear in the execution trace displayed when the _Entry&nbsp;Procedure_ is reached.
```
Private Sub Any(ByVal arg1 As String, ByVal arg2 as Single)
    Const PROC = "Any"
    
    On Error Goto eh
    BoP ErrSrc(PROC), "arg1=", arg1, "arg2=", arg2

    ' .... ' any code

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: ErrMsg ErrSrc(PROC)
End Sub
```

## Other services
### The _ErrSrc_ service
The following procedure will be copied into any component which uses an _ErrMsg_ service in order to identify any procedure by the returned \<component-name>.\<procedure-name>.
```VB
Private Function ErrSrc(ByVal s As String) As String
   ErrSrc = "<component-name>." & s
End Function
```
### Error description with additional information
An _Application Error_ provides the error description with the `err.Raise` statement. When this description has text concatenated by a || this part of the description is displayed as an extra "About the error:" section.

## Contribution, development, test, maintenance
The dedicated _Common Component Workbook_ [ErH.xlsm][16] is used for development, test, and maintenance. This Workbook is kept in a dedicated folder which is the local equivalent (in github terminology the clone) of the public GitHub repo [Common-VBA-Error-Services][8]. The module [_mTest_][15] contains all test procedures obligatory being executed when the code is modified. Code modifying contributions will be handled by means of a branch merged to the master once a code change has successfully passed regression testing.

A code modifying contribution is very well supported by the [GitHub Desktop for Windows][9] which is the environment I use.

[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Error-Services/master/source/mErH.bas
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Message-Service/master/source/fMsg.frm
[3]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Message-Service/master/source/fMsg.frx
[4]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Message-Service/master/source/mMsg.bas
[5]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Message-Service/master/source/mTrc.bas
[15]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Error-Services/master/source/mTest.bas
[16]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Error-Services/master/source/ErH.xlsm
[17]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Error-Services/master/source/mBasic.bas

[6]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
[7]:https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings
[8]:https://github.com/warbe-maker/Common-VBA-Error-Services
[9]:https://desktop.github.com
[10]:https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1
[11]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/14/Common-VBA-Execution-Trace-Service.html
[12]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
[14]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html#the-buttons-service
[18]:https://warbe-maker.github.io/vba/common/2021/02/19/Common-VBA-Components.html
