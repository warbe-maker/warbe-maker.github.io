---
layout: post
title: Common VBA Error Hamdler
subtitle: An Error Handler assembled from the best which can be found in foruns
date: 2020-10-02 16:00 +0200
categories: vba common
---


In this post<br>
[Methods](#methods)<br>
[Syntax](#syntax)<br>
[Installation](#installation)<br>
[Usage](#usage)<br>
[Development, test, maintenance](#development-test-maintenance)


### Methods
The method is very similar to the  display of an error message:

`ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl, vbOkOnly
`

but provides a possibly surprising result:
- an error message which discriminates _VB Runtime Errors_, _Application Error_, and _Database-Error_
- a clear indication of the error source in the form <module>.<procedure>
- a path to the error or better a path from the error source to the _Entry Procedure_
- optionally an execution time trace
- an optional writing of an error log file is still missing.

 with just 4 code lines in a procedure (see [Basic usage](#basic-usage):


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

### Installation
- Download and import [_mErrHndlr_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mErrHndlr.bas)
- Download and import [_clsCallStack_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/clsCallStack.cls)
- Download and import [_clsCallStackItem_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/clsCallStackItem.cls)
- With the alternative VBA MsgBox is used (download [fMsg.frm](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm) and [fMsg.frx](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsf.frx), import _fMsg.frm_) the options become very interesting (see [Error handler, the alternative MsgBox and debugging](#error-handler-alternative-msgbox-and debugging)

and in the module _mErrHndlr_ set the local Conditional Compile Argument:<br>`#Const AlternateMsgBox = 1`
### Usage
#### Basic usage
 ```vbscript
   Const PROC = "the name of the procedure" ' for the identification of the error source
   On Error Go-to on_error ' obligatory anyway
   BoP ErrSrc(PROC) ' indicates the begin of an "Entry Procedure"
   
   .... any code

exit_proc:
   EoP ErrSrc(PROC) ' indicates the end of an "Entry Procedure"
   
   Exit Sub/Function
on_error:
   ErrHndlr .... ' see syntax and examples
End Sub/Function
```

### Usage
#### Basic
#### Elaborated
#### Debugging
When the Conditional Compile Argument Conditional CompileArgument:<br>`#Const AlternateMsgBox = 1` in module mErrHndlr is set the fMsg UserForm is required (see [Installation](#installation) which provides an outperforming means for debugging. When familiar with the "trick"
```vbs
on_error:
#If Debugging Then
    Debug.Print Err Description: Stop: Resume
#End If
```
the combination mErrHndlr with fMsg offers something even more convenient.

#### Error handler, alternative MsgBox,  and debugging

#### Difference in display of the error message
##### Using the VB MsgBox
![](Assets/ErrorMsgMsgBox.png)
##### Using the Alternative VB MsgBox
![](Assets/ErrMsgAlternativeMsgBox.png)

### Development, test, maintenance
- The dedicated _Common Component Workbook_ ErrHndlr.xlsm is the development, test, and maintenance environment (see the Guthub repo [Common-VBA-Errror-Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler).
- The module _mTest_ contains all test procedures