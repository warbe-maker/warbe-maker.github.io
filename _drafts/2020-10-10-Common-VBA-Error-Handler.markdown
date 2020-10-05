---
layout: post
title: Common VBA Error Hamdler
subtitle: An Error Handler assembled from the best which can be found in foruns
date: 2020-10-02 16:00 +0200
categories: vba common
---
In this post<br>
[Method](#method)<br>
[Syntax](#syntax)<br>
[Settings](#settings)<br>
[Installation](#installation)<br>
[Usage examples](#usage-examples)<br>
&nbsp&nbsp[Error handler, alternative MsgBox and debugging](#error-handler-alternative-msgbox-and-debuggung)<br>
[Development, test, maintenance](#development-test-maintenance)

### Method
The method is as simple as displaying an error message:
```vbscript
ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl, vbOkOnly
```
but provides a possibly surprising result:
- an error message which discriminates _VB Runtime Errors, _Application Error_, and _Database-Error_
- a clear indication of the error source in the form <module>.<procedure>
- a path to the error or better a path from the error source to the _Entry Procedure_
- ptionally an execution time trace
- an optional writing of an error log file is still missing.

 with just 4 code lines in a procedure (see [Basic usage](#basic-usage):


### Syntax
`ErrHndlr errornumber, errorsource, errordescription, errorline[, buttons]

### Settings
The _ErrHndlr_ has the 

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
### Installation
- Download and import [_mErrHndlr_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mErrHndlr.bas)
- Download and import [_clsCallStack_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/clsCallStack.cls)
- Download and import [_clsCallStackItem_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/clsCallStackItem.cls)
- With the alternative VBA MsgBox is used (download [fMsg.frm](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm) and [fMsg.frx](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsf.frx), import _fMsg.frm_) the options become very interesting (see [Error handler, the alternative MsgBox and debugging](#error-handler-alternative-msgbox-and debugging)

and in the module _mErrHndlr_ set the local Conditional Compile Argument:<br>`#Const AlternateMsgBox = 1`
#### Complete usage

#### Error handler, alternative MsgBox,  and debugging

#### Difference in display of the error message
##### Using the VB MsgBox
![](Assets/ErrorMsgMsgBox.png)
##### Using the Alternative VB MsgBox
![](Assets/ErrMsgAlternativeMsgBox.png)

### Development, test, maintenance
- The dedicated _Common Component Workbook_ ErrHndlr.xlsm is the development, test, and maintenance environment (see the Guthub repo [Common-VBA-Errror-Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler).
- The module _mTest_ contains all test procedures