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
[Usage examples](#usage-examples)<br>
[Installation](#installation)<br>
[Development, test, maintenance](#development-test-maintenance)

### Method
The method is as simple as displaying an error message, the result however is as complete as one can think.
The _ErrHndlr_ in module _mErrHndlr_ provides:
- a comprehensive error message
- a path to the error
- a complete execution trace

### Syntax

### Usage

### Installation
- Download and import [_mErrHndlr_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mErrHndlr.bas)
- Download and import [_clsCallStack_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/clsCallStack.cls)
- Download and import [_clsCallStackItem_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/clsCallStackItem.cls)
- In case a more appealing kind of error message is preferred (see the [difference in the display](#difference-in-display-of-the-error-message)) also download [fMsg.frm](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm) and [fMsg.frx](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsf.frx), import _fMsg.frm_ and in the module _mErrHndlr_ set the local Conditional Compile Argument:<br>`#Const AlternateMsgBox = 1`

#### Difference in display of the error message
##### Using the VB MsgBox
![](Assets/ErrorMsgMsgBox.png)
##### Using the Alternative VB MsgBox
![](Assets/ErrMsgAlternativeMsgBox.png)

### Development, test, maintenance
- The dedicated _Common Component Workbook_ ErrHndlr.xlsm is the development, test, and maintenance environment (see the Guthub repo [Common-VBA-Errror-Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler).
- The module _mTest_ contains all test procedures