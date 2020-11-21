---
layout: post
title: Common VBA Execution Trace
subtitle: When performance has become an issue
date: 2020-11-14
categories: vba common
---

In this post
[Service](#service)<br>
[Installation](#installation)<br>
[Usage](#usage)<br>
[Options, Properties](#options-properties)<br>

### Service
When the optional module _mTrc_ is installed it provides an execution trace whenever the processing reaches an [_Entry Procedure_](#the-entry-procedure).

### Installation
- Download and import the module  [_mTrc_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Handler/master/mTrc.bas) 
- For the unique identification of a traced procedure copy the following into every module of which procedures are to be traced:
```vbscript
Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTrc." & sProc
End Function
```

- When the execution trace module _mTrc_ is used stand-alone (i.e. without the error handler see blog post []()) some more steps for the installation are required:
-- Copy the following into the module section of any standard module:
```vbscript
Public Const CONCAT = "||"
' ----------------------------------------------------------------------
' Deklarations for the use of the fMsg UserForm (Alternative VBA MsgBox)
Public Enum StartupPosition         ' ---------------------------
    Manual = 0                      ' Used to position the
    CenterOwner = 1                 ' final setup message form
    CenterScreen = 2                ' horizontally and vertically
    WindowsDefault = 3              ' centered on the screen
End Enum                            ' ---------------------------

Public Type tSection                ' ------------------
       sLabel As String             ' Structure of the
       sText As String              ' UserForm's
       bMonspaced As Boolean        ' message area which
End Type                            ' consists of
Public Type tMessage                ' three message
       section(1 To 3) As tSection  ' sections
End Type                            ' -------------------
' ----------------------------------------------------------------------
```

### Usage
Set the Conditional Compile Argument `ExecTrace = 1` and make sure any BoP/EoP, BoC/EoC statements are fully qualified with `mTrc.` That's it. Any executed procedure with an<br> `mErH.BoP ErrSrc(PROC)`<br>at the beginning and an<br> `mErH.EoP ErrSrc(PROC)` <br>code line at the end of a procedure plus any code lines with a BoC at the beginning and EoC at the end will be included in the displayed trace result. The trace result is automatically displayed whenever the execution has returned to the _Entry Procedure_.
#### Usage example
```vbscript
Private Sub Any(ByVal anyarg As String)
    Const PROC = "Any"
    On Error Goto eh
    BoP ErrSrc(PROC), anyarg ' Begin trace with value of the argument
    '               ! Use Goto xt instead of Exit Sub !
    ' .... any code ! in order not to                 !
    '               ! bypass the EoP statement        !
xt: EoP ErrSrc(PROC)
    Exit Sub
eh:
    ' any error handling
End Sub
```

### Options, Properties
#### _Compact_ (default) or _Detailed_ trace result
The default is a trace display like the following:
![](../Assets/ExecutionTrace.png)
![](/Assets/ExecutionTrace.png)<br>

However, for those who do not believe in the displayed figures a detailed view may be of interest. With `mTrc.DisplayedInfo = Detailed` (yes, standard modules may have properties but they are just not auto-sensed) the following kind of trace information is displayed:
![](../Assets/ExecutionTraceDetailed.png)
![](/Assets/ExecutionTraceDetailed.png)<br>

#### Precision
The seconds precision defaults to 6 decimals which should be far enough since results will differ anyway from trace to trace due to system conditions. The property _Precision_ may be used to change the default however.

### Contribution, development, test, maintenance
The dedicated _Common Component Workbook_ **Trc.xlsm** is used for development, test, and maintenance of the _mTrc_ module. This Workbook is kept in a dedicated folder which is [GitHub's public repo Common-VBA-Errror-Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler) clone folder. The module **_mTest_** contains all test procedures of which the execution is obligatory when the code is modified. The module **_fMsg_** is used to display the trace results without any limit in the width (a vertical scroll bar is displayed when the maximum message window width is exceeded). Any changes are preferably managed by means of a branch which is merged to the master once the change/modification has successfully passed the regression test.

Those interested not only in using the Execution Trace but also motivated to modify/improve it are kindly asked to fork it and create a pull request before start with the modification. I may invite you to contribute.