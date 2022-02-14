---
layout: post
title: Common VBA Components
date:          2021-02-19
modified_date: 2022-02-14
categories:    vba common
---
A true development performance boost provided they are well designed, continuously maintained and carefully tested.
<!--more-->

## Introduction
### Disambiguation
> A _Common Component_ has the same content in any VB-Project using it. It is developed, maintained, and tested in ***one*** specific -  preferably dedicated - Workbook/VB-Project.<br>A component/module just having the same name with different code is ***not*** a _Common Component_ in the subsequent sense.

### My _Common Components_ 
- had initially been developed when it seemed appropriate
- had been maintained and extended every now and then
- has its dedicated VB-Project which includes a test environment and an unattended Regression Test
- is kept in a public GitHub repo of which I use clones
- meets a consistent coding standard and follows clean code principals (no defaults, early binding, avoiding unintended 'case' changes, etc.)

### How to keep them up-to-date in VB-Projects using them?
I use a _[Common Component Management][1]_ Workbook which is saved as _Addin_ and provides - amongst others - the service to _Update Outdated Common Components_. A bit sophisticated but well for the  job.

## Personal and public use of (my) _Common Components_
### The challenge
I do not like the idea maintaining different code versions of _Common Components_, one which I use in my VB-Projects and another 'public' version. On the other hand I do not want to bother users of my _Common Components_ with an environment and features they may not like/use.

### Managing the splits
The primary goal is to provide _Common Components_ which are as autonomous as possible by allowing to optionally use them in a more sophisticated environment (my preference). This is achieved by using a couple of [Procedures with environment flexibility](#procedures-with-environment-flexibility) which use other _Common Components_ when installed which is indicated by the use of a couple of _[Conditional Compile Arguments](#conditional-compile-arguments)_.

#### Conditional Compile Arguments

| Conditional<br>Compile&nbsp;Argument | Purpose |
| ------------------------------------ | ------- |
| _Debugging_                          | Indicates that error messages should be displayed with a debugging option allowing to resume the error line |
| _ExecTrace_                          | Indicates that the _[mTrc][4]_ module is installed
| _MsgComp_                            | indicates that the _[mMsg][3]_, _[fMsg.frm][1]_, and _[fMsg.frx][2]_ are installed |
| _ErHComp_                            | Indicates that the _[mErH][6]_ is installed |

#### Procedures providing environment flexibility
##### The _BoP/EoP_ procedures
Copied to components/modules which have procedures prepared for optionally being traced (when _mTrc_ is installed and the _Conditional Compile Argument_ `ExecTrace = 1` the statements BoP/EoP may also be used by the mErH module when installed and activated by the _Conditional Compile Argument_ `ErHComp = 1`. See the [Common VBA Execution Trace Services][4].
```vb
Private Sub BoP(ByVal b_proc As String, _
                ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' service.
' Has no effect unless the Conditional Compile Argument 'ExecTrace = 1' (when
' the Common Execution Trace Component (mTrc) is installed. Serves for the
' Common Error Handling Component (mErH) when installed and the Conditional
' Compile Arguments 'ExecTrace = 1'.
' ------------------------------------------------------------------------------
    Dim s As String:    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")

#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case only the mTrc is installed but not the merH.
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If

End Sub

Private Sub EoP(ByVal e_proc As String, _
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

##### The "universal" _ErrMsg_ function
See inline comments below.
```vb
Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_no As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option active
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section displaying text connected to an error
' message by two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function
```

## My _Common Components_ (overview)
|Component                          |Module(s)       |Status                 |Comment               |
| --------------------------------- | -------------- | --------------------- | -------------------- |
|Common VBA Message Services        |mMsg, fMsg      |[public GitHub repo][2]|Used by mErH (optionally by mTrc |
|Common VBA Error Services          |mErH, mMsg, fMsg|[public GitHub repo][3]|Optionally uses mTrc|
|Common VBA Execution Trace Services|mTrc            |[public GitHub repo][4]|stand-alone or as optional component of mErH|
|Common VBA Excel Workbook Services |mWrkbk          |[public GitHub repo][5]|Existence/open check over multiple Excel instances, open services and other|
|Common VBA File Services           |mFile           |[public GitHub repo][6]|Existence check, etc.|
|Common VBA Basic Services          |mBasic          |private GitHub repo    | 
|Common VBA Registry Services       |mReg            |private GitHub repo    | Read/write named values simplified to the max |
 
 
 

[1]:https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services
[2]:https://github.com/warbe-maker/Common-VBA-Message-Service
[3]:https://github.com/warbe-maker/Common-VBA-Error-Services
[4]:https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
[5]:https://github.com/warbe-maker/Common-VBA-Excel-Workbook-Services
[6]:https://github.com/warbe-maker/Common-VBA-File-Services
[7]:https://github.com/warbe-maker/Common-VBA-Basic-Services
[8]:https://github.com/warbe-maker/Common-VBA-Registry-Services