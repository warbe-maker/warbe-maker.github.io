---
layout: post
title: Personal and public use of my _Common Components_
date:          2022-02-15
modified_date: 2023-07-13
categories:    vba common
---
Managing the balancing act: _Common Components_ designed for a best possible fit with my own VB-Projects using them but without bothering others with my way of using/integrating them. However, maintaining different code versions of a _Common Component_, one which I use in my VB-Projects and another 'public' version is not worthwhile.<!--more-->

## Preface
My _Common Component's_ aim is to function as autonomous as possible, i.e. not requiring any additional installed component. In other words, any additional component I use with/for them needs to remain optional.

## Managing the balancing act
1. Any additional component, i.e. another _Common Component_ remains optional by means of _[Conditional Compile Arguments](#conditional-compile-arguments)_ indication that a component is installed/available and should be used.
2. "Interface" procedures call optional components only when indicated installed/available.

### _Conditional Compile Arguments_

| Cond.&nbsp;Comp.&nbsp;Arg. | Purpose |
| -------------------------- | ------- |
| `Debugging = 1`            | Indicates that an error messages should be displayed with a [_debugging option_](#the-debugging-option). This error message option is already available with the "Interface" procedure _[ErrMsg](#the-errmsg-interface-procedure)_ even when no other components are installed/available. |
| `ExecTrace = 1`            | Indicates that the _[Common VBA Execution Trace Service_][3] is installed and will actively be used |
|  `MsgComp = 1`             | Indicates that [Common VBA Message Service][1] is installed so that the _mMsg.Dsply_ service can be used as alternative to the `VBA.MsgBox` |
| `ErHComp = 1`              | Indicates that the [Common VBA Error Services][2] is installed which is able to display the 'path-to-the-error' |


### Procedures providing the environment flexibility
#### The _ErrMsg_ interface procedure
A copy of the _ErrMsg_ "interface" procedure is used in each _Common Component_ for the display of an error message with the following options:
- a [debugging option](#the-debugging-option) button when the _Conditional Compile Argument_ 'Debugging = 1'
- an optional additional "About:" section when the err_dscrptn argument has an additional string concatenated by two vertical bars (||)
- the display of the error message by means of the _[Common VBA Message Service][1]_ when installed and active (_Conditional Compile Argument_ `MsgComp = 1`)
- the display of the possibly most comprehensive and well designed error message by the  [Common VBA Error Service][2] when installed (_Conditional Compile Argument_ `MsgComp = 1`)
- the display of an error message by means of the `VBA.MsgBox` when neither the [Common VBA Message Service][1] nor the [Common VBA Error Service][2] is installed.

```vb
Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_no As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. Obligatory copy Private for any
' VB-Component using the common error service but not having the mBasic common
' component installed.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, June 2023
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
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
    '~~ About
    ErrDesc = err_dscrptn
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    End If
    '~~ Type of error
    If err_no < 0 Then
        ErrType = "Application Error ": ErrNo = AppErr(err_no)
    Else
        ErrType = "VB Runtime Error ":  ErrNo = err_no
        If err_dscrptn Like "*DAO*" _
        Or err_dscrptn Like "*ODBC*" _
        Or err_dscrptn Like "*Oracle*" _
        Then ErrType = "Database Error "
    End If
    
    '~~ Title
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")
    '~~ Description
    ErrText = "Error: " & vbLf & ErrDesc
    '~~ About
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging = 1 Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function
```
#### The _BoP/EoP_ interface procedures
A copy of the below procedures in each _Common Component_ keeps the installation/availability of the _[Common VBA Error Services][2]_ and the _[Common VBA Execution Trace Service][3]_ optional. See how these procedures (the corresponding code lines respectively effects the display of the _[path to the error][6]_ displayed when the _[Common VBA Error Services][2]_ is installed.

```vb
Public Sub BoP(ByVal b_proc As String, _
      Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP b_proc, b_args
#ElseIf XcTrc_clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.BoP b_proc, b_args
#ElseIf XcTrc_mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP e_proc, e_args
#ElseIf XcTrc_clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP e_proc, e_args
#ElseIf XcTrc_mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP e_proc, e_args
#End If
End Sub
```
### The "Debugging Option"
See the _[Used Common Components_][5] section in the README of the _[Common VBA Error Services][2]_ for how the option is displayed depending on the used (or not used) _Common Components_.

## Comments
Comments are welcome. I apologize for the fact that commenting requires a login to GitHub. This seems to be the only way to keep away spammers.

[1]:https://github.com/warbe-maker/VBA-Message
[2]:https://github.com/warbe-maker/VBA-Error
[3]:https://github.com/warbe-maker/VBA-Trace
[4]:https://github.com/warbe-maker/VBA-Basics
[5]:https://github.com/warbe-maker/VBA-Error#used-common-components
[6]:https://github.com/warbe-maker/VBA-Error#the-path-to-the-error