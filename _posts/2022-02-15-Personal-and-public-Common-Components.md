---
layout: post
title: Personal and public use of my _Common Components_
date:          2022-02-15
modified_date: 2022-02-15
categories:    vba common
---
Managing the splits: Aiming for _Common Components_ well designed for being used in my own VB-Projects without bothering other users with my (more sophisticated) use of them. <!--more-->

## Preface
I do not like the idea maintaining different code versions of _Common Components_, one which I use in my VB-Projects and another 'public' one. On the other hand I do not want to bother users with integration issues when my _Common Components_ make use of other _Common Components_.

My primary goal is to provide _Common Components_ which function as autonomous as possible while integrating and interacting with other _Common Components_ in my VB-Projects: The [Common VBA Message Service][1] which provides the desired flexibility the VBA.MsgBox does not have and the [Common VBA Error Services][2] which displays errors with the _path to the error_.

## Managing the splits
Optionally available/installed/used _Common Components_ are indicated by a couple of _[Conditional Compile Arguments](#conditional-compile-arguments)_ which are used in [Procedures providing the environment flexibility](#procedures-providing-the-environment-flexibility) procedures also by a couple of procedures which only optionally use other _Common Components_ only when installed.

### _Conditional Compile Arguments_

| Cond.&nbsp;Comp.&nbsp;Arg. | Purpose |
| -------------------------- | ------- |
| `Debugging = 1`            | Indicates that the error messages should be displayed with a debugging option allowing to resume the error line |
| `ExecTrace = 1`            | Indicates that the  _[Common VBA Execution Trace Service][3] is installed and will actively be used |
|  `MsgComp = 1`             | Indicates that [Common VBA Message Service][1] is installed so that the _mMsg.Dsply_ service can be used as alternative to the `VBA.MsgBox` |
| `ErHComp = 1`              | Indicates that the [Common VBA Error Services][2] is installed which is able to display the 'path-to-the-error' |


### Procedures providing the environment flexibility
#### _ErrMsg_
Used in each _Common Component_ for the display of an error message. When neither the [Common VBA Message Service][1] nor the [Common VBA Error Service][2] is installed the error message is displayed by means of the `VBA.MsgBox`, else the 'better option' is used.

```vb
Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
'
' - Displays a debugging option button when the Conditional Compile Argument
'   'Debugging = 1'
' - Displays an optional additional "About the error:" section when a string is
'   concatenated with the error message by two vertical bars (||)
' - Invokes mErH.ErrMsg when the Conditional Compile Argument ErHComp = !
' - Invokes mMsg.ErrMsg when the Conditional Compile Argument MsgComp = ! (and
'   the mErH module is not installed / MsgComp not set)
' - Displays the error message by means of VBA.MsgBox when neither of the two
'   components is installed
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to
'          turn them into negative and in the error message back into a
'          positive number.
' - ErrSrc To provide an unambiguous procedure name by prefixing is with the
'          module name.
'
' See:
' https://github.com/warbe-maker/Common-VBA-Error-Services
'
' W. Rauschenberger Berlin, Feb 2022
' ------------------------------------------------------------------------------' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
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
    
    '~~ Consider extra information is provided with the error description
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
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function
```
#### _BoP/EoP_
Keeps the use of the mErH module and the mTrc module optional.
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

### Example
The _[Common VBA Error Services][1]_ and the _[Common VBA Execution Trace Services][3]_ have the following in common:
1. Both use in each component/module the `ErrSrc` function to uniquely identify a procedure's name (i.e. prefix it with the component's name) and the _AppErr_ function for Application Error numbers not conflicting with system errors.
3. Both use _BoP/EoP_ statements to indicate the <u>B</u>egin and <u>E</u>nd <u>o</u>f a <u>P</u>rocedure.<br>The execution trace uses the statements to begin/end the trace of a procedure<br>the error uses the statements to indicate an 'entry procedure' to which the error is passed on for being displayed (which allows gathering the 'path to the error'.

### Comments
Comments are welcome. I apologize for the fact that commenting requires a login to GitHub. This seems to be the only way to keep away spammers.

[1]:https://github.com/warbe-maker/Common-VBA-Message-Service
[2]:https://github.com/warbe-maker/Common-VBA-Error-Services
[3]:https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service

[4]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mTrc.bas
[5]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mMsg.bas
[6]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frm
[7]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frx
[6]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/mErH.bas
