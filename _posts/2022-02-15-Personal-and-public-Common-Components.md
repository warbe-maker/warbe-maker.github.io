---
layout: post
title: Personal and public use of my _Common Components_
date:          2022-02-15
modified_date: 2022-02-15
categories:    vba common
---
Managing the splits: Aiming for _Common Components_ well designed for being used in my own VB-Projects without bothering other users with my (more sophisticated) use of them. <!--more-->

## Preface
I do not like the idea maintaining different code versions of _Common Components_, one which I use in my VB-Projects and another 'public' one. On the other hand I do not want to bother users of my _Common Components_ with other _Common Components_ regularly use.

### Managing the splits
My primary goal is to provide _Common Components_ which function as autonomous as possible - and also to optionally use them together with the/my [Common VBA Message Services][1] and the [Common VBA Error Services][2]. This 'optionally installed' is primarily achieved by the use of a couple of _Conditional Compile Arguments_ and procedures also by a couple of procedures which only optionally use other _Common Components_ only when installed.

#### _Conditional Compile Arguments_ with _Common Components_

| Conditional Compile Argument | Purpose |
| ---------------------------- | ------- |
| _Debugging_                  | Indicates that error messages should be displayed with a debugging option allowing to resume the error line |
| _ExecTrace_                  | Indicates that the _[mTrc][4]_ module is installed
| _MsgComp_                    | indicates that the _[mMsg][5]_, _[fMsg.frm][6]_, and _[fMsg.frx][7]_ are installed |
| _ErHComp_                    | Indicates that the _[mErH][8]_ is installed |

#### Procedures providing an optional environment
##### _ErrMsg_
Used in each _Common Component_, provides the display of an error message using the VBA.MsgBox when no other components are installed but still provides a very useful debugging option allowing to resume the error line. When the _[Common VBA Message Services][1]_ (components _[mMsg][5]_, _[fMsg.frm][6]_, and _[fMsg.frx][7]_) are installed the Error Message will look much more convenient, when additionall the [Common VBA Error Services][2] (component _[mErH][8]_) is installed the Error Message will provide a 'path to the error')
```vb
Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option
' (Conditional Compile Argument 'Debugging = 1') and an optional additional
' "about the error" information which may be connected to an error message by
' two vertical bars (||).
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
