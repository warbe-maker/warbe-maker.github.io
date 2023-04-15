---
layout:        post
title:         All the matter for a VB error message
date:          2020-11-21
modified_date: 2023-04-15
---
For professionally and semi-professionally developed VB-Projects this post considers the absolute minimum about a proper, i.e. debug supporting error handling/message.
<!--more-->

## The Error Number
The _Number_ property of the _Err_ object may indicate a VB Runtime, a Database, or an Application Error. The latter is one explicitly raised by `Err.Raise`. Microsoft documentation says, the error number raised by means of `Err.Raise` should be the sum of the application error n +  [_vbObjectError_](<https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1>) in order to avoid conflicts with  is a _VB Runtime Error_. I call such an error an _Application Error_ of which the number is set by:
```vb
Public Function AppErr(ByVal a_err_no As Long) As Long
' ------------------------------------------------------------------------------
' To ensure a programmed 'Application' error number (raised by Err.Raise) not
' conflicts with a 'VB Runtime Error' or any other system error the function
' returns a given positive number (a_err_no) into a negative one by adding the
' system constant 'vbObjectError' and returns the original 'Application Error'
' number when called with a negative error number.
' ------------------------------------------------------------------------------
    If a_err_no >= 0 Then AppErr = a_err_no + vbObjectError Else AppErr = Abs(a_err_no - vbObjectError)
End Function
```

The error handling may investigate the number as follows:
```vb
    Dim ErrNo   As long
    Dim ErrType As String
    
    If Err.Number < 0 Then
        ErrNo = AppErr(err_no) ' converts the error number set with Err.Raise AppErr(n) back to its origin number
        ErrType = "Application Error "
    Else
        ErrNo = err_no
        ErrType = ' requires further investigation (see below)
    End If
```

## The Error Type
An error message should preferably distinguish between _VB Runtime error_, _Application error_, and _Database error_. This distinction requires the analysis of the `Err.Number` and the `Err.Description` and the `AppErr` service.
```vb
    Dim ErrNo   As Long
    Dim ErrType As String
    
    Select Case Err.Number
        Case Is < 0
            ErrNo = AppErr(err_no) ' converts the error number set with Err.Raise AppErr(n) back to its origin number
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
```

## The Error Source
The source of the error is essential! Unfortunately VBA's `Err.Source` only returns the application name. The only way to provide a more meaningful information is the following in each module:
```vb
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "<module-name>." & s
End Function
```
## The Error Line
Unfortunately, in case of an error, VBA's _Erl_ variable only returns the code line which caused the error when there is one - which usually isn't the case. The following is a godsend when needed to locate the line where the error occurred:
```vb
    ...
    On error Goto eh
    ...
xt: Exit Function/Sub

eh:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrMsg ....
End Function/Sub
```
When the _Conditional Compile Argument_ `Debugging = 1` the ***F5*** key will go straight to the line the error occurred - with the disadvantage that the error will loop until solved by a code modification. This may be avoided by displaying a "Resume" and a "Terminate" button (see [Conclusion](#conclusion) below.

## Summary
The below procedure addresses all the mentioned above:

```vb
Public Function ErrMsg(ByVal err_src As String, _
              Optional ByVal err_dsc As String = vbNullString) As Variant
' ------------------------------------------------------------------------------
' Universal error message including a debugging option button (when Conditional
' Compile Argument 'Debugging = 1') and an optional additional "About:" section
' when an error description argument (err_dsc) is provided with an additional
' string concatenated by two vertical bars (||).
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into negative and in the error message back into
'               its origin positive number.
'       ErrSrc  To provide an unambiguous procedure name prefixed with the
'               module's name.
' ------------------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_src = vbNullString Then err_src = Err.source
    If err_dsc = vbNullString Then err_dsc = Err.Description
    If err_dsc = vbNullString Then err_dsc = "--- No error description available ---"
    
    '~~ Consider extra information is provided with the error description
    If InStr(err_dsc, "||") <> 0 Then
        ErrDesc = Split(err_dsc, "||")(0)
        ErrAbout = Split(err_dsc, "||")(1)
    Else
        ErrDesc = err_dsc
    End If
    
    '~~ Determine the type of error
    Select Case Err.Number
        Case Is < 0
            ErrNo = AppErr(Err.Number)
            ErrType = "Application Error "
        Case Else
            ErrNo = Err.Number
            If err_dsc Like "*DAO*" _
            Or err_dsc Like "*ODBC*" _
            Or err_dsc Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_src <> vbNullString Then ErrSrc = " in: """ & err_src & """" ' assemble ErrSrc from available information"
    If Erl <> 0 Then ErrAtLine = " at line " & Erl                      ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ") ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_src & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)

End Function
```

The below test example
```vb
Public Sub Test_ErrMsg()
    Const PROC = "Test_ErrMsg"
    
    On Error GoTo eh
    Err.Raise AppErr(10), ErrSrc(PROC), "This is an application error" & "||" & "This is an optional additional info about the error."
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
```
displays an error message where a _Yes_ reply and the _F5_ key goes straight to the error line - provided the _Conditional Compile Argument_ `Debugging = 1`.

## Perspective
I personally prefer the [Common VBA Error Services][1] in all my VBProjects. 

[1]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/error/handling/2021/01/16/Common-VBA-Error-Services.html