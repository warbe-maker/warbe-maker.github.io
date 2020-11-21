---
layout: post
title: All the matter for a VB error message
subtitle: Error number, Error type, Error source, etc.
date: 2020-11-21
---

<small>All aspects of this post are part of the [Common VBA Error Handler][1]</small>

## The error number
The _Number_ property of the _Err_ object may indicate a VB Runtime, a Database, or an Application Error. The latter is one explicitly raised by `Err.Raise`. Microsoft documentation says, the error number raised by means of `Err.Raise` should be the sum of the application error n +  [_vbObjectError_](<https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1>) in order to avoid conflicts with  is a _VB Runtime Error_. I call such an error an _Application Error_ of which the number is set by:
```
Public Function AppErr(ByVal err_no As Long) As Long
' --------------------------------------------------
' Used with Err.Raise AppErr(n), ...  to translate 
' the number into a negative 'Application Error
' Number' and in reverse, when the error number is
' negative translate it back into the original
' positive 'Application Error Number'.
' -------------------------------------------------
    AppErr = IIf(errno < 0, err_no - vbObjectError, vbObjectError + err_no)
End Function
```

The error handling may investigate the number as follows:
```
   Select Case Err.Number
       Case AppErr(n) ' an error erased by Err.Raise
       Case n         ' a Database or VB runtime error
   End Select
```

## The source of the error
The source of the error is the most important information in a displayed error message. Unfortunately the _Source_ property of the _err_ object does not deliver what it's name promises but just the application name. Thus this information needs to be provided in each module via:
```
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "module-name." & s
End Function
```
## The error line
The more code lines in a procedure the more desired in case of an error. Unfortunately the _Erl_ provided by VBA only delivers the code line which caused the error when there are line numbers. In by far the most cases when they are desired they are missed. The following however does the trick:
```
    On error Goto eh
    ...
    
eh:
#If Debugging Then ' Debugging is the Conditional Compile Argument Debugging = 1
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrMsg ....
End Function/Sub
```

Where I have found this the guy called it a godsend when needed. At that's what it is. The only disadvantage I found, it will loop until the error is eliminated or bypassed by any kind of code modification. Without a code modification the above may be achieved when the error message displayed comes with tow extra buttons: One called "Resume" and the other one called "Resume Next". This service is provided by my [Common VBA Error Handler][1].

## The type of error
An error message should preferably distinguish between _VB Runtime error_, _Application error_, and _Database error_. This distinction requires the analysis of the Err.Number and the Err.Description.

### All matter for an error message
The below procedure delivers/returns all the above mentioned in a way it can be used to build a proper error message:
```
Private Sub ErrMsgMatter(ByVal err_source As String, _
                         ByVal err_no As Long, _
                         ByVal err_line As Long, _
                         ByVal err_dscrptn As String, _
                Optional ByRef msg_title As String, _
                Optional ByRef msg_type As String, _
                Optional ByRef msg_line As String, _
                Optional ByRef msg_no As Long, _
                Optional ByRef msg_details As String, _
                Optional ByRef msg_dscrptn As String, _
                Optional ByRef msg_info As String)
' ---------------------------------------------------------------------------------
' Returns all matter to build a proper error message.
' msg_line:    at line <err_line>
' msg_no:      1 to n (an Application error translated back into its origin number)
' msg_title:   <error type> <error number> in:  <error source>
' msg_details: <error type> <error number> in <error source> [(at line <err_line>)]
' msg_dscrptn: the error description
' msg_info:    any text which follows the description concatenated by a ||
' ---------------------------------------------------------------------------------
    If InStr(1, err_source, "DAO") <> 0 _
    Or InStr(1, err_source, "ODBC Teradata Driver") <> 0 _
    Or InStr(1, err_source, "ODBC") <> 0 _
    Or InStr(1, err_source, "Oracle") <> 0 Then
        msg_type = "Database Error "
    Else
      msg_type = IIf(err_no > 0, "VB-Runtime Error ", "Application Error ")
    End If
   
    msg_line = IIf(err_line <> 0, "at line " & err_line, vbNullString)
    msg_no = IIf(err_no < 0, err_no - vbObjectError, err_no)
    msg_title = msg_type & msg_no & " in:  " & err_source
    msg_details = IIf(err_line <> 0, msg_type & msg_no & " in: " & err_source & " (" & msg_line & ")", msg_type & msg_no & " in " & err_source)
    msg_dscrptn = IIf(InStr(err_dscrptn, CONCAT) <> 0, Split(err_dscrptn, CONCAT)(0), err_dscrptn)
    If InStr(err_dscrptn, CONCAT) <> 0 Then msg_info = Split(err_dscrptn, CONCAT)(1) Else msg_info = vbNullString

End Sub
```

Used in a procedure which displays an error message will look as follows:

```
Private Sub ErrMsg(ByVal err_no As Long, _
                   ByVal err_source As String, _
                   ByVal err_dscrptn As String, _
                   ByVal err_line As Long)
' -----------------------------------------------
' Displays a proper error message.
' -----------------------------------------------
    Dim sTitle      As String
    Dim sDetails    As String
    
    ErrMsgMatter err_source:=err_source, err_no:=err_no, err_line:=err_line, err_dscrptn:=err_dscrptn, _                         msg_title:=sTitle, msg_details:=sDetails
    
    MsgBox Prompt:="Error description:" & vbLf & _
                    err_dscrptn & vbLf & vbLf & _
                   "Error source/details:" & vbLf & _
                   sDetails, _
           buttons:=vbOKOnly, _
           Title:=sTitle
End Sub
```
[1]: https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/20/Common-VBA-Error-Handler.html