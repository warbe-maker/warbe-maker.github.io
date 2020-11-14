---
layout: post
<<<<<<< HEAD
title: All the matter if a VB error message
subtitle: Error number, Error type, Error source, etc.
=======
title: The VBA error message matter
subtitle: Error number, Error source, Error type, etc. in VB projects
>>>>>>> branch 'master' of https://github.com/warbe-maker/warbe-maker.github.io
date: 2020-11-13
---
<small>All aspects of this post are part of the [Common VBA Error Handler](https://warbe-maker.github.io/vba/common/2020/11/07/Comprehensive-Common-VBA-Error-Handler.html)</small>

### Err.Number
The _Number_ property of the _Err_ object may indicate a VB Runtime, a Database, or an Application error. The latter is one explicitly raised by `Err Raise`. Microsoft documentation says, the error number raised by means of `Err.Raise` should be the sum of the application error n +  [_vbObjectError_](<https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1>) in order to avoid conflicts with  is a _VB Runtime Error_. I call such an error an _Application Error_ of which the number is set by:
```vbs
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
```vbs
   Select Case err.Number
       Case AppErr(n) ' an error erased by err.Raise
       Case n         ' a Database or VB runtime error
   End Select
```

### Error source
Provided the source of the error is known and displayed with the error message, Application Errors may range from 1 to n in each procedure which keeps the error number handling as simple as possible. Because the _Source_ property of the _Err_ object  (`err.Source`) unfortunately does not deliver what it's name promises we require the following function in each module:
```vbs
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "module-name." & s
End Function
```

### Error type
A proper error message will display the type of error with the number as<br>
\<error type> \<error number><br>
whereby the error type may be _Application error_, _VB Runtime error_, or _Database error_.


### All matter for an error message

The below procedure returns everything potentially usefully to display a proper error message:

```vbs

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

```vbs
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
