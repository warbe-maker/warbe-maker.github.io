---
layout: post
title: Err.Number, vbObjectError, and Err.Source
subtitle: Error numbers and error source in VBA projects

---
<small>All aspects of this post are part of the [Common VBA Error Handler](https://warbe-maker.github.io/vba/common/2020/11/07/Comprehensive-Common-VBA-Error-Handler.html)</small>

### vbObjectError and Err.Number
- Microsoft documentation says, the error number raised by means of ```Err.Raise``` should be the sum of n +  [_vbObjectError_](<https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1>) in order to avoid conflicts with  is a _VB Runtime Error_. I call such an error an _Application Error_
- When the source of the error is known, i.e. the module and the procedure application errors can be source specific and range from number 1 to n

What does this mean for an error handling?

1. A negative Err.Number can be identified as an _Application Error_ and thus can be displayed in the title of the error message as<br>```Application Error n in <error source> at line m```
2. Because a huge negative number is pretty inappropriate when displayed in an error message the vbObjectErr should be subtracted from the Err.Number when it is negative which translates it back to the original positive _Application Error_ number.

### _Err.Source_
The _Source_ property of the _Err_ object unfortunately does not provide what it's name promises. The only way to get the source of the error as \<module>.\<procedure> is to do it "manually"

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
' -------------------------------------------------------
' Returns all the matter to build a proper error message.
' -------------------------------------------------------
                
    If InStr(1, err_source, "DAO") <> 0 _
    Or InStr(1, err_source, "ODBC Teradata Driver") <> 0 _
    Or InStr(1, err_source, "ODBC") <> 0 _
    Or InStr(1, err_source, "Oracle") <> 0 Then
        msg_type = "Database Error "
    Else
      msg_type = IIf(err_no > 0, "VB-Runtime Error ", "Application Error ")
    End If
   
    msg_line = IIf(err_line <> 0, "at line " & err_line, vbNullString)     ' Message error line
    msg_no = IIf(err_no < 0, err_no - vbObjectError, err_no)                ' Message error number
    msg_title = msg_type & msg_no & " in " & err_source & " " & msg_line             ' Message title
    msg_details = IIf(err_line <> 0, msg_type & msg_no & " in " & err_source & " (at line " & err_line & ")", msg_type & msg_no & " in " & err_source)
    msg_dscrptn = IIf(InStr(err_dscrptn, CONCAT) <> 0, Split(err_dscrptn, CONCAT)(0), err_dscrptn)
    If InStr(err_dscrptn, CONCAT) <> 0 Then msg_info = Split(err_dscrptn, CONCAT)(1)

End Sub
```

Used in a procedure which displays an error message will look as follows:

```vbs
Private Sub ErrMsg(ByVal err_no As Long, _
                   ByVal err_source As String, _
                   ByVal err_dscrptn As String, _
                   ByVal err_line As Long)

    Dim sTitle As String
    
    ErrMsgMatter err_source:=err_source, err_no:=err_no, err_line:=err_line, err_dscrptn:=err_dscrptn, msg_title:=sTitle
    
    MsgBox Prompt:="Error description:" & vbLf & _
                    err_dscrptn & vbLf & vbLf & _
                   "Error source:" & vbLf & _
                   err_source, _
           buttons:=vbOKOnly, _
           Title:=sTitle
End Sub
```