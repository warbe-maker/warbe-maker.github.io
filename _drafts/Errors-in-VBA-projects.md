---
layout: post
title: Err.Number, vbObjectError, and Err.Source
subtitle: Error numbers and error source in VBA projects
---
<small>All aspects of this post are part of  the public [Common VBA Error Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler)</small>. 

### vbObjectError and Err.Number
- Microsoft documentation says, the error number raised by means of ```Err.Raise``` should be the sum of n +  [_vbObjectError_](<https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1>) in order to avoid conflicts with  is a _VB Runtime Error_. I call such an error an _Application Error_
- When the source of the error is known, i.e. the module and the procedure application errors can be source specific and range from number 1 to n

What does this mean for an error handling?

1. A negative Err.Number can be identified as an _Application Error_ and thus can be displayed in the title of the error message as<br>```Application Error n in <error source> at line m```
2. Because a huge negative number is pretty inappropriate when displayed vbObjectErr should be subtracted which translates it back to the original _Application Error_ number.

### _Err.Source_
The _Source_ property of the _Err_ object unfortunately not provides what it's pretending. The only way to get the source of the error as \<module>.\<procedure> is to do it "manually" (see [Error handling in any procedure](#error-handling-in-any-procedure)

### _Err.Number_ and _Err.Source_ wrapped up in an ErrHndlr Module
In fact the Err.Source may only be used to identify a _Database Error_ so that it is not used for the "pretended" purpose.
```vbscript
Public Function AppErr(ByVal lNo As Long) As Long
    If lNo < 0 _
    Then AppErr = lNo - vbObjectError _
    Else AppErr = lNo + vbObjectError lNo
End Function

Private Function ErrMsgTitle( _
                 ByVal errornumber As Long, _
                 ByVal errorsource As String, _ 
        Optional ByVal errorline As Long = 0) As String
        
   If Err.Number < 0 _
   Then ErrMsgTitle = "Application Error " & Err.Number + vbObjectError _
   Else ErrMsgTitle = "VB Runtime Error " & Err.Number
   ErrMsgTitle = ErrMsgTitle & " in " & errorsource
   If errorline <> 0 _
   Then ErrMsgTitle = ErrMsgTitle & " at line " & errorline
   
End Function

Public Function ErrDisplay( _
                ByVal errornumber As Long, _
                ByVal errorsource As String, _
                ByVal errordescription As String Long _
       Optional ByVal continue As Variant = vbOkOnly) As Variant

   ErrDisplay = _
      MsgBox(title:=ErrMsgTitle(errornumber,errorsource, errorline, continue), prompt:=errordescription, buttons:=replies)
   
End Function
```
#### Error handling in any procedure
```vbscript
Private Sub Any
   Const PROC = "Any"
   On Error Goto on_error
   ...
   If ... _
   Then Err.Raise AppErr(1), "Error description" ' example of an "Application Error"
   ...

end_proc:
   Exit Sub
   
on_error:
   ErrDisplay Err.Number, ErrSrc(PROC), Err.Description, Erl, vbOkOnly
   ' or in case of a continue option
   ' Select Case ErrDisplay(...., continue:=vbYesNo)
   '    Case vbYes
   '    Case vbNo
   ' End Select
End Sub

Private Property ErrSrc(Optional ByVal s As String) As String
    ErrSrc = "<module name>" & "." & s
End Function
```
