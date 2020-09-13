---
layout: post
title:  "Application versus Visual Basic error"
date:   2020-09-11 13:49:00 +0200
categories: vba excel snippet
---
This is my very first blogging exercise!

I didn't use Err.Raise because of an uncertainty regarding the handling of programmed application errors in contrast to Visual Basic errors. This mainly had got to do with the use of the vbObjectError constant which appeared bit strange to me. The code below now shows a proper way for dealing with programmed errors.
```vbscript
Public Function AppErr(ByVal lNo As Long, _
              Optional ByRef sError As String = vbNullString) As Long
' -------------------------------------------------------------------
' Usage example when a programmed application error occurs:
'    If ..... Then Err.Raise AppErr(1), ....
' Example when the error message is displayed:
'   AppErr Err.Number, sErrTitle
'   MsgBox title:=sErrTitle
' --------------------------------------------------------------------
    If lNo < 0 Then
        '~~ This is an application error number which had been turned into a negative
        '~~ number in order to avoid any conflict with a VB error. The function returns
        '~~ the original positive application error number and a corresponding title
        AppErr = lNo - vbObjectError
        If Not IsMissing(sError) Then sError = "Application error " & AppErr
    Else
        '~~ This is a positive error number which is regarded as a programmed application
        '~~ error. The function thus returns a negative number in order to avoid any
        '~~  conflict with a VB error.
        AppErr = vbObjectError + lNo
        '~~ For the case the positive lNo is a Visual Basic Error a corresponding string is returned 
        If Not IsMissing(sError) Then sError = "Microsoft Visual Basic runtime error " & lNo
    End If
End Function
```
There may be a slightly better way but this should be a good start. The following two test examples show the usage of the AppErr function - by the way outlining some error handling in general.
```vbscript
Public Sub Test_AppErr()
    On Error GoTo on_error
    Const PROC = "Test_AppErr"
    Dim sErrTitle   As String
    
    Err.Raise AppErr(1), ErrSrc(PROC), "This is a programmed application error raised in the procedure metioned below as the error source."
exit_proc:
    Exit Sub
on_error:
    AppErr Err.Number, sErrTitle
    MsgBox title:=sErrTitle & " in " & Err.Source, _
           prompt:="Error description: " & vbLf & vbLf & Err.Description & vbLf & vbLf & vbLf & _
                   "Error source:" & vbLf & vbLf & ErrSrc(PROC)
End Sub

Public Sub Test_VBErr()
    On Error GoTo on_error
    
    Const PROC = "Test_VBErr"
    Dim sErrTitle   As String
    Dim l           As Long
    
    l = l / 0
    
exit_proc:
    Exit Sub
on_error:
    AppErr Err.Number, sErrTitle
    MsgBox title:=sErrTitle & " in " & ErrSrc(PROC), _
           prompt:="Error description: " & vbLf & vbLf & Err.Description & vbLf & vbLf & vbLf & _
                   "Error source:" & vbLf & vbLf & ErrSrc(PROC)
End Sub
```
The little function ErrSrc addresses the fact that VB only knows the project as the error source and has got no idea of in which module/procedure the error occurred. I do have this function in each module by default.
The advantage of identifying the error source is that each procedure can have its own application error numbered from 1 to n. And thus no need to maintain a list of application error numbers used in a project. 
```vbscript
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTest." & s
End Function
```
When the error source is regarded a property of a module the following will as be appropriate:
```vbscript
Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mTest." & s:  End Property

```

