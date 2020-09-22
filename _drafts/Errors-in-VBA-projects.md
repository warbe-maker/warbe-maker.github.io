---
layout: post
title: Errors in VBA projects
subtitle: Error numbers and error source in VBA projects
---
<small>All aspects of this post are part of  the public [Common VBA Error Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler)</small>. 

### VB Runtime and other errors
First of all I prefer to distinguish _VB Runtime Errors_ from _Application Errors_.

- VB Runtime Errors are raised by VB and are caused by coding deficiencies.
- Application Errors are caused by an incorrect application or usage of any kind of procedure, foreseeable and thus can be trapped by the explicit raise of an error (```Err.Raise```)

### Error Handler covering both
An error handling is able to distinguish _VB Runtime_ from _Application Errors_ by means of the [vbObjectError](<https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1>) constant (-2147221504) which is to be added to the _Application Error Number_.<br>
As a result an _Application Error Number_ 1 becomes the number -2147221503 - which is quite inappropriate to be displayed in an error message. When the error is displayed, a negative number can be identified as an _Application Error_ and translated back into the origin positive number by subtracting vbObjectError. Both directions are provided by:
```vbscript
Public Function AppErr(ByVal lNo As Long) As Long

    IIf lNo < 0, _
        AppErr = lNo - vbObjectError, _
        AppErr = vbObjectError + lNo
End Function
```
The error is raised by ```Err.Raise AppErr(1), ....``` and translated back when the error us displayed. In connection with a clear identification of the error source each procedure can have its own _Application Error Numbers_ ranging from 1 to n.

#### Error source
Since VB does not provide any means to obtain the procedures and the modules name the following should be mandatory for procedure:

and the following for each module:
```vbscript
Private Property ErrSrc(Optional ByVal s As String) As String
    ErrSrc = "<module name>" & "." & s
End Function
```
Used as follows:
```vbscript
code
```