---
layout: post
title: Errors in VBA projects
subtitle: Error numbers and error source in VBA projects
---
<small>The aspects of this blog are part of  the [github repo Common VBA Error Handler](https://github.com/warbe-maker/Common-VBA-Error-Handler)</small>. 

First of all _VB Runtime Errors_ should be distinguished from _Application Errors_ because they have a different reason and a different handling.

### VB Runtime Errors 
- Are caused by an incorrect use of Visual Basic and/or VBA
- Are (or should be) trapped by an error handler (On Error Goto ...)
- Can only be avoided by sufficient testing  (white box and boundary testing at the minimum)

### Application Errors
- Are caused by an incorrect application or usage of any kind of procedure usually by the passed arguments
- Are foreseeable during coding and thus can be handled by the explicit raise of an error (Err.Raise)
- May be avoided by making it impossible to pass invalid arguments or are trapped by an error handler (On Error Go-to ...)

### Error Handler covering both
When both kinds of error are handled by the same error handler it makes sense to distinguish them.<br>
For this VBA offers the [vbObjectError](<https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1>) constant (-2147221504) which is to be added to the _Application Error Number_ for the distinction.<br>
When the _vbObjectError_ constant is added to an _Application Error Number_ let's say 1 the result is an error number -2147221503 - quite inappropriate to remember. When the error is displayed, a negative number can be identified as an _Application Error_ but should be translated back into the origin positive number by subtracting vbObjectError. Both directions are provided by:
```vbscript
Public Function AppErr(ByVal lNo As Long) As Long

    IIf lNo < 0, _
        AppErr = lNo - vbObjectError, _
        AppErr = vbObjectError + lNo
End Function
```
The advantage of this approach is that each procedure can have its own _Application Error Numbers_ ranging from 1 to n.

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
