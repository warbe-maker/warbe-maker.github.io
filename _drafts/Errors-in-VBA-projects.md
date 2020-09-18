---
layout: post
title: Errors in VBA projects
subtitle: Error numbers and error source in VBA projects
---
First of all _VB Run Time Errors_ should be distinguished from _Application Errors_ because they have a different reason and a different handling.

### VB Run Time Errors 
- Are caused by an incorrect use of Visual Basic and/or VBA
- Are (or should be) trapped by an error handler (On Error Goto ...)
- Can only be avoided by sufficient testing  (white box and boundary testing at the minimum)

### Application Errors
- Are caused by an incorrect application or usage of any kind of procedure usually by the passed arguments
- Are foreseeable during coding and thus can be handled by the explicit raise of an error (Err.Raise)
- May be avoided by making it impossible to pass invalid arguments or are trapped by an error handler (On Error Go-to ...)

### Error Handler covering both
When both kinds of error are handled by the only one error handler it makes sens to distinguish them.

VBA offers the [vbObjectError](<https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1>) constant (-2147221504) to be used for the distinction of the two kinds of error. While the reason to use it is clear I found it difficult to implement it's use appropriately.

When the vbObjectError constant is added to an Application Error Number let's say 1 the result is an error number -2147221503. When the error is displayed, a negative number can thus be identified as an _Application Error_ and turned back into the origin number:
```vbscript

```
The advantage of this approach is obvious. Provided the error source (the procedure where the error occurred) is known and displayed with the error message each procedure can have its own _Application Error Numbers_ ranging from 1 to n.

Since VB does not provide any means to obtain the procedures and the modules name the following should be mandatory for each module and procedure:

