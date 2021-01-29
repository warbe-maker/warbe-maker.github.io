---
layout: post
title: "VBA Properties: An underestimated and undervalued concept"
date:   2021-02-01
categories: vba property optional argument
---

## Abstract
The commonly used example for a _Property_ in a _Class Module_:
```
Option Explicit
Private sCustName As String
Public Property Let CustName(ByVal s As String)
    sCustName = s
End Property
Public Property Get CustName() As String
    CustName = sCustName
End Property
```
provides a very basic understanding of the concept but unfortunately hides it's potential it is able to unfold when also used in a _Standard Module_ particularly when combined with optional arguments.

## Example service based on Property Let/Get with optional arguments
In a _Standard Module_ called _mFile_ a service called _Txt_ returns the content of a text file as string or writes a string to a file, optionally appended. The module [mFile][1d1] may be downloaded for the full implementation.

```VB
Public Property Let Txt( _
         Optional ByVal ft_file As Variant, _
         Optional ByVal ft_append As Boolean = True, _
         Optional ByRef ft_split As String, _
                  ByVal ft_string As String)
' -----------------------------------------------------
' Writes the string (ft_string) into the file (ft_file)
' which might be a file object of a file's full name.
' Note: ft_split is not used but specified to comply
'       with the Get Property declaration.
' -----------------------------------------------------

    ' see the download for the full implementation

End Property

Public Property Get Txt( _
         Optional ByVal ft_file  As Variant, _
         Optional ByVal ft_append As Boolean = True, _
         Optional ByRef ft_split As String) As String
' ----------------------------------------------------
' Returns the text file's (ft_file) content as string
' with VBA.Split() string in (ft_split).
' Note: ft_append is not used but specified to comply
'       with the Get Property declaration.
' ----------------------------------------------------

    ' see the download for the full implementation
    
End Property
```
## Usage of the service
## Example Read
```VB
s = mFile.Txt(ft_file:=myfile)
```

## Example write
```VB
mFile.Txt(ft_file:=myfile) = s
```

## Summary
1. _Properties_ are not only for _Class Modules_ but can play a significant role in _Standard Modules_
2. _Properties_ may have any number of optional arguments provided
   - they are declared consistently in the Get and Let Property
   - they are declared in the Let Property <u>before</u> the assignment variable.<br><br>Attention! This is in contrast to 'normal' declarations where they have to be declared at last


[1d1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-File-Services/master/mFile.bas