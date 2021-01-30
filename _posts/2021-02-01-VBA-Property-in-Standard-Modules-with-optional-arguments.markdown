---
layout: post
title: "VBA Properties: An underestimated and undervalued concept"
date:   2021-02-01
categories: vba property optional argument
excerpt_separator: <!--end-of-excerpt-->
---
A significant extention of VBA _Property_, underestimated, undervalued, and potentially be missed when ignored.
<!--end-of-excerpt-->

## The very basics
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
provides a very basic understanding of the concept of a 'bi-directional _Function_' but may hide it's full potential unfold when it's also used in a _Standard Module_ and particularly when combined with optional arguments. In the example above, in a _Class Module_ the _Property_ value is saved to and read from a _Private_ declared variable. However, the 'value' may as well be an object and the source/target for it may be a Collection, a Dictionary, an Excel Worksheet, a File, ... .

## The full potential unfold

In the following example the _Propertty_ value is written to and read from a _TextFile_. In a _Standard Module_ called _mFile_ the service _Txt_ (the module [mFile][1d1] may be downloaded for the full implementation).

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
The usage of the above outlined example service emphasizes the potential and elegance of _Properties_ with _optional arguments_ in _Standard Modules_.

## Read text string from file
The optionally returned split string argument can be used to transfer the string with VBA.Split into an array.
```VB
Dim s As String
Dim v As Variant
s = mFile.Txt(ft_file:=myfile, ft_split:=sSplit)
v = VBA.Split(s, sSplit)
```

## Write text string to file
When the string is intermitted with _vbCrLf_ the string may represent the full content written to the file. The optional arguments allows appending the string to the file.
```VB
mFile.Txt(ft_file:=myfile, ft_append:=True) = s
```

## Summary
1. _Properties_ are not only for _Class Modules_ but can play a significant role in _Standard Modules_
2. _Properties_ may have any number of optional arguments provided
   - they are declared consistently in the Get and Let Property
   - they are declared in the Let Property <u>before</u> the assignment variable.<br><br>Attention! This is in contrast to 'normal' declarations where they have to be declared at last
3. Considering the fact that the 'value' may be anything and the source/target may be a wide range of things (Worksheet, File, Collection, Dictionary, etc, the applications for _Properties_ 'explode'.


[1d1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-File-Services/master/mFile.bas