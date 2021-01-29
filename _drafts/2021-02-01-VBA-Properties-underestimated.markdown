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

## Properties as a service providing means
Let's start with a _Standard Module_ called _mFile_ with a service called _Txt_ which read from and write to a file a string, optionally intermitted by CrLf. It's use:<br>
 `s = mFile.Txt(ft_file:=myfile) ' read`<br>
 `mFile.Txt(ft_file:=myfile) = s ' write`
demonstrates the potential for an optimum transparent code - not only in Class Modules but for services in common.

## Implementation of the service

```
Option Explicit

Public Property Let Txt( _
         Optional ByVal ft_file As File, _
                  ByVal s As String)
    
End Property

Public Property Get Txt( _
         Optional ByVal ft_file As File) As String

    
End Property
```

## Conclusion