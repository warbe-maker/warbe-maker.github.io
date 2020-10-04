---
layout: post
title: Dictionary Compare
subtitle: Adding item to a Dictionary by any sequence
date:   2020-09-25 16:00 +0200
categories: vba basic
---

In this post<br>
[Method](#method)<br>
[Syntax](#syntax)<br>
[Settings](#settinhs)<br>
[Examples](#examples)<br>
[Development, test, maintenance](#development-test-maintenance)

### Method


### Syntax

`DctDiff dict1, dict2[, criteria][, sense]`

The procedure has these names arguments:

| Part         | Description |
| ------------ | ----------- |
| dict1, dict2 | Obligatory. The two Dictionary objects to be compared
| criteria.    | Optional. Defaults to item when omitted. Specifies, what is compared to determine a difference
| sense        | Optional. Defaults to case sensitive


### Settings

The order argument settings are:

| Argument | Constant   | Description |
| -------- | ---------- | ----------- |
| criteria | crit_bykey |             |
|          | crit_byitem|             |
| sense    | sense_caseignored   |             |
|          | sense_casesensitive |             |


### Examples
#### Entry sequence
In the below example the _VBComponents_ of _ThisWorkbook_ are added ordered in entry sequence (the default):
```vbscript
Private Sub DctAddExample()

   Dim dct As Dictionary
   Dim vbc As VBComponent
   
   For each vbc in ThisWorkbook.VBProject.VBComponents
      DctAdd dct, vbc, vbc ' key and item is an object       
   Next vbc
   
End Sub
```
#### Ascending by key case sensitive
In the below example the _VBComponents_ of _ThisWorkbook_ are added ordered in ascending sequence case sensitive. The order criteria is the name property of the key object:
```vbscript
Private Sub DctAddExample()

   Dim dct As Dictionary
   Dim vbc As VBComponent
   
   For each vbc in ThisWorkbook.VBProject.VBComponents
      DctAdd dct, vbc, vbc.name, ascending_bykey        
   Next vbc
   
End Sub
```

### Development, test, maintenance
- The dedicated _Common Component Workbook_ Dct.xlsm is the development, test, and maintenance environment.
- The procedure _Test\_DctAdd_ in module _mTest_ provides a fully automated regression test, obligatory after any kind of code modification
- The procedure _Test\_DctAdd\_99\_Performance_ in module _mTest_ provides an example for a performance test. In order to trace the execution time the tests make use of  the _mErrHndlr_ module (not required for the _DctAdd_ procedure)
- The _DctAdd_ procedure uses the _ErrMsg_ procedure in module _mBasic_