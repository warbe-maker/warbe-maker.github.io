---
layout: post
title: DctAdd: Add key/item pairs to a Dictionary "instantly ordered"
subtitle: Adding item to a Dictionary by any sequence order without extra sorting
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
In many cases, specifically when entries to be added are not several hundreds, collecting items in a Dictionary instantly ordered is an option - especially when this method offers uncommon options. The procedure _DctAdd_ in the module _mDct.bas_ offers ascending/descending either by key or by item whereby both may also be an object, provided the object has a name property. It also offers the explicit add before/after a specific target entry (key or item) and all either case sensitive or case ignored.

### Syntax

`DctAdd dictionary, key, item[, order][, seq][, sense][, target][, staywithfirst]`

The procedure has these names arguments:

| Part | Description |
| -------- | ----------- |
| dct      |  	Required. Always the name of a Dictionary variable or object. When not an object a new Dictionary is established. Dictionary object  returned with the provided key/item pair added.|
| key      | Required. The key associated with the item being added. May be numeric, string, or an object.  |
| item    | Required. The item associated with the key being added. May be numeric, string, or an object. |
| order | Optional. Defaults to by key when omitted. |
| seq    | Optional. Defaults to entry sequence when omitted. |
| sense   | Optional. Defaults to case sensiticve when omitted.|
| target | Optional. Target key or item when seq is add before or add after. When omitted:<br>When the sequence is add before the sequence is changed to descending<b>When the sequence is add after it is changed to ascending. |


### Settings

The order argument settings are:

| Argument | Constant            | Description |
| -------- | ------------------- | ----------- |
| order    | order_bykey         |             |
|          | order_byitem        |             |
| seq      | seq_ascending       | Performs an add operation with the key/item pair added/inserted ascending by key.|
|          | seq_descending      |             |
|          | seq_aftertarget     |             |
|          | seq_beforetarget    |             |
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
- The procedure _Test\_DctAddPerfornance_ in module _mTest_ provides an example for a performance test. In order to trace the execution time the tests make use of  the _mErrHndlr_ module (not required for the _DctAdd_ procedure)
- The _DctAdd_ procedure uses the _ErrMsg_ procedure in module _mBasic_