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

`DctAdd dictionary, key, item[, order][,casesensitive][, duplicates]`

The procedure has these names arguments:

| Part | Description |
| -------- | ----------- |
| dct      |  	Required. Always the name of a Dictionary variable or object. When not an object a new Dictionary is established. Dictionary object  returned with the provided key/item pair added.|
| dctkey   | Required. The key associated with the item being added. May be numeric, string, or an object.  |
| dctitem. | Required. The item associated with the key being added. May be numeric, string, or an object. |
| dctorder | If dctorder is omitted it defaults entry sequence.   |
| dctcasesensitive | Optional. Boolean. Defaults to True. When provided False the key/item is added case ignored |
| keepduplicates | Optional. Boolean. Defaults to True.<br>False = when the same item is added with a different key the item is replaced by the key/item pair<br>True = when the same item is added with a key which does not exist, the key/item pair is added|
| dcttarget | Optional. Target when the dctorder is add before by key/item or add after key/item. When not provided along with such a dctorder the dctorder is changed to addascending_bykey/item adddescending_bykey/item  |


### Settings

The dctorder argument settings are:

| Constant          	| Description |
| ----------------- | ----------- |
| after_byitem.     | |
| after_bykey       | |
| before_byitem     | |
| before_bykey      | |
| ascending_bykey   | Performs an add operation with the key/item pair added/inserted ascending by key.|
| ascending_byitem  | |
| descending_bykey  | |
| descending_byitem | |


### Examples
#### Entry sequence
In the below example the _VBComponents_ of _ThisWorkbook_ are added ordered in entry sequence (the default):
```vbscript
Private Sub DctAddExample()

   Dim dct As Dictionary
   Dim vbc As VBComponent
   
   For each vbc in ThisWorkbook.VBProject.VBComponents
      DctAdd dct, vbc, vbc.name        
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