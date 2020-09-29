---
layout: post
title: Add items to a Dictionary "ordered"
subtitle: Adding item to a Dictionary by any sequence order without extra sorting
date:   2020-09-25 16:00 +0200
categories: vba basic
---

Demands regarding quality, stability, completeness, etc. escalate when a VBA procedure or module is about to be published.

In many cases, i.e. when entries are not several hundreds, collecting items in a Dictionary instantly ordered is an option - instead of sorting them finally. The procedure _DctAdd_ in the module _mBasic.bas_ offers ascending/descending either by key or by item whereby both may also be an object - provided the object has a name property. It also offers the explicit add before/after a specific target entry (key or item) and all either case sensitive or case ignored.

A full test environment is available in the dedicated _Common Component Workbook_ Basic.xlsm. Testing had proven that the procedure works fine even when the ordered key or item is an object.

### Syntax
```
DctAdd dictionary, key, item, order[, keepduplicates]
```
| Argument | Description |
| -------- | ----------- |
| dct      | A variable declared as Dictionary. Initialized when not already done. Returned by the procedure with the provided item added.|
| dctkey   | |
| dctitem. | |
| dctorder | |
| dcttarget | |
| keepduplicates | |


### Named arguments


### Usage example
The _VBComponents_ of ThisWorkbook are added ordered in ascending sequence case sensitive (requires a reference to "") 
```vbscript
Private Sub DctAddExample()

   Dim dct As Dictionary
   Dim vbc As VBComponent
   
   For each vbc in This workbook.VBProject.VBComponents
      ' DctAdd dct, vbc, vbc.name
      ' would be the equivalent of
      ' dct.Add cbc, cbc.Name
      ' The items are added in entry
      DctAdd dct, vbc, vbc.name, dct_ascendingcasesensitive
      ' adds each item in ascending
      ' sequence         
   Next vbc
```