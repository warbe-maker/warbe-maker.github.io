---
layout: post
title: Add key/item pairs to a Dictionary "instantly ordered"
subtitle: Adding item to a Dictionary by any sequence order without extra sorting
date: 2020-10-02 16:00 +0200
categories: vba dictionary common
---
In this post<br>
[Method](#method)<br>
[Syntax](#syntax)<br>
[Settings](#settinhs)<br>
[Usage examples](#usage-examples)<br>
[Installation](#installation)<br>
[Development, test, maintenance](#development-test-maintenance)

### Method
In many cases, specifically when items to be collected in a Dictionary are not several hundreds, adding them instantly ordered is an option. The procedure _DctAdd_ in  module _mDct.bas_ provides:
- ascending, descending, and entry sequence
-  order by key or by item whereby both may also be an object, provided the object has a name property
- case sensitive and case ignored
- add before or add after a specific target entry (by key or by item)
- unique or duplicate items even when sort order is by item, provided the items have different keys
- item update for duplicate keys when ordered by key

### Syntax

`DctAdd dct, key, item[, order][, seq][, sense][, target][, staywithfirst]`

Note: Without any optional argument are the result is identical with<br>`dct.Add key, item`

The procedure has these names arguments:

| Part | Description |
| -------- | ----------- |
| dct      |  	Required. Always the name of a Dictionary variable or object. When not an object a new Dictionary is established. Dictionary object  is returned by the method with the provided key/item pair added.|
| key      | Required. The key associated with the item being added. May be numeric, string, or an object.<br><br>**Note:** When the key is the order criteria and it is an object, the object must have a name property which is used as the sort value. If not an error is raised.  |
| item    | Required. The item associated with the key being added. May be numeric, string, or an object.<br><br>**Note:** When the item is the order criteria and it is an object, the object must have a name property which is used as the sort value. If not an error is raised. |
| order | Optional. Defaults to _order\_bykey_ when omitted. |
| seq    | Optional. Defaults to entry sequence (_seq\_entry_) when omitted. |
| sense   | Optional. Defaults to _case\_sensitive_ when omitted.|
| target | Optional. An existing key or item. When omitted:<br>-When the sequence is _seq\_beforekey_, or _seq\_beforeitem_, the sequence is changed to _seq\_descending_<br>- When the sequence  _seq\_afterkey_, _seq\_afteritem_ the sequence is changed to _seq\_ascending_ |
| staywithfirst | Optional. Boolean. Defaults to False.<br>False:<br>- With _order\_bykey_ any add of an existing key updates the item<br>- With _order\_byitem_ any add if the same item is added provided it has a new key.<br>True:<br>- With _order\_bykey_ any add for an existing key is ignored<br>- With _order\_byitem_ Attention!!! Any add if an existing item is ignored - even when it has a new unique key !!!|

### Settings

| Argument | Constant   | Description |
| -------- | ---------- | ----------- |
| order    | order_bykey         | Items added are ordered by key (default for ascending or descending sequence)   |
|          | order_byitem        | Items added are ordered by item                  |
| seq      | seq_ascending       | Items are added in ascending sequence            |
|          | seq_descending      | Items are added in descending sequence           |
|          | seq_aftertarget     | The item is added after a specified target entry |
|          | seq_beforetarget    | The item is added before a specified target entry|
|          | seq_entry.          | Items are added in entry sequence (default)      | 
| sense    | sense_caseignored   | Items are ordered with case ignored              |
|          | sense_casesensitive | Items are ordered with case sensitive (default)  |


### Usage examples
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
      ' order_bykey and case_sensitive are defaults
      DctAdd dct, vbc, vbc.name, seq:=seq_ascending       
   Next vbc
   
End Sub
```
### Installation
Download [_mDct.bas_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Dictionary-Procedures/master/mDct.bas) and import it into your VB-Project. Alternatively you may fork the [Github repo Common-VBA-Dictionary-Procedures](https://github.com/warbe-maker/Common-VBA-Dictionary-Procedures).

### Development, test, maintenance
- The dedicated _Common Component Workbook_ Dct.xlsm (see Github repo) is the development, test, and maintenance environment.
- The procedure _Test\_DctAdd_ in module _mTest_ provides a fully automated regression test, obligatory after any kind of code modification
- The procedure _Test\_DctAddPerfornance_ in module _mTest_ provides an example for a performance test. In order to trace the execution time the tests make use of  the _mErrHndlr_ module (not required for the _DctAdd_ procedure)
- The _DctAdd_ procedure uses the _ErrMsg_ procedure in module _mBasic_