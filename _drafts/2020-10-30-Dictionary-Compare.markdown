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
[Example](#example)<br>
[Development, test, maintenance](#development-test-maintenance)

### Method
Comparing Dictionaries may not be worth another post unless the function's options make a difference:
- Compare either the Keys, the Items or both
- Compare case sensitive or case ignored (for items if type String)
- Accept objects for both, keys an items whereby objects can only constitute a difference when they have a name property
- optionally ignore/skip empty items

### Syntax

`DctDiff dict1, dict2[, criteria][, sense]`

The procedure has these names arguments:

| Part         | Description |
| ------------ | ----------- |
| dict1, dict2 | Obligatory. The two Dictionary objects to be compared
| criteria     | Optional. Defaults to item when omitted. Specifies, what is compared to determine a difference
| sense        | Optional. Defaults to case sensitive |
| ignoreemptyitems| Optional. Boolean. Defaults to False when omitted. When True and the compare criterion is byitem or byentry any empty items `Trim(item) = vbNullString` are skipped.|

### Settings

The order argument settings are:

| Argument | Constant   | Description |
| -------- | ---------- | ----------- |
| criteria | crit_bykey | only the keys are compared           |
|          | crit_byitem| only the items are compared           |
|          | crit_entry |             |
| sense    | sense_caseignored |      |
|          | sense_casesensitive |    |

### Example

### Development, test, maintenance
- The dedicated _Common Component Workbook_ Dct.xlsm is the development, test, and maintenance environment for _DctDiff_ (see the GitHub repo [Common VBA Dictionary Procedures](https://github.com/warbe-maker/Common-VBA-Dictionary-Procedures).
- The procedures _Test\_DctDiff*_ in module _mTest_ provide tests, obligatory after any kind of code modification.