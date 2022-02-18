---
layout:        post
title:         Worry free merged areas
date:          2021-12-06
modified_date: 2021-12-06
---
<!--more-->

Get rid of any worry with merged cells in VBA Projects.

Not using the great design feature 'merged cells' is the credo of numerous developers throughout forums. And they are right - as long as they have not found a way to properly manage this kind of 'obstruction'. I was not prepared to stay away from merged cells and implemented a way to manage them.

A service called _MergedAreas_ temporarily _eliminates_ (un-merges) merged cells and finally _restores_ (re-merges) them in the meantime even allowing to delete concerned rows or insert new rows straight in the middle of a merged area.

The _MergedAreas_ service (among others) is available in the [_mObstructions.bas_][1] which may be downloaded and imported into any Excel VBA Project. See the README in the [public Github repository][2] for details.

This is one of the obstruction services successfully used in a _Common Rows Component_. The component provides row services like _move up_/_move down_, _insert new_, or _delete_ rows in protected Worksheets specifically focusing on the preservation of total formulas. However, that's another story still pending to go public.

[1]:https://gitcdn.link/cdn/warbe-maker/Common-Excel-VBA-Obstructions-Services/master/source/mObstructions.bas
[2]:https://github.com/warbe-maker/Common-Excel-VBA-Obstructions-Services