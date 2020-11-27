---
layout: post
title: Common VBA Components
subtitle: Ready for use, highly reusable, carefully tested
date: 2020-11-27
categories: vba common
---

## Introduction
_Common VBA Components_ provide an enormous advantage for VB-Projects provided they are well designed and carefully tested. Released from any economic pressure (retired) I do spend most of my time on their development, test, improvement, and (recently started) publication. Although they all had been used for a long time the decision to publish them adds a considerable effort to ensure quality, completeness, and consistency.

## The environment
Ever since I've developed, maintained, and tested each _Common VBA Components_ as an individual VB-Project in a dedicated _Common Component Workbook_ in a dedicated _Common Component Project Folder_. With the (late in life) move to Github the folder became the repo clone. Consequently, now I try to do any modification via a branch in order not to interfere with productive VB-Projects using them.

## The management of Common Components
When agreed that _Common Components_ are fine the crucial question is how to keep them up-to-date in VB-Projects using them. Replacing a component in a VB-Project by a more up-to-date version is quite tricky because it cannot be done by the VB-Project themselves but requires a second instance providing this service.

## The Common VBA Components Manager
Publishing the _Common Component Management_ (CompMan) Workbook I use as Add-In to keep used Common Components up-to-date in VB-Projects using them will be one of my future tasks.

## My Common Components
|Commmon| Download|GitHub repo|Service| Description |
|---------|---------|------|-------|----|
|Common VBA Message Service|fMsg<br>mMsg.bas|[public][3r]|Dsply | Display a message|
|Common VBA Error Handling Services|[mErH.bas][1d1]<br>[fMsg.frm][1d2]<br>[fMsg.frx][1d3]<br>[mMsg.bas][1d4]|[public][1r]|[ErrMsg][1s1]|Processing the error message|
||||[BoP, EoP][1s2]| Indicate Begin and End of a Procedure|
||||BoTP| Indicate Begin of Test Procedure|
|Common VBA Execution Trace Srvices|mTrc<br>fMsg<br>mMsg|[public][2]|BOP | Indicate Begin of Procedure |
||||EoP | Indicate End of Procedure|
||||BoC | Indicate Begin of Code|
||||EoC| Indicate End of Code|
|Common VBA Excel Workbook Services|mWbk|private| | |
|Common VBA File Services|mFile|private|- Exists|File existence check|
||||sDiff| Compare Files|
||||ToArray| File to array|
||||FileSelect| File select dialog|
|Common VBA Excel Worksheet Services|mWsh|private| | |
|Common VBA Excel Obstructions Service|mObstrctns|private|| |
|Common VBA Excel Rows Services|mRows|private| | |
|Common VBA Excel Range Services|mRng|private| | |
|Common VB Project Services|mVBP<br>clsVBP|private| | |
|Common VBA Basic Services|mBasic|private|||
|||||
|||||

still to be continued.

[1r]:https://github.com/warbe-maker/Common-VBA-Error-Handler-Services
[1s1]: https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/21/Common-VBA-Error-Handler.html#the-errmsg-service
[1b]: https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/21/Common-VBA-Error-Handler.html#the-beginend-of-procedure-services-bop-eop
[1d1]: (https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mErH.bas)
[1d2]: (https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm)
[1d3]: (https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frx)
[1d4]: (https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mMsg.bas)
[2]: https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
[3r]: https://github.com/warbe-maker/Common-VBA-Message-Service