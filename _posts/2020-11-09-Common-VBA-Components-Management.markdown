---
layout: post
title: Common VBA Components Management
subtitle: Ready for use, highly reusable, completely tested
date: 2020-11-09
categories: vba common
notoc: true
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
|Commmon VBA ... Services|GitHub repo|Services|
|---------|---------|------|-------|
|Message<br>-fMsg<br>-mMsg|public|Dsply |
|Error Handling<br>mErH|public|ErrMsg<br>Bop<br>EoP<br>BoTP|
|Execution Trace<br>mTrc|public|BOP<br>EoP<br>BoC<br>EoC|
 |Workbook<br>mWbk|private| |
 | File<br>mFile|private repo|Exists<br>sDiff<br>ToArray<br>FileSelect<br>|
 |Worksheet<br>mWsh|private|
 |Excel Obstructions<br>mObstrctns|private||
 |Excel Rows<br>mRows|private||
 |Excel Range<br>mRng|private||
 |Project<br>mVBP|private|| 

still to be continued.