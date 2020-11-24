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
|Component|Module(s)|Status|Comment|
|---------|---------|------|-------|
|Common VBA Message Form|fMsg (mMsg)|public repo|Used by mErH, mTrc |
|Common VBA Error Handler|mErH, fMsg|public repo|Optionally uses mTrc|
|Common VBA Execution Trace|mTrc, fMsg|public repo|stand-alone or as optional component of mErH|
 |Common VBA Workbooks|mWrkbk, mErH, fMsg|private repo|Existence/open check over multiple Excel instances, open services and other|
 |Common VBA File|mFile, mErH, fMsg|private repo|Existence check, etc.|
 |Common VBA Worksheet|mWsh, mErH, fMsg|private repo|
 |Common VBA Excel Obstructions|mObstrctns|private repo||
 |Common VBA Excel Rows|mRows|private repo||
 |Common VBA Excel Range|mRng|private repo||
 |Common VB-Project|mVBP|private repo|| 

still to be continued.