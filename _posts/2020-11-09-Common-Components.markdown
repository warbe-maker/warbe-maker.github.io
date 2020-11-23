---
layout: post
title: Common VBA Components
subtitle: Ready for use, highly reusable, completely tested
date: 2020-11-09
categories: vba common
notoc: true
---

_Common VBA Components_ had become a real passion. Released from any economic pressure (retired) I do spend most of my time on their development, test, improvement, and (recently started) publication. Although they all had been used for a long time the decision to publish them adds a considerable effort to ensure quality, completeness, and consistency.

Ever since I've developed, maintained, and tested each _Common VBA Components_ as an individual VB-Project in a dedicated _Common Component Workbook_ in a dedicated _Common Component Project Folder_. With the (late in life) move to Github the folder became the repo clone. Consequently, now I try to do any modification via a branch in order not to interfere with productive VB-Projects using them.

A _Common VBA Component_ is potentially used in many VB-Projects and thus deserves every managable testing effort. A properly setup egression test performing all individual test in one go is the most economic way which assures a desired quality. The setup is an effort but it's pretty satisfactory performing a regression test with all assertions automatically provided.

Publishing the _Common Component Management_ (CompMan) Workbook I use as Add-In to keep used Common Components up-to-date in VB-Projects using them will be one of my future tasks.

|Component|Module(s)|Status|Comment|
|---------|---------|------|-------|
|Common VBA Message UserForm|fMsg|public repo|Used by mErH, mTrc |
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