---
layout: post
title: Common Components
subtitle: Ready for use, highly reusable, completely tested
date: 2020-11-09
categories: vba common
---

They had become a real passion - and an environment on their own. Free from any economic pressure (retired) I do spend most of my time on their development, test, improvement, and publication. Though already used for a long time the decision to publish them adds a significant effort to ensure a certain quality.

Each of these _Common Components_ is developed, maintained, tested, and finally published as an individual VB-Project hosted as repo on GitHub and locally represented as clone in a dedicated _Common Component Project Folder_. Any modification is done via a branch in order not to interfere with  productive VB-Projects using them.

A _Common Component Management_ Workbook used as Add-In provides the means to keep _Common Components_ up-to-date in all VB-Projects using them.

|Component|Module(s)|Status|Comment|
|---------|---------|------|-------|
|Common VBA Message UserForm|fMsg|public repo|Used by mErH, mTrc |
|Common VBA Error Handler|mErH, fMsg|public repo|Optionally uses mTrc|
|Common VBA Execution Trace|mTrc, fMsg|public repo|stand-alone or as optional component of mErH|
 |Common VBA Workbooks|mWrkbk, mErH, fMsg|private repo|Existence/open check over multiple Excel instances, open services and other|
 |Common VBA File|mFile, mErH, fMsg|private repo|Existence check, etc.|
 |Common VBA Worksheet|mWsh, mErH, fMsg|private repo|

still to be continued.


