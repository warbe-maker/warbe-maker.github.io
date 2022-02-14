---
layout: post
title: Common VBA Components
date:          2021-02-19
modified_date: 2022-02-14
categories:    vba common
---
A true development performance boost provided they are well designed, continuously maintained and carefully tested.
<!--more-->

## Introduction
### Disambiguation
> A _Common Component_ has the same content in any VB-Project using it. It is developed, maintained, and tested in ***one*** specific -  preferably dedicated - Workbook/VB-Project.<br>A component/module just having the same name with different code is ***not*** a _Common Component_ in the subsequent sense.

### My _Common Components_ 
- had initially been developed when it seemed appropriate
- had been maintained and extended every now and then
- has its dedicated VB-Project which includes a test environment and an unattended Regression Test
- is kept in a public GitHub repo of which I use clones
- meets a consistent coding standard and follows clean code principals (no defaults, early binding, avoiding unintended 'case' changes, etc.)

### How to keep them up-to-date in VB-Projects using them?
I use a _[Common Component Management][1]_ Workbook which is saved as _Addin_ and provides - amongst others - the service to _Update Outdated Common Components_. A bit sophisticated but well for the  job.

## Personal and public use of (my) _Common Components_
I do not like the idea maintaining different code versions of _Common Components_, one which I use in my VB-Projects and another 'public' version. On the other hand I do not want to urge users of my _Common Components_ to also use the other _Common Components_ which have become a de facto standard for me.

### Managing the splits
The primary goal is to provide _Common Components_ which are as autonomous as possible by allowing to optionally use them in a more sophisticated environment. This is achieved by a couple of procedures which only optionally use other _Common Components_ when also installed which is indicated by the use of a couple of _Conditional Compile Arguments_:

| Conditional<br>Compile&nbsp;Argument | Purpose |
| ------------------------------------ | ------- |
| _Debugging_                          | Indicates that error messages should be displayed with a debugging option allowing to resume the error line |
| _ExecTrace_                          | Indicates that the _[mTrc][4]_ module is installed
| _MsgComp_                            | indicates that the _[mMsg][3]_, _[fMsg.frm][1]_, and _[fMsg.frx][2]_ are installed |
| _ErHComp_                            | Indicates that the _[mErH][6]_ is installed |

By these means other users are no bothered by my personal preferences - or are only as little as possible :-).

## _Common Components_ overview
|Component|Module(s)|Status|Comment|
|---------|---------|------|-------|
|Common VBA Message Services |mMsg, fMsg |[public GitHub repo][2]|Used by mErH (optionally by mTrc |
|Common VBA Error Services|mErH, mMsg, fMsg|[public GitHub  repo][3]|Optionally uses mTrc|
|Common VBA Execution Trace Services|mTrc |[public GitHub repo][4]|stand-alone or as optional component of mErH|
 |Common VBA Excel Workbook Services|mWrkbk|[public GitHub repo][5]|Existence/open check over multiple Excel instances, open services and other|
 |Common VBA File Services|mFile|[public GitHub repo][6]|Existence check, etc.|
 |Common VBA Basic Services|mBasic|private GitHub repo| 
 [Common VBA Registry Services|mReg|private GitHub repo| Read/write named values simplified to the max
 

[1]:https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services
[2]:https://github.com/warbe-maker/Common-VBA-Message-Service
[3]:https://github.com/warbe-maker/Common-VBA-Error-Services
[4]:https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
[5]:https://github.com/warbe-maker/Common-VBA-Excel-Workbook-Services
[6]:https://github.com/warbe-maker/Common-VBA-File-Services
[7]:https://github.com/warbe-maker/Common-VBA-Basic-Services
[8]:https://github.com/warbe-maker/Common-VBA-Registry-Services