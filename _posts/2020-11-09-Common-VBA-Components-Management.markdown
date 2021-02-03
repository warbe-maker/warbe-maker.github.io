---
layout: post
title: Common VBA Components
subtitle: Ready for use, highly reusable, carefully tested
date: 2020-11-27
categories: vba common
show_excerpts: False
---

## Introduction
_Common VBA Components_ provide an enormous advantage for VB-Projects provided they are well designed and carefully tested. Released from any economic pressure (retired) I've spent many hours a on their development, test, improvement, and (recently started) publication using GitHub. Although most of them had already been used for a long time, publishing required an extra effort to ensure quality, completeness, and consistency.

## The environment
Ever since development, maintenance, and test is done via individual _Common VBA Component_ VB-Projects in a dedicated _Common Component Workbook_ in a dedicated _Common Component Folder_. This dedication has paid off it's effort last but not least because it allowed the implementation of regression tests for each component. With the (late in life) move to GitHub these folders became the _repo clone_. Using GitHub for the versioning has proofed a developer's dream. Consequently, I now try to do any modification via a branch in order not to interfere with any productive VB-Projects using them.

## The management
When agreed that _Common Components_ are a great means for productivity one of the crucial questions is how to keep them up-to-date in VB-Projects using them. Replacing a component in a VB-Project by a more up-to-date version is a bit tricky because it cannot be done by the VB-Project itself but requires a second instance providing this service.

## The _Common VBA Components Manager_
Publishing the _Common Component Management_ (CompMan) Workbook I use as Add-In to keep used Common Components up-to-date in VB-Projects using them will be one of my future tasks.

## My Common Components

|         Common VBA ...    |Download/Install|GitHub repo|     Service    |      Description                 |
|---------------------------|----------------|-----------|----------------|----------------------------------|
| Basic Services            |mBasic          |private    |                |                                  |
|                           |                |           |                |                                  |
| Error Handling Services   |[mErH.bas][1d1] |[public][1]|-[ErrMsg][1s1]  | Display or pass on error         |
|                           |[fMsg.frm][1d2] |           |-BoP, EoP       | Indicate Begin/End of a Procedure|
|                           |[fMsg.frx][1d3] |           |-BoTP           | Indicate Begin of Test Procedure |
|                           |[mMsg.bas][1d4] |           |                |                                  |
| Execution Trace Services  |[mTrc.bas][2d1] |[public][2]|-BOP            |Indicate Begin of Procedure       |
|                           |[fMsg.frm][2d2] |           |-EoP            |Indicate End of Procedure         |
|                           |[fMsg.frx][2d3] |           |-BoC            |Indicate Begin of Code            | 
|                           |[mMsg.bas][2d4] |           |-EoC            |Indicate End of Code              |
|VBA File Services          |[mFile][4d1]    |[public][4]|-Exists         | File existence check             |
|                           |                |           |-Differs        | Compare Files                    |
|                           |                |           |-Arry           | File to/from array               |
|                           |                |           |-FileSelect     | File select dialog               |
|                           |                |           |-Tmp            | File select dialog               |
|                           |                |           |-Txt            | File to/from text                |
|VBA Message Service        |fMsg            |[public][3]|-Dsply          | Display a structured message     |
|                           |mMsg.bas        |           |-Box            | Display (Msg)Box analog message  |
| Excel Obstructions Service|mObstrctns      |private    |                |                                  |
| Excel Range Services      |mRng            |private    |                |                                  |
| Excel Rows Services       |mRows           |private    |                |                                  |
| Excel Workbook Services   |mWbk            |private    |                |                                  |
| Excel Worksheet Services  |mWs             |private    |                |                                  |
| Project Services          |mVBP<br>clsVBP  |private    |                |                                  |

still to be continued.

[1]:https://github.com/warbe-maker/Common-VBA-Error-Services
[1r]:https://github.com/warbe-maker/Common-VBA-Error-Handler-Services
[1s1]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/21/Common-VBA-Error-Handler.html#the-errmsg-service
[1b]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/21/Common-VBA-Error-Handler.html#the-beginend-of-procedure-services-bop-eop
[1d1]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mErH.bas
[1d2]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm
[1d3]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frx
[1d4]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mMsg.bas
[2]:https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
[2d1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Execution-Trace-Service/master/mTrc.bas
[2d2]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Execution-Trace-Service/master/fMsg.frm
[2d3]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Execution-Trace-Service/master/fMsg.frx
[2d4]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Execution-Trace-Service/master/mMsg.bas
[3]:https://github.com/warbe-maker/Common-VBA-Message-Service
[4]:https://github.com/warbe-maker/Common-VBA-File-Services
[4d1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-File-Services/master/mFile.bas