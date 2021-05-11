---
layout: post
title: Common VBA Components
date:          2021-02-19
modified_date: 2021-04-29
categories:    vba common
---
A great advantage for the  development of VB-Projects - provided well designed, continuously maintained and carefully tested.
<!--more-->

## Introduction
Keeping _Common VBA Components_ up-to-date in VB-Projects using them is  cumbersome - unless done by a service when a Workbook is opened. _Synchronizing_ the code of whole VB-Projects is probability the 'supreme discipline' in this regard but that's the matter if another post.

## Environment
Development, maintenance, and test of  _Common VBA Component_ , is done via dedicated VB-Projects which claim the original/raw component code 'hosted'. This dedication pais off it's effort because it it allowes the implementation of regression tests performed with every code modification. Using GitHub for the versioning has proofed a developer's dream. Consequently, I now try to do any modification via a branch in order not to interfere with any productive VB-Projects using them.


## Management services
Services are provided by a _Common Components Management_ (CompMan) Workbook, setup as Add-In.

## (My) Common VBA Components

|         Common VBA ...    |Download and import|GitHub repo|     Service    |      Description                 |
|---------------------------|----------------|-----------|----------------|----------------------------------|
| Basic Services            |mBasic          |private    |                |                                  |
|                           |                |           |                |                                  |
| [Error Handling Services][1s1] |[mErH.bas][1d1] |[public][1]|-ErrMsg | Display or pass on error to the caller        |
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
[1s1]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/error/handling/2021/01/16/Common-VBA-Error-Services.html
[1b]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/21/Common-VBA-Error-Handler.html#the-beginend-of-procedure-services-bop-eop
[1d1]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/source/mErH.bas
[1d2]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/source/fMsg.frm
[1d3]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/source/fMsg.frx
[1d4]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/source/mMsg.bas
[2]:https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
[2d1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mTrc.bas
[2d2]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frm
[2d3]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frx
[2d4]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mMsg.bas
[3]:https://github.com/warbe-maker/Common-VBA-Message-Service
[4]:https://github.com/warbe-maker/Common-VBA-File-Services
[4d1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-File-Services/master/source/mFile.bas