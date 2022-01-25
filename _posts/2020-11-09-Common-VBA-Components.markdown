---
layout: post
title: Common VBA Components
date:          2021-02-19
modified_date: 2022-01-25
categories:    vba common
---
A great advantage for the  development of VB-Projects - provided they are well designed, continuously maintained and carefully tested.
<!--more-->

## Introduction
### Clarification of the term _Common Component_
A component just having the same name in various VBProjects is **not** a _Common Component_ in the sense of this post unless these _Used Common Components_ are all based on a _Raw Common Component_ one Workbook/VP-Project has claimed hosting it. And any _Used Common Component_ in whichever VB-Project is a copy of this _Raw Common Component_. The Workbook/VP-Project which hosts the _Raw Common Component_ is responsible for the development, maintenance, and testing - why it preferably should be dedicated.

The very next question: How can it be guaranteed that the _Used Common Components_ are not outdated? The only appropriate solution: When the Workbook is opened for development/maintenance (not for production!) all outdated _Used Common Components_ are updated with the _Raw Common Component_ - which cannot be done by the Workbook itself but only by a "third party Workbook" dedicated providing this update service.

I have started over several times implementing such a Workbook and I've given up the try as many times. Finally I do have a fairly stable and solid _[Excel VBA Component Management][5]_ in place, as a public Github repo for use by everyone. 

## My Common VBA Components
### Public versus personal: Principals of design
I don't like the idea of providing a 'public' version of a component apart from the version I use. _Common Components_ serving personal preferences in my own VB-Projects while being absolutely autonomous in any 'foreign' VB-Project. That's the challenge. To manage this balancing act all my _Common Components_ do use a couple of procedures which either perfectly work in both environments - or simply do nothing. Concerned of this approach are:
#### _Error Handling_
The procedures _AppErr_, _ErrSrc_, and _ErrMsg_ provide a fairly elaborated error handling which includes a debugging option (display of an error with the option to resume the error line) when activated by the _Conditional Compile Argument_ `Debugging = 1`. [^1]

#### _Execution Trace_
Procedures: _BoP_, _EoP_. Corresponding statements do absolutely nothing unless activated by the _Conditional Compile Argument_ `ExecTrace = 1` which requires the _Common Component_ [mTrc.bas][2d1] being downloaded and imported. 




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


[^1]: Im my personal VB-Projects I do use the _Common Components_ _mMsg_, _fMsg_, and _mErH_ which do provide a much "nicer" display of an error, activated by the _Conditional Compile Argument_ `MsgComp = 1` and `ErHComp = 1`.


[1]:https://github.com/warbe-maker/Common-VBA-Error-Services
[1r]:https://github.com/warbe-maker/Common-VBA-Error-Handler-Services
[1s1]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/error/handling/2021/01/16/Common-VBA-Error-Services.html
[1b]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/21/Common-VBA-Error-Handler.html#the-beginend-of-procedure-services-bop-eop
[1d1]:https://gitcdn.link/cdn/warbe-maker/VBA-MsgBox-alternative/master/source/mErH.bas
[1d2]:https://gitcdn.link/cdn/warbe-maker/VBA-MsgBox-alternative/master/source/fMsg.frm
[1d3]:https://gitcdn.link/cdn/warbe-maker/VBA-MsgBox-alternative/master/source/fMsg.frx
[1d4]:https://gitcdn.link/cdn/warbe-maker/VBA-MsgBox-alternative/master/source/mMsg.bas
[2]:https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
[2d1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mTrc.bas
[2d2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frm
[2d3]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frx
[2d4]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mMsg.bas
[3]:https://github.com/warbe-maker/Common-VBA-Message-Service
[4]:https://github.com/warbe-maker/Common-VBA-File-Services
[4d1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-File-Services/master/source/mFile.bas
[5]:https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services