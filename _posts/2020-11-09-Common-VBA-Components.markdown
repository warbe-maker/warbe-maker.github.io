---
layout: post
title: Common VBA Components
date:          2021-02-19
modified_date: 2022-02-15
categories:    vba common
---
A true development performance boost provided they are well designed, continuously maintained and carefully tested.
<!--more-->

## Preface
### Disambiguation
> A _Common Component_ has the same content in any VB-Project using it. It is developed, maintained, and tested in ***one*** specific -  preferably dedicated - Workbook/VB-Project.<br>A component/module just having the same name with different code is ***not*** a _Common Component_ in the subsequent sense.

### _Common Components_ 
- had initially been developed when it seemed appropriate
- had been maintained and extended every now and then
- has its dedicated VB-Project which includes a test environment and an unattended Regression Test
- is kept in a public GitHub repo of which I use clones
- meets a consistent coding standard and follows clean code principals (no defaults, early binding, avoiding unintended 'case' changes, etc.)

## Managing _Common Components_
I use a _[Common Component Management][1]_ Workbook which is saved as _Addin_ and provides - amongst others - the service to _Update Outdated Common Components_. A bit sophisticated but well for the  job.

## My _Common Components_ (overview)

|Component                          |Module(s)       |Status                 |Comment               |
| --------------------------------- | -------------- | --------------------- | -------------------- |
|Common VBA Message Services        |mMsg, fMsg      |[public GitHub repo][2]|Used by mErH (optionally by mTrc |
|Common VBA Error Services          |mErH, mMsg, fMsg|[public GitHub repo][3]|Optionally uses mTrc|
|Common VBA Execution Trace Services|mTrc            |[public GitHub repo][4]|stand-alone or as optional component of mErH|
|Common VBA Excel Workbook Services |mWrkbk          |[public GitHub repo][5]|Existence/open check over multiple Excel instances, open services and other|
|Common VBA File Services           |mFile           |[public GitHub repo][6]|Existence check, etc.|
|Common VBA Basic Services          |mBasic          |private GitHub repo    | 
|Common VBA Registry Services       |mReg            |private GitHub repo    | Read/write named values simplified to the max |
 
See also: [Conflicts with personal and public _Common Components_][9]

### Comments
Comments are welcome. I apologize for the fact that commenting requires a login to GitHub. This seems to be the only way to keep away spammers.
 

[1]:https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services
[2]:https://github.com/warbe-maker/Common-VBA-Message-Service
[3]:https://github.com/warbe-maker/Common-VBA-Error-Services
[4]:https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
[5]:https://github.com/warbe-maker/Common-VBA-Excel-Workbook-Services
[6]:https://github.com/warbe-maker/Common-VBA-File-Services
[7]:https://github.com/warbe-maker/Common-VBA-Basic-Services
[8]:https://github.com/warbe-maker/Common-VBA-Registry-Services
[9]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html