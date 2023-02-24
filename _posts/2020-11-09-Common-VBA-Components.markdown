---
layout: post
title: Common VBA Components
date:          2021-02-19
modified_date: 2023-02-24
categories:    vba common
---
A true development performance boost provided they are well designed, continuously maintained and carefully tested.
<!--more-->

## Preface
### Disambiguation
> A _Common Component_ has the same content in any VB-Project using it. It is developed, maintained, and tested in ***one*** specific -  preferably dedicated - Workbook/VB-Project.<br>A component/module just having the same name with different code is ***not*** a _Common Component_ in the subsequent sense.

### My _Common Components_ 
- had initially been developed when it seemed appropriate
- had been maintained and extended every now and then
- has its dedicated VB-Project which includes a test environment and an unattended Regression Test
- is kept in a public GitHub repo of which I use clones
- meets a consistent coding standard and follows clean code principals (no defaults, early binding, avoiding unintended 'case' changes, etc.)

## My management of _Common Components_
I use a _[Common Component Management][1]_ Workbook (in a public GitHub repository) which optionally may be saved as _Addin_. It provides the (not only) the service to _Update Outdated Common Components_. A Somehow sophisticated approach but it does the job already for years - and is still supported.

## My _Common Components_ (overview)

|Component                           |Module(s)               | Status                 |Comment               |
| ---------------------------------- | ---------------------- | ---------------------- | -------------------- |
|Common VBA Message Services         |mMsg, fMsg              |[public GitHub repo][2] |Universal message service, used by _mErH_ for instance|
|Common VBA Error Services           |mErH, mMsg, fMsg        |[public GitHub repo][3] |Optionally uses _mTrc_|
|Common VBA Execution Trace Services |mTrc                    |[public GitHub repo][4] |stand-alone or as optional component of mErH|
|Common VBA Excel Workbook Services  |mWbk                    |[public GitHub repo][5] |Existence/open check over multiple Excel instances, open services and other|
|Common VBA File Services            |mFso                    |[public GitHub repo][6] |Files and folder services including PrivateProfile file services|
|Common VBA Basic Services           |mBasic                  |private GitHub repo     |The code is visible via the [CompMan Workbook][1] where the component is used|
|Common VBA Queue and Stack services |mQ, clsQ<br>mStck,&nbsp;clsStck|[public GitHub repo][10]| Stack an Queue usage unified|
|Common VBA Registry Services        |mReg                    |private GitHub repo     | Read/write named values simplified to the max |
 
See also: [Conflicts with personal and public _Common Components_][9]

### Comments
Comments are welcome. I apologize for the fact that commenting requires a login to GitHub. This seems to be the only way to keep away spammers.
 

[1]:https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services/blob/master/README.md?#management-of-excel-vb-project-components
[2]:https://github.com/warbe-maker/Common-VBA-Message-Service
[3]:https://github.com/warbe-maker/Common-VBA-Error-Services
[4]:https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
[5]:https://github.com/warbe-maker/Common-VBA-Excel-Workbook-Services
[6]:https://github.com/warbe-maker/Common-VBA-File-Services
[7]:https://github.com/warbe-maker/Common-VBA-Basic-Services
[8]:https://github.com/warbe-maker/Common-VBA-Registry-Services
[9]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
[10]:https://github.com/warbe-maker/Common-VBA-Queue-and-Stack-services