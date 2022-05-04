---
layout: post
title: Common VBA Error Services
date:          2021-01-16
categories: vba common error handling
modified_date: 2022-05-04
---
An inviting error message with features! This will make a difference for the development of a VBA-Project. Error services inspired by the-best-of-the-web, worth being considered not only by professionals.
<!--more-->

## Preface
The _[Common VBA Error Services][1]_ introduced by this post might appear overdone, too complicated, not worth the effort, etc.. It became my standard throughout all VB-Projects however. On the long run it proved worth the effort, foremost because it helps locating and eliminating an error a snap. The [README][4] in the corresponding [public GiHub repo][1] provides all the details for installation and usage. For experienced developers it will take not more than 30 minutes to finish the task. A true investment for future VBA development. 

## Disambiguation
| Term            | Meaning                                         |
| --------------- | ----------------------------------------------- |
|_Application&nbsp;Error_| An error which had been raised by an `err.Raise` statement distinguished from any system error like VB-Run-time or Database error. In order to avoid conflicts with system error numbers the `vbObjectError` is added to turn it into a negative number. The error service uses an AppErr function for this which turns the negative back into its original positive number for the display of the error. |
|_Entry&nbsp;Procedure_| For the error display this is one of the key issues for the assembling of the _path-to-the-error_.|
|_Error&nbsp;Source_   | An unambiguous name of the procedure which raised an error - by prefixing the procedure name with the module name.|
|_Common&nbsp;Components, Component| Term used for all my [Common VBA Components][2] of which the code is kept identical with all VB-Projects using them. A tough matter but successfully implemented and used. See:  |

## Features
### Summary
The advantage of the error display service may best be depicted by the following:
1. The display of an error message by means of the VBA.MsgBox - already enriched with the debugging option
![](../Assets/DemoAppErrByVBAMsgBox.png)
![](/Assets/DemoAppErrByVBAMsgBox.png)
This is my standard without using any extra modules. Pure VBA in fact.

2. The display of an error message by means of the mMsg.ErrMsg service:
![](../Assets/DemoAppErrByMsgDsply.png)
![](/Assets/DemoAppErrByMsgDsply.png)
Looks much better and the debugging option is much more self explanatory.

3. The display of an error message by means of the _mErH.ErrMsg_ service:
![](../Assets/DemoAppErrByErhErrMsg.png)
![](/Assets/DemoAppErrByErhErrMsg.png)
For complex VBA-Projects and/or those having a (recommended by professionals) fine module structure, having the "path-to-the-error" displayed can make a difference.

### Conlusion
A error message displayed by the mErH.ErrMsg service provides:
- A inviting appearance by design
- A [Path to the error][5]
- An optional _Debugging Button_ for going [straight to the error line][3]
- An optional _About the error_ section for _Application Errors_ (those raised by `Err.Raise`)

## Comments
Any comments are welcome whether here in this blog or in the [public GiHub repo][1] which is open for Discussions. I appologize for the fact that whichever way is used requires an acoount and a login. This is the only way to hinder spammers.

[1]:https://github.com/warbe-maker/Common-VBA-Error-Services
[2]:https://warbe-maker.github.io/vba/common/2021/02/19/Common-VBA-Components.html
[3]:https://warbe-maker.github.io/vba/common/error/handling/2022/02/16/Straight-to-the-Error-Line.html
[4]:https://github.com/warbe-maker/Common-VBA-Error-Services/blob/master/README.md
[5]:https://github.com/warbe-maker/Common-VBA-Error-Services/blob/master/README.md#the-path-to-the-error 