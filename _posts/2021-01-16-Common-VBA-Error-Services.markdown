---
layout: post
title: Common VBA Error Services
date:          2021-01-16
categories: vba common error handling
modified_date: 2022-02-18
---
No more worry with error messages popping up. An error handling/service inspired by (the best of) the web, worth being considered for any VB-Project.
<!--more-->

## Preface
The _[Common VBA Error Services][1]_ introduced by this post might appear overdone, too complicated, not worth the effort, etc.. However, since this is my personal standard I am using throughout all VB-Projects and modules I am confident it is worth every effort. Eliminating errors can't be faster. The [README][5] in the corresponding [public GiHub repo][1] provides all the details for installation and usage. For experienced developers it will take not more than 30 minutes to finish the task. A true investment for future VBA development. 

## Disambiguation
| Term            | Meaning                                         |
| --------------- | ----------------------------------------------- |
|_Application&nbsp;Error_| An error which had been raised by an `err.Raise` statement distinguished from any system error like VB-Run-time or Database error. In order to avoid conflicts with system error numbers the `vbObjectError` is added to turn it into a negative number. The error service uses an AppErr function for this which turns the negative back into its original positive number for the display of the error. |
|_Entry&nbsp;Procedure_| For the error display this is one of the key issues for the assembling of the _path-to-the-error_.|
|_Error&nbsp;Source_   | An unambiguous name of the procedure which raised an error - by prefixing the procedure name with the module name.|
|_Common&nbsp;Components, Component| Term used for all my [Common VBA Components][2] of which the code is kept identical with all VB-Projects using them. A tough matter but successfully implemented and used. See:  |

## Features
I picture shall be better than many words:
![](../Assets/ErrMsgWithDebuggingOption.png)
![](/Assets/ErrMsgWithDebuggingOption.png)

- Welcome layout and design of the error message making use of the _Common VBA Message Service_
- [Path to the error][6]
- Optional _Debugging Button_ for resuming the error line<br>see [Straight to the error line][4]
- An optional _About the error_ section for _Application Errors_


## Comments
Any comment is very welcome. I apologize for the fact that it requires a GitHub account/login but this is the only appropriate way to ban spammers.

[1]:https://github.com/warbe-maker/Common-VBA-Error-Services
[2]:https://warbe-maker.github.io/vba/common/2021/02/19/Common-VBA-Components.html
[3]:https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services
[4]:https://warbe-maker.github.io/vba/common/error/handling/2022/02/16/Straight-to-the-Error-Line.html
[5]:https://github.com/warbe-maker/Common-VBA-Error-Services/blob/master/README.md
[6]:https://github.com/warbe-maker/Common-VBA-Error-Services/blob/master/README.md#the-path-to-the-error 