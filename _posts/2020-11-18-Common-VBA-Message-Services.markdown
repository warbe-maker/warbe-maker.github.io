---
layout: post
title: Common VBA Message Service
subtitle: An alternative for the VBA MsgBox with less constraints, more options, and a better display
date:          2020-11-17
modified_date: 2022-02-18
categories:    vba common
---
An alternative to the VBA.MsgBox of which the limits hardly will ever been reached.
<!--more-->
## Preface
The quick-and-dirty VBA.MsgBox is fine for many things but it has its limits. I found it almost impossible to display anything a little bit nicer designed. And because I've got the time and the ambition I implemented one without all the limits. The _[Common VBA Message Service README][1]_ public GitHub repo provides all information on how to install and use it. 

## Features
![](../Assets/CommMsgServiceDemo.png)
![](/Assets/CommMsgServiceDemo.png)
- 4 message sections, each with an optional label, both free in color, font (proportional or mono-spaced), font size, bold, italic, underline, etc.
- Width and height adjusted up to a specifiable maximum (within a min/max range)
- Minimum width specifiable
- Up to 7 x 7 reply buttons with any caption text plus all the VBA.MsgBox button values
- Vertical and horizontal scroll-bars when maximum width/height is exeeded (proportional message sections adapt, mono-spaced sections determine their widht by the longest line)
- An optional mode-less display allows the use of any number of message displays in parallel.

## Usage examples
### Error Message
![](../Assets/ErrMsgWithDebuggingOption.png)
![](/Assets/ErrMsgWithDebuggingOption.png)

### Display of execution trace
![](../Assets/ExecutionTrace.png)
![](/Assets/ExecutionTrace.png)

### Process monitoring
See the [_Monitor_ service demonstration][2]

### Process monitoring instances
See the [_Monitor_ service instances demonstration][3]. It makes use of the _MsgInstance_ service to position the message windows on the display. 


## Comments
Comments of any kind are more than welcome. I apologize for the fact that it requires a GitHup account/login but this is appropriate for keeping away spammers.


[1]:https://github.com/warbe-maker/Common-VBA-Message-Service/blob/master/README.md
[2]:https://github.com/warbe-maker/Common-VBA-Message-Service/blob/master/README.md#monitor-service-demonstration
[3]:https://github.com/warbe-maker/Common-VBA-Message-Service/blob/master/README.md#demo-of-the-monitor-service-using-the-msginstance-service
