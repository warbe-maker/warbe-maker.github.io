---
layout: post
title: Common VBA Message Service
subtitle: An alternative for the VBA MsgBox with less constraints, more options, and a better display
date:          2020-11-17
modified_date: 2023-04-18
categories:    vba common
---
An alternative to the VBA.MsgBox. Amongst others it has the advantages of an almost unlimited message size.
<!--more-->

## Preface
The VBA.MsgBox is fine for ad-hoc messages but one has to live with its limits. There's hardly any way to provide appealing messages and the size is very limited. I've got the time and motivation to implemented an alternative without all the limits. The _[Common VBA Message Service README][1]_ public GitHub repo provides all information on how to install and use it. 

## Features
![](../Assets/CommMsgServiceDemo.png)
![](/Assets/CommMsgServiceDemo.png)
- 4 message sections, each with an optional label, both free in color, font (proportional or mono-spaced), font size, bold, italic, underline, etc.
- Width and height adjusted up to a specifiable maximum (within a min/max range)
- Minimum width specifiable
- Up to 7 x 7 reply buttons with any caption text plus all the VBA.MsgBox button values
- Vertical and horizontal scroll-bars when maximum width/height is exceeded (proportional message sections adapt, mono-spaced sections determine their width by the longest line)
- An optional mode-less display allows the use of any number of message displays in parallel (i.e. instances of the message UserForm).

## Examples
The following display examples show the great flexibility of the _[Common VBA Message Services][1]_
### Display of an Error Message 
![](../Assets/ErrMsgWithDebuggingOption.png)
![](/Assets/ErrMsgWithDebuggingOption.png)<br>
<small>Note: The path-to-the-error is a service provided by the _Common VBA Error Services_! The example is shown because the error service uses the message service the error message display.</small>

### Display of an Execution trace
![](../Assets/ExecutionTrace.png)
![](/Assets/ExecutionTrace.png)

### Display of a Process monitor
![](../Assets/DemoMonitorService.gif)
![](/Assets/DemoMonitorService.gif)

### Display of several Process monitoring instances
![](../Assets/DemoMonitorServiceInstances.gif)
![](/Assets/DemoMonitorServiceInstances.gif)

### Mode-less displayed Buttons with an Application.Run service
Displaying a message mode-less is rather useless when the buttons do not call services via `Application.Run`. The _[Common VBA Message Service][1]_ offers a service to provide displayed buttons with `Application.Run` information executed when clicked - provided the message is displayed mode-less. This service substantially extends a displayed messages functionality. It may be used to trigger actions/services provided by any open Workbook, i.e. not only the Workbook which displayed the message. [^1]


## Comments
Comments of any kind are more than welcome. I apologize for the fact that it requires a GitHup account/login but this is appropriate for keeping away spammers.

[^1]: A _CommandButton_ on a Worksheet can only be assigned a macro without arguments! In contrast, Buttons displayed by the [Common VBA Message Services][1] may be provided with any number of arguments.

[1]:https://github.com/warbe-maker/Common-VBA-Message-Service/blob/master/README.md

