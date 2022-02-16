---
layout: post
title: Common VBA Execution Trace
date: 2020-11-14
modified_date: 2022-02-16
categories: vba common
---
Monitoring which code component has been executed and how much time (highest precision!) it took.
<!--more-->

### Service
The [Common-VBA-Execution-Trace-Service][1] provides the means to trace the execution of any procedure or code snippet, writing the trace result to a log file which defaults to ThisWorkbook's parent folder named _ExecTrace.log_.

The below example resulted from the module's regression test:
![](../Assets/ExecutionTrace.png)<br>


### Installation
Download [_mTrc.bas_][1] and import it into your VB-Project.  

See the README in [Common-VBA-Execution-Trace-Service][1] for detailed information about how to use it.

[1]:https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mTrc.bas
[3]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/mMsg.bas
[4]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frm
[5]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Execution-Trace-Service/master/source/fMsg.frx

