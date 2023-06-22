---
layout: post
title: Common VBA Stack and Queue Services
subtitle: An alternative for the VBA MsgBox with less constraints, more options, and a better display
date:          2023-02-07
modified_date: 2023-02-07
categories:    vba common
---
Queue and stack services appear pretty trivial at first. However, a closer look, and it is worth a comprehensive (code and forget) implementation. This post focuses on Queue services.
<!--more-->

## Preface
 _[Common VBA Stack and Queue Services README][1]_ 
With four lines of code the implementation of a queue (an _Enqueue_ and a _Dequeue_ service) based on a Collection appears - and is - trivial:
```vb
    Dim MyQueue As New Collection
    MyQueue.Add ...     ' enqueue
    v = MyQueue(1)      ' get (non-object) queue item
  ' Set v = My queue(1) ' get (object) queue item
    MyQueue.Remove 1    ' dequeue
```
Looks like not worth many more effort. However a true comprehensive solution should address some more aspects of the matter and provide all potentially useful services. Considered should be:
- a mixture of items which represent an object and other type of variables all together in one queue
- not only the first but also a specific item should be able to be _de-queued_ (in case not unique in the queue optionally by its position) <span id="a1">[¹](#1)</span>

Given these and some more aspects a much more comprehensive (what I like to call 'code and forget') solution is worth being implemented.
The below private procedures are an extract from the StandardModule ([mQ.bas][5] for download) and are also used in the Class module ([clsQ.cls][4] for download):
```
```
The inline documentation in the above code should suffice. For more see the [GitHub repository][1] which is just a Workbook (QaS.xlsb for download) in its dedicated folder. Alternatively just see the _[README][2]_ which provides a Queue and a Stack, both as Common Modules, plus a (fully traced) Regression-Test. 


## Comments
Comments of any kind are more than welcome. I apologize for the fact that it requires a GitHub account/login but this is the only appropriate way avoiding spams.

<br><span id="1">¹</span> In contrast to a stack where only push and pop matters, a queue should provided a service to de-queue a specific and not only the first item.[⏎](#a1)<br>

[1]:https://github.com/warbe-maker/VBA-Queue-and-Stack
[2]:https://github.com/warbe-maker/VBA-Queue-and-Stack/blob/master/README.md
[4]:https://gitcdn.link/cdn/warbe-maker/VBA-Queue-and-Stack/master/source/clsQ.cls
[5]:https://gitcdn.link/cdn/warbe-maker/VBA-Queue-and-Stack/master/source/mQ.bas
[6]:https://gitcdn.link/cdn/warbe-maker/VBA-Queue-and-Stack/master/source/clsStk.cls
[7]:https://gitcdn.link/cdn/warbe-maker/VBA-Queue-and-Stack/master/source/mStk.bas