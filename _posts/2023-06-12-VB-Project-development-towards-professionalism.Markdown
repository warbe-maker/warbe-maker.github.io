---
layout: post
title: Excel VB-Project development towards professionalism
subtitle: 
date:       2023-06-12
modified:   2023-07-13
categories: vba excel component management common components
---
It's the heritage of may active professional IT life time, a passion for Excel VB-Projects, and the result of my retirement life: A bunch of ***Common Components*** to increase the performance and quality of Excel VB-Project development, a ***Component Management*** Addin which keeps ***Common Components** up-to-date in all VB-Projects using them and last but not least a _[professional ***Error Handling***](#attraction-for-a-professional-error-handling-with-a-comprehensive-error-message-display)_ with a comprehensive ***Error Message*** displayed.  
> Those feeling overwhelmed will likely benefit most. Both, _Common Components_ and _Component Management_ is able to sustainable changes Excel VP-Project development life from tinkered to professional.

### Attraction for _Common Components_
#### Disambiguation
A _Common Component_ in this context is a _Standard Module_ or _Class Module_) with procedures regarding a certain context (`Sub`, `Function`, `Property`) designed for being used in any VB-Project.
#### Benefit, potential
_Common Componets_ are a true efficacy boost for the VB-Project development not only when used/imported as a whole but also as a rich source of procedures for being copied into other modules. To unfold its full potential a _Common Component_ needs to be thorowly and extensively tested (see [regression testing](#regression-testing) which is possible best in a dedicated Workbook hosting it (a non dedicated Workbook will hardly ever provide a complete, comprehensive test environment which also supports regression testing. Last but not least, a _Common Component_ should be kept identical in all Workbooks/VB-Projects using it.
#### Component Management (Excel only)
Not only for the management of _Common Components_ but also for maintaining a basis for versioning:
- An **Update outdated Common Components** service keeps used **Common Components** up-to-date by updating them when outdated and the Workbook is opened for development/maintenance (in a dedicated folder!)
- An **Export of changed VB-Components** service supports versioning - in combination with versioning means like GitHub (what I prefer) by exporting any changed component
- A **Synchronization VB-Projects** service supports development and maintenance without the unavoidable downtime for the productive Workbook otherwise.  

See the _[Common VBA Components Management][3] for more information.

#### My _Common Components_
All my _Common Component_ development aim for working completely autonomous, i.e. they are not requiring any other module, the use of other _Common Components_ remains optional respectively. The most frequently used three of them are mentioned in this blog. Others are also available as public GitHub repos.

### Attraction for a professional _Error Handling_ with a comprehensive _Error Message_
An ideal error message in this respect will look as follows&nbsp;[^1]:  
![](../Assets/DemoAppErrByErhErrMsg.png)
It provides the best possible support for debugging by means of a comprehensive error message and the means to resume the code line which raised/caused the error. Furthermore the _Error Handling_ should provide a _[regression testing](#regression-testing)_ support, i.e. the ability to suppress the display of _asserted_ error (errors which are explicitly tested and thus expected).

### Attraction for _Regression Testing_
Regression testing could be considered the Holy Grail of IT quality management. For VB-Projects as a whole it may be not possible or too cumbersome. _Common Components_ in contrast are perfect candidates for it (I've provided one for each of mine), specifically when hosted in a dedicated Workbook which allows a dedicated development environment. Something which hardly can be achieved when a component is just declared as _common_ in a non-dedicated Workbook&nbsp;[^2]. Regression testing is supported by the **[Common VBA Error Services][2]** and the **[Common VBA Execution Trace Service][1]**.  
See a complete example with the **[Common VBA Basic Services][8]**.
The [RegressionTest.ExecTrace.log][4] file shows the result of the execution trace. The [mBasicTest.bas][5] is the test component which produces the execution trace result. See how to [fork][6] the public [GitHub repo][8] to make it all available as your own personal GitHub repo which you the may clone to your computer [by means of [GitHub DeskTop][7]) of which makes the whole thing available in a dedicated folder on your computer.[^3]


[^1]: The **[Common VBA Error Services][2]** provides a best possible support for debugging and regression testing.
[^2]:The _Component Management_ supports both however, even the ad-hoc modification of _Common Component_ in its using VB-Project.
[^3]:Having a look on what GitHub provides - also for Excel VB-Project development will open a complete new world for those never heard if it - or having it ignored for a long time like me.


[1]:https://github.com/warbe-maker/VBA-Trace
[2]:https://github.com/warbe-maker/VBA-Error
[3]:https://github.com/warbe-maker/VBA-Component-Management
[4]:https://github.com/warbe-maker/VBA-Basics/blob/master/RegressionTest.ExecTrace.log
[5]:https://github.com/warbe-maker/VBA-Basics/blob/master/source/mBasicTest.bas
[6]:https://www.freecodecamp.org/news/how-to-fork-a-github-repository/#:~:text=Forking%20a%20repository%20is%20the%20matter%20of%20clicking,forked%20repository%20gets%20created%20under%20your%20GitHub%20account.
[7]:https://desktop.github.com
[8]:https://github.com/warbe-maker/VBA-Basics