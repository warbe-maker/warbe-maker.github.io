---
layout: post
title: Excel VB-Project development towards professionalism
subtitle: 
date:       2023-06-12
modified:   2023-06-26
categories: vba excel component management common components
---
It's the heritage of may active professional IT life time and the result of my retirement life: Means substantially supporting - mainly Excel - VB-Project development: A bunch of ***Common Components*** to increase efficiency of development and a ***Component Management*** which keeps them up-to-date in all using VB-Projects.  
> Those possibly overwhelmed likely will benefit most. It is able to sustainable change their Excel VP-Project development life from tinkered to professional.

### Attraction for _Common Components_
Disambiguation: Procedure (Function, Methods, code snippets, etc.) with useful properties for more than just one VB-Project, collected in a _Standard Module_ or a _Class Module_.
Such components are ideal for:
- The use of the **[Common VBA Execution Trace Service][1]** which supports the identification of performance flaws and last but not least the  documentation (extent of) a [regression test](#regression-testing)
- The use of the **[Common VBA Error Services][2]** which supports debugging and regression testing
- and last but not least management means which keeps them up-to-date in all VB-Projects using them.

### Component Management (Excel only)
Not only for the management of _Common Components_ but also for maintaining a versioning basis.
- An **Update outdated Common Components** service keeps used **Common Components** up-to-date by updating them when outdated and the Workbook is opened for development/maintenance (in a dedicated folder!)
- An **Export of changed VB-Components** service supports versioning - in combination with versioning means like GitHub (what I prefer) by exporting any changed component
- A **Synchronization VB-Projects** service supports development and maintenance without the unavoidable downtime for the productive Workbook otherwise  

See the _[Common VBA Components Management][3] for more information.

### Regression testing
It is the nature and the advantage of a _Common Component_ that it allows a dedicated development environment which enables regression-testing. Something which hardly can be achieved when a component is just declared as such in a non-dedicated Workbook. Regression testing is supported by the **[Common VBA Error Services][2]** and the **[Common VBA Execution Trace Service][1]**.  
See a complete example with the **[Common VBA Basic Services][8]**.
The [RegressionTest.ExecTrace.log][4] file shows the result of the execution trace. The [mBasicTest.bas][5] is the test component which produces the execution trace result. See how to [fork][6] the public [GitHub repo][8] to make it all available as your own personal GitHub repo which you the may clone to your computer [by means of [GitHub DeskTop][7]) of which makes the whole thing available in a dedicated folder on your computer.[^1]

### My _Common Components_
All my _Common Component_ development aim for working completely autonomous, i.e. they are not requiring other module. The use of other _Common Components_ remains optional. The first three of them are mentioned in this blog. Others are also available as public GitHub repos.

[^1]:Having a look on what GitHub provides - also for Excel VB-Project development will open a complete new world for those never heard if it - or having it ignored for a long time like me.

[1]:https://github.com/warbe-maker/VBA-Trace
[2]:https://github.com/warbe-maker/VBA-Error
[3]:https://github.com/warbe-maker/VBA-Component-Management
[4]:https://github.com/warbe-maker/VBA-Basics/blob/master/RegressionTest.ExecTrace.log
[5]:https://github.com/warbe-maker/VBA-Basics/blob/master/source/mBasicTest.bas
[6]:https://www.freecodecamp.org/news/how-to-fork-a-github-repository/#:~:text=Forking%20a%20repository%20is%20the%20matter%20of%20clicking,forked%20repository%20gets%20created%20under%20your%20GitHub%20account.
[7]:https://desktop.github.com
[8]:https://github.com/warbe-maker/VBA-Basics