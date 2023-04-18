---
layout: post
title: "VB-Projects and Github"
date:   2020-09-30 09:11:20 +0200
categories: vba excel management
---
All my _Common Components_ hosting Workbooks [^1] are **GitHub** repositories. The positive experience over more than two years are worth sharing it.
<!--more-->

## Common advantages
1. **GitHub** is free of any charge no matter whether the VB-Project remains private or is made public
2. **GitHub** versions any changes, provided changed components are - preferably automated along with the `Workbook_BeforeSave` event - exported. This kind of automation may be done by means of the [Common Components Management][1] service _Export Changed Components_ [^2]. 
3. **GitHub** provides the means to make a code change in a branch first and once successfully tested merge the branch into the master.
4. When turned public other GitHub users may make use of the project or even contribute to it
5. An optional but rather obligatory README document provides the place for the description, installation and usage of the project
6. All outside the Vb-Project documents use Markdown which is easy to use and sufficiently capable for a user friendly appearance. A good example for this is Microsoft's VB Documentation.


## Specifics for VB-Projects
While for _Common Components_ **GitHub** is perfect, using it for productive VB-Projects seems to be an absolute nonsense. Excels tight coupling of data and code needs to be loosen in order to make it possible to maintain a productive Workbook's VB-Project - and the only way is, to synchronize the productive VB-Project with the maintained VB-Project. That's why the Common Components Management provides a VB-Project Synchronization service.

## Turning a dedicated VB-Project folder into a (cloned) GitHub repo
### Preconditions
1. A GitHub user account
2. Each Workbook lives in a dedicated parent folder [^3]

### Steps
1. Install and open [GitHub Desktop][3]
2. Use _File|Add local repository ..._ to turn a dedicated Workbook folder into a GitHub repo
3. Create a folder dedicated for Export-Files, e.g. name it _source_
4. Either export changed VBComponents manually (e.g. by MZTools or the native VB-Editor) or have this automated using the corresponding [Common Component Management service][1] service
5. Use **GitHub Desktop** to ***Commit*** made changes and ***Push*** them to your repo in github.com/<your github user id>/<your repo> 

[^1]: See the corresponding [README section][2] in the Common Components Management GitHub repo
[^2]: There are a couple of [other means][4] available which are worth being considered. An advantage of the [Common Components Management][1] is that it also provides a ***VB-Project Synchronization*** and an ***automated update for Common Components*** when there is a more up-to-date version available. 
[^3]: I myself keep Workbooks/VB-Projects in a dedicated Workbook folder already for a long time. The folder has a sub-folder which keeps the Export-Files of all changed VBComponents, whereby this export may be managed manually, e.g. using MZTools or fully automated using the [Common Component Management service][1]. This dedicated Export-folder also serves as a backup in case the VB-Project gets corrupted and is no longer readable - a true godsend in case. 

[1]:https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services
[2]:https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services/blob/master/README.md?#the-concept-of-hosted-common-components
[3]:https://docs.github.com/en/desktop/installing-and-configuring-github-desktop/installing-and-authenticating-to-github-desktop/installing-github-desktop
[4]:https://stackoverflow.com/questions/2996995/how-to-use-version-control-with-vba-code


