---
layout: post
title: "VB-Projects and GitHub"
date:          2020-09-30
modified_date: 2023-07-13
categories: vba excel management
---
All my VB-Projects, including all my _Common Components_ hosting Workbooks [^1] are ***[GitHub][5]*** repositories since several years and I found the (positive) experience  worth sharing it.
<!--more-->

## Common advantages
1. ***[GitHub][5]*** is free of any charge no matter whether the VB-Project remains private or is made public
2. ***[GitHub][5]*** versions any changes, provided changed components are - preferably automated along with the `Workbook_BeforeSave` event - exported. This kind of automation may be done by means of the [Common Components Management][1] service _Export Changed Components_ [^2]. 
3. ***[GitHub][5]*** provides the means to make a code change in a branch first and once successfully tested merge the branch into the master.
4. When turned public other GitHub users may make use of the project or even contribute to it
5. An optional but rather obligatory README document provides the place for the description, installation and usage of the project
6. All outside the Vb-Project documents use Markdown which is easy to use and sufficiently capable for a user friendly appearance. A good example for this is Microsoft's VB Documentation.

## Specifics for "productive" Workbooks/VB-Projects
Considering ***[GitHub][5]*** for _Common Components_ (specifically when hosted in a dedicated Workbook) is definitely an option since those kind of components usually do not imply data. Considering ***[GitHub][5]*** for a ++productive++ Workbook/VB-Project appears much less obvious if not an absolute nonsense at first glance. What makes it considerable is a [_VB-Project Synchronization Service_][6] like the one provided by my [_Common Components Management_][1] Workbook/Addin: While the productive Workbook remains in place, its VB-Project is maintained/modified in a copy and the productive VB-Project is finally  synchronized with the copy, thereby minimizing the productive Workbook's downtime. Less development stress = better (tested!) result.

## Turning a dedicated VB-Project folder into a (cloned) GitHub repo
### Preconditions
1. A GitHub user account
2. The Workbook "lives" in a dedicated parent folder [^3]

### Steps
1. Install and open [_GitHub Desktop_][3]
2. Use _File|Add local repository ..._ to turn a dedicated Workbook folder into a GitHub repo
3. Create a folder dedicated for Export-Files, e.g. name it _source_
4. Either export changed VBComponents manually (e.g. by [_MZTools_][7] or the native VB-Editor) or have this automated using the corresponding [Common Component Management service][1] service
5. Use [_GitHub Desktop_][3] to ***Commit*** made changes and ***Push*** them to your GitHub repo (https://github.com/<your-github-user-id>/<your-repo> 

[^1]: See the corresponding [README section][2] in the Common Components Management GitHub repo
[^2]: There are a couple of [other means][4] available which are worth being considered. An advantage of the [Common Components Management][1] is that it also provides a ***VB-Project Synchronization*** and an ***automated update for Common Components*** when there is a more up-to-date version available. 
[^3]: I myself keep Workbooks/VB-Projects in a dedicated Workbook folder already for a long time. The folder has a sub-folder which keeps the Export-Files of all changed VB-Components, whereby this export is automatically managed by a [_Common Component Management Service][1]. This dedicated Export-folder not only serves as a backup in case the VB-Project gets corrupted (a true godsend in case) but is also the source for a versioning provided by [_GitHub Desktop_][3]. 

[1]:https://github.com/warbe-maker/VBA-Component-Management
[2]:https://github.com/warbe-maker/VBA-Component-Management/blob/master/README.md?#the-concept-of-hosted-common-components
[3]:https://docs.github.com/en/desktop/installing-and-configuring-github-desktop/installing-and-authenticating-to-github-desktop/installing-github-desktop?platform=windows
[4]:https://stackoverflow.com/questions/2996995/how-to-use-version-control-with-vba-code
[5]:https://github.com
[6]:https://github.com/warbe-maker/VBA-Component-Management#synchronize-vb-project
[7]:https://www.mztools.com