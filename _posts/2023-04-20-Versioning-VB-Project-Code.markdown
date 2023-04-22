---
layout: post
title: Excel VB Component Management
subtitle: Keeping Common Components up-to-date in Vb-Projects using them
date:   2023-04-20
modified_date: 2023-04-20
categories: vba excel vb-project versioning
---
My [Component Management Workbook][7] requires a single code line in combination with an imported Export-File for a fully automated export of changed VBComponents .
<!--more-->
## The fully automated Export service
When the [Component Management Workbook][7] is downloaded and opened it provides its own default environment of files and folders and is immediately ready for servicing Workbooks which meet the required [preconditions][10]. The corresponding [README][8] in the [public GitHub repository][9] provides all required information not only for the fully automated _Export of Changed Components_ service here in the focus. 

## Service only when applicable
Supposing that the development/maintenance of a VB-Project happens at different locations, code change versioning is useless for productive Workbooks [^1] &nbsp; Thus, the automated code export service is only provided when the Workbook resides in a defined location and productively used Workbooks are not bothered with it, i.e. not even recognize that the service is enabled for the Workbook.

## Versioning approaches and tools
There are a number of [Alternatives](#alternatives-some) and so the final chosen versioning tool depends on personal preferences. Most of the alternatives are based on Export-Files provided when the code has changed [^2] &nbsp; I use [GitHub Desktop][3] for Windows as user interface for [GitHub][2]. GitHub is free and only requires 2 clicks (_Commit_ and _Push_) to complete the versioning task which by saving the changes into a GitHub _repository_ which may be private or public. The below focuses on my solution which I am using now for more than two years - continously improving it.

[^1]: It may be an often practiced approach not to separate the productive use from the development task but it comes with the risk of an - at least  temporarily unusable Workbook. The resulting stress contradicts careful coding and testing. The risky approach is triggered by the fact that when a Workbook is used while its code is maintained means that the code changes have to be transferred to - synchronized with respectively - the productive Workbook which requires a proper [code synchronization service][4] also provided by the [Component Management Workbook][7]

## Some alternative versioning approaches/tools
The below alternatives are just a very first look with all provided information just taken directly from the solution provider. A more complete list with a closer look may be worth some effort however. 

| Alternative | Short description (derived from source) |
|------------------|-------------------|
|[vbaDeveloper][5] | VbaDeveloper is an excel addin for easy version control of all your vba code. If you write VBA code in excel, all your files are stored in binary format. You can commit those, but a version control system cannot do much more than that with them. Merging code from different branches, reverting commits (other than the last one), or viewing differences between two commits is very troublesome for binary files. The VbaDeveloper Addin aims to solve this problem.|
|[VBASync][1]      | Maybe the only solution I know which does not touch the Workbook at all and not even uses Export-Files. However, this solution requires an additional manual step.|
| [XLTools][6]     | Powerful Excel add-in designed for business users.|

[^2]: [VBASync][1] is the only solution I know which does not touch the Workbook at all and not even uses Export-Files. However, this solution requires an additional manual step.

[1]: https://github.com/chelh/VBASync
[2]: https://github.com
[3]: https://docs.github.com/en/desktop/installing-and-configuring-github-desktop/installing-and-authenticating-to-github-desktop/installing-github-desktop
[4]: https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services/blob/master/README.md?#synchronize-vb-project
[5]: https://github.com/hilkoc/vbaDeveloper
[6]: https://xltools.net/
[7]: https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services/blob/master/CompMan.xlsb?raw=true
[8]: https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services/blob/master/README.md
[9]: https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services
[10]: https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services/blob/master/README.md#enabling-the-services-serviced-or-not-serviced
