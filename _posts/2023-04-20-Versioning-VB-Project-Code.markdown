---
layout: post
title: Excel VB Component Management
subtitle: Keeping Common Components up-to-date in Vb-Projects using them
date:   2023-04-20
modified_date: 2023-04-20
categories: vba excel vb-project versioning
---
A single code line in combination with an imported Export-File is all what is required for a fully automated export of changed VBComponents.
<!--more-->

## Versioning approaches and tools
There are a number of [Alternatives](#alternatives-some) and so the final chosen versioning approach/tool depends on personal preferences. In general the alternatives can be split ito two components:
- Export-Files created when the code has changed [^2]
- a versioning tool (Git, CVS, SVN, etc.) which uses/versions them

I use [GitHub Desktop][3] as user interface for [GitHub][2]. The tool is free and only requires to clicks (_Commit_ and _Push_) to complete the versioning task by saving the changes into a GitHub _repository_ which may be private or public. The below focuses on my solution which I am using now for more than two years - continously improving it.

## Service only when applicable
Supposing that the development/maintenance of a VB-Project happens at different locations, code change versioning is useless for productive Workbooks [^1] &nbsp; Thus, the automated code export service is only provided when the Workbook resides in a defined location and productively used Workbooks are not bothered with it, i.e. not even recognize that the service is enabled for the Workbook.

[^1]: It may be an often practiced approach not to separate the productive use from the development task but it comes with the risk of a temporarily unusable Workbook. The resulting stress contradicts careful coding and testing. The risky approach is triggered by the fact that when a Workbook is used while its code is maintained means that the code changes have to be transferred to the productive Workbook which requires a proper [code synchronization service][4]. 

## Alternatives (some)
A closer look at the below alternatives may be worth some effort. However, the below is just a very first look at the subject with information taken directly from the solution provider.  

| Alternative | Short description (derived from source) |
|------------------|-------------------|
|[vbaDeveloper][5] | VbaDeveloper is an excel addin for easy version control of all your vba code. If you write VBA code in excel, all your files are stored in binary format. You can commit those, but a version control system cannot do much more than that with them. Merging code from different branches, reverting commits (other than the last one), or viewing differences between two commits is very troublesome for binary files. The VbaDeveloper Addin aims to solve this problem.|
|[VBASync][1]      | Maybe the only solution I know which does not touch the Workbook at all and not even uses Export-Files. However, this solution requires an additional manual step.|
| [XLTools][6]     | Powerful Excel add-in designed for business users.|

The aim of this post was to publish the solution I've provided and used for more than two years by continuously improving it. However, I shy away from adding it to the above list. The solution is based on a [CompMan.xlsb][7] (Component Management) Workbook which can be configured for auto-open or as an Add-in. When first opened after download it sets up its own environment (folders and files structure) and is immediately ready for servicing Workbooks which meet the [preconditions][10]. All information required is provided by the [README][8] in the [public GitHub repository][9].

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
