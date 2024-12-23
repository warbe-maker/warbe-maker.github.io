---
layout: post
title: An easy way to Private Profile files in VB-Projects
date:          2021-02-03
modified_date: 2024-12-23
categories:    vba Office Excel
---
<!--more-->
Private Profile files made easy and transparent.

## Appropriateness
Microsoft suggests to use the Registry instead of Private Profile files. A _Private Profile file_ may still worth being considered because:
- Anything written to the registry by a VB-Project is personal and not available for other users but a Workbook/VB-Project may be used by multiple individuals which should make use of the same configuration.
- When the VB-Project has been developed by a professional contractor the configuration should no be an integrated part of the implementation. Any modified VB-Project will then use the client's configuration. And last but not least may the contractor's configuration differ from the client's.
- A "common" Workbook/VB-Project may be used by different clients, each with an individual configurations.

## Implementation
There are a couple of library calls available for the maintenance of a _Private Profile file_. Their interface is somewhat unusual and has some limitations however. An easy going implementation (at least from my point of view) would be a Class Module with a _Value_ Get and a _Value_ Let Property with the three parameters:
- _file\_full\_name_
- _section\_name_
- _value\_name_
- _value\_default_ (used for _Value_ Get only, value returned as default when not available)

For comprehensiveness there are a couple of other methods and properties thinkable.

## An example
The autonomous Class Module _[clsPrivProf][1]_ may be downloaded (see [how to](#download-from-githbub) when unfamiliar with GitHub) and imported from the public GitHub repository or the code may directly be copied into a Class Module named _clsProvProf_. The module runs without any specific library calls and provides the following advantages:
- An unlimited string length (no hassle with buffer size)
- Sections and Value-Names are maintained in ascending order
- Optional comment lines (file header, file footer, section comment, and value comment)
- Sections are separated by an empty line
- Easy to use methods and properties since all use the same set of three parameters.

## Usage
- Value write: `Value(<value-name>[, <section-name>][, file-full_name]) = "any"`
- Value read: `result = Value(<value-name>[, <section-name>][, <file-full-name>]) `
  - _value\_name_ (obligatory)
  - _section\_name_ (optional)
  - _file\_full\_name_ (optional)
Note: Both, the section-name and the file-name are optional. Specifically the _file\_full\_name_ will usually be omitted since specified once when the class instance is established.

For further information see the corresponding [README][2] and a supplementing [SpecsAndUse][3] document.

## Download from GitHbub
For those unfamiliar or discomfort with GitHub: The link displays the module's Export file which allows to download it or copy the code  
![downloaded](../Assets/GitHubDownload.png)  
![downloaded](/Assets/GitHubDownload.png)  or imported

[1]: https://github.com/warbe-maker/Common-VBA-Private-Profile-Services/blob/main/CompMan/source/clsPrivProf.cls
[2]: https://github.com/warbe-maker/Common-VBA-Private-Profile-Services
[3]: https://github.com/warbe-maker/Common-VBA-Private-Profile-Services/SpecsAndUse.md
