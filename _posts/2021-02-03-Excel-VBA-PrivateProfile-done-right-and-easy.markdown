---
layout: post
title: "Excel VBA PrivateProfile done right - and easy"
date:   2020-09-30 09:11:20 +0200
categories: VBA Office Excel
---

## Introduction
_PrivateProfile_ is the term used for information in a file organized as<br>[section]<br>\<valuen-ame>=\<value><br>structure, typically for config- or ini- files. Word provides for example [System.PrivateProfileString][4] with a perfect syntax. Excel unfortunately offers only things like [GetPrivateProfileString][3] with a much less comfortable syntax. The Standard Module _mFile_ provides 'Word-like' services which as well mainly deal with the arguments: file, section, value-name and value.

## The _mFile_ PrivateProfile services

### The _Value_ service
Syntax read: `value = mFile.Value(file, section, value-name)`<br>
Syntax write: `mFile.Value(file, section, value-name) = value`

### The _ValueExists_ service
Syntax: `If mFile.ValueExists(file[, section], value) Then`

### The _NameExists_ service
Syntax: `If mFile.NameExists(file[, section], value-name) Then`

### The _SectionExists_ service
Syntax: `If mFile.SectionExists(file, section) Then`

### The _SectionsCopy_ service
Syntax: `mFile.SectionsCopy source, target, sections`

## The services have the following named arguments

| Argument      | Description | Services |
| ------------- | ----------- | -------- |
| pp_file       | String expression, obligatory, specifies the full name of the _PrivateProfile file, automatically created with the first write if a named value.| All |
| pp_sections   | Variant, optional, defaults to 'all sections in file' when omitted. Section names may be provided as a comma delimited string, or a Dictionary or Collection of name items.  | SectionsCopy<br>SectionsRemove<br>ValueExists<br>ValueNameExists|
| pp_replace    | Optional, boolean, defaults to false (i.e. the copied section is merged in the target file. | SectionsCopy |
| pp_section    | Obligatory, String expression, identifies the section| NameRemove<br>SectionExists<br>SectionRemove<br>Value<br> |
| pp_source     | String expression, obligatory, specifies the full name of the source _PrivateProfile_ file | SectionsCopy |
| pp_target     | String expression, obligatory, specifies the full name of the target _PrivateProfile_ file | SectionsCopy |
| pp_value_name | | NameRemove<br>Value |
| pp_value      | Variant expression, the value written to the _PrivateProfile_ file | Value<br>ValueExists |


## Installation
- Download and import [mFile.bas][1]
- Download and import [mDct.bas][2]

## Usage
The services may best be used in a Standard Module dedicated to the file used for the required application specific values, whereby each value preferably is implemented as a Property. The following example provides a read service for a property called _RootFolder_ in a module called _mCfg_.
```VB
Option Explicit

Private Const CFG_SECTION = "Basic"
Private Const VAL_NAME_ROOT_FOLDER = "RootFolder"

Private Property Get CfgFile() As String
    CfgFile = ...... ' specifying the ProvateProfile file's full name
End Property

Private Property Get RootFolder() As String
    RootFolder = mFile.Value(CfgFile, CFG_SECTION, VAL_NAME_ROOT_FOLDER)
End Property

```
This service will be used subsequently in the project:<br>
`sRoot = mCfg.RootFolder`

the matter of all the above is a pretty nice example how the implementation of a couple of modules each providing a service on a higher abstraction layer. I.e. each module provides a service while hiding the used technical means.  

[1]: https://gitcdn.link/repo/warbe-maker/Common-VBA-File-Services/master/mFile.bas
[2]: https://gitcdn.link/repo/warbe-maker/Common-VBA-Dirctory-Services/master/mDct.bas
[3]: https://docs.microsoft.com/de-de/windows/win32/api/winbase/nf-winbase-getprivateprofilestring?redirectedfrom=MSDN
[4]: https://docs.microsoft.com/de-de/office/vba/api/word.system.privateprofilestring