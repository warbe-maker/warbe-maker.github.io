---
layout: post
title: "Code updates with VBA"
date:   2021-02-05
categories: vba excel component management
---

Programmatically updating the code of a VB project is not straight forward like removing and re-importing a component.


## The hurdles
1. There is no safe and stable way for a VB project to uodate it's own code other than delegating this service to another VB project.
2. A component cannot be simply removed and replaced by importing an _Export File_ because the removal of a component is postponed by the system until the running process has ended. However, renaming and removing does the trick: The rename puts the component out of the way.
3. A service to update another VB projects code is only available when needed when running as _Component Management Services_ Addin.

## Disambiguation of used terms
| Term | Meaning
|------|--------
| _Component_ | Generic _VB-Project_ term for a _Class Module_ or _Data Module_, _Standard Module_, or _  |
_Common Component_ | A _Component_ which is shared ong two or more VB-Projects |
| _Raw_ | The instance of a _Common Component_ which is regarded the original. In other words the component in a Workbook/VB-project which is dedicated to its development,  maintenance and test. I.e. the Workbook which has the means to ensure the desired quality of the services the component provides |
| _Clone_ | The copy of a _Raw_ component in any Workbook/VP-Project using it |
|_Host_ | The Workbook/VP-Project which hosts the _Raw_ component |
|_Service_| A generic term for any _Public_ Property, Sub, or Funtion of  _Component_ |
| _Template Project_ | A Workbook/VP-Project of which all components are regarded _Raw_ component. A _Template Project_ is a "code-only-project and does not have any data other than static data |
| _Clone Project_ | A Workbook/VP-Project derived from a _TemplateProject_ |
| _Workbook Folder_ | A folder dedicated to a Workbook/VB-Project with all its Export Files and other project specific means. Such a folder is the equivalent of a Git-Repo-Clone (provided Git is used for the project's versioning which is recommendable |## The _UpdateOutdatedClones_ service
A component which is developed, maintained and tested in another VB project can be called the _Raw Component_ component. The copy of this component used by another VB project can be called a _Clone Component_.
The service checks for any clone of which the raw has changed and replaces it.

## The _ExportChangedComponents_ service
The service is evoked with the Workbook_Before_Save event, compares the code of all components in a VB project with its last _Export File_, and exports them when different. VB projects which host _Raw Components_ as well as VB projects using them use the service. Raws hosting VB projects register these components as available for other VB projects using them. 

## Installation

## Usage