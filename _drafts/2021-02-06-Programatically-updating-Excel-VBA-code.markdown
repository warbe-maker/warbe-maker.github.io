---
layout: post
title: "Code updates with VBA"
date:   2021-02-05
categories: vba excel component management
---

Programmatically updating the code of a VB project is not straight forward like removing and re-importing a component.


## The challenge
1. There is no safe and stable way for a _VB-Project_ to uodate it's own code other than delegating this service to another _VB-Project_.
2. A component cannot be simply removed and replaced by importing an _Export File_ because the removal of a component is postponed by the system until the running process has ended. However, renaming and removing does the trick: The rename puts the component out of the way.
3. A service to update another _VB-Project's code is only always  available when needed when running as Addin - which is the birth of a  _Component-Management-Services_ Addin.

## Disambiguation of used terms
| Term | Meaning
|------|--------
|_Component_       | Generic _VB-Project_ term for a _Class Module_, a  _Data Module_, a _Standard Module_, or a _UserForm_  |
|_Common Component_| A _Component_ which is used by two or more VB-Projects |
| _Raw_,<br>_Raw-Component_ | The instance of a _Common Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated Workbook. |
| _Clone_,<br>_Clone-Component_ | The copy of a _Raw_ Component_ in any Workbook/_VP-Project_ using it |
|_Clone-Project_ | A Workbook/_VP project_ derived from a _Clone-Project_ |
|_Host_          | The Workbook/_VP-Project_ which hosts the _Raw-Component_ |
|_Raw-Project_   | A Workbook/_VP project_ of which all components are regarded _Raw_ Components_. A _Master Project_ is mainly a 'code-only-project' which does not have any other but static data |
|_Service_       | Generic term for any _Public Property_, _Public Sub_, or _Public Funtion_ of a _Component_ |
|_VB-Project_     | In the present case this term is used synonymously with Workbook |
| _Workbook-, or<br>VB-Project-Folder_ | A folder dedicated to a Workbook/VB-Project with all its Export Files and other project specific means. Such a folder is the equivalent of a Git-Repo-Clone (provided Git is used for the project's versioning which is recommendable |


## The _ExportChangedComponents_ service
The service is used with the _Workbook_Before_Save_ event. It compares the code of any component in a _VB-Project_ with its last _Export File_ and re-exports it when different. Usage of the service by _VB-Projects_ which host _Raw-Components_ is essential. The general usage of it for any BP-Project in a development status is appropriate as it is not only a code backup but also serves versioning. Any _Component_ indicated a _hosted Raw-Component is registered as such with its _Export File_ as the main property.<br>
The service also checks a _Clone-Component_ modified within the VB-Project using it a offers updating the _Raw-Component_ in order to make the modification permanent. Testing the modification will be a task performed with the raw hosting project.

The service has the following Syntax:

The service has the following named arguments:


## The _UpdateOutdatedClones_ service
The service is used with the _Workbook\_Open_ event. It checks each _Component_ for being known/registered as _Raw_  _hosted_ by another _VB-Project_. If yes, its code is compared with the _Raw's Export File and suggested for being updated if different.

## Installation

## Usage
### Usage of the _ExportChangedComponents_ service

### Usage of the _UpdateOutdatedClones_ service
