---
layout: post
title: Programmatically modifying the code of an  Excel VB-Project
date:   2021-02-19
categories: vba excel code component management
---

## Introduction
Programmatically updating the code of a _VB-Project_ is not straight forward like removing and re-importing a component. Synchronizing all the code between two _VB-Projects_ is an even more ambitious service. Re-started several times I've finally ended up with a set of satisfyingly stable services provided via an Addin-Workbook.


## Challenges
1. There is no safe and stable way to programmatically modify the code of a _VB-Project_  other than delegating this service to another _VB-Project_.
2. A component cannot be simply removed and replaced by importing an _Export File_ because the removal of a component is postponed by the system until the running process has ended. However, renaming and removing does the trick because the rename puts the component out of the way for the import.
3. An update service may be available
   - through an open Workbook via Application.Run
   - through an Addin-Workbook which may automatically be opened by another Workbook referencing it.
4. _Document Modules_ come with extra challenges
   - They code can only be changed by transferring the code from an _Export-File_ line by line
   - While renaming the _Workbook Document Module_ is pretty straight forward the _Worksheet Document Module_ has two names, the sheets _Name_ and the sheets _CodeName_. When both are changed the assignment becomes a challenge, uncertain, or even impossible depending on whether the number of sheets is still equal or different
   - _Worksheets_ may have new or outdated _Controls_ and their properties may have changed, all leading to additional challenges not considered in the first place
   - _Workheets_ may - and often will - come with range names which can only be updated in concert with a sheet's design change which will remains a manual task


Conclusion: Updating _VBA Components_ developed, maintained (and appropriately tested!) in one _VB-Project_ and used by others plus the synchronization of a productive Workbook with a temporary development copy is the aim of my _CompMan_ Workbook/VB-Project.

## Disambiguation
The terms below are used in all posts regarding this matter and in the _[Excel-VB-Components-Management][2]_ VB-Project.


| Term             | Meaning                  |
|------------------|------------------------- |
|_Component_       | Generic _VB-Project_ term for a _Class Module_, a  _Data Module_, a _Standard Module_, or a _UserForm_  |
|_Common Component_| A _Component_ which is used by two or more VB-Projects |
| _Raw_,<br>_Raw-Component_ | The instance of a _Common Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw-Host_ Workbook. |
| _Clone_,<br>_Clone-Component_,<br>_Raw-Clone_ | The copy of a _Raw- Component_ in a _VP-Project_ using it |
|_VB-Clone-Project_ | A _VP-Project_ derived from a _Raw-Project_ |
|_Procedure_     | Any - Public or Private _Property_, _Sub_, or _Funtion_ of a _Component_. See also _Service_.
|_Raw-Host_.     | The Workbook/_VP-Project_ which hosts the _Raw-Component_ |
|_VB-Raw-Project_ | The copy of a _VB-Clone-Project_ temporarily used for the development and test of the _VB-Clone-Project_. When the development had finished, the source for the code synchronization.|
|_Service_       | Generic term for any _Public Property_, _Public Sub_, or _Public Funtion_ of a _Component_ |
|_VB-Project_     | In the present case this term is used synonymously with Workbook |
| _Workbook-_, or<br>_VB-Project-Folder_ | A folder dedicated to a Workbook/VB-Project with all its Export Files and other project specific means. Such a folder is the equivalent of a Git-Repo-Clone (provided Git is used for the project's versioning which is recommendable |


## Services
### _ExportChangedComponents_ service
Used with the _Workbook_Before_Save_ event it compares the code of any component in a _VB-Project_ with its last _Export File_ and re-exports it when different. The service is essential for _VB-Projects_ which host _Raw-Components_ in order to get them registered as available for other _VB-Projects_. Usage by any _VB-Project_ in a development status is appropriate as it is not only a code backup but also perfectly serves versioning - even when using [GitHub][]. Any _Component_ indicated a _hosted Raw-Component is registered as such with its _Export File_ as the main property.<br>
The service also checks a _Clone-Component_ modified within the VB-Project using it a offers updating the _Raw-Component_ in order to make the modification permanent. Testing the modification will be a task performed with the raw hosting project.

For the service's syntax and named arguments see [Usage of the _ExportChangedComponents_ service](#usage-of-the-exportchangedcomponents-service).

### _UpdateRawClones_ service
The service is used with the _Workbook\_Open_ event. It checks each _Component_ for being known/registered as _Raw_  _hosted_ by another _VB-Project_. If yes, its code is compared with the _Raw's Export File and suggested for being updated if different.

For the service's syntax and named arguments see [Usage of the  _UpdateRawClones_ service](#usage-of-the-updaterawclones-service).

### _SynVbProjects_ service
Under construction

## Installation
### _CompMan_ Add-in
1. Download and open [CompManDev.xlsb][1]
2. Follow the instructions to identify a location for the Add-in - preferably a dedicated folder like ../CompMan/Add-in. The folder will hold the following files:
   - CompMan.cfg    ' the basic configuration
   - CompMan.xlam   ' the Add-in
   - HostedRaws.dat ' the specified raws hosted in any Workbook
   - RawHost.dat    ' the Workbooks which claim raws hosted
   
3. Follow the instructions to identify a 'serviced root'
4. Use the built-in Command button to run the _Renew_ service. It will:
   - ask to confirm or change the basic configuration
   - initially setup or subsequently renew the CompMan Add-in by saving a copy  of the development instance as Add-in (mind the fact that this is a multi-step process which may take some seconds)

Once the Add-in is established it will automatically be loaded with the first Workbook opened having it referenced. See the Usage below for further required preconditions.

### Installation for Workbooks/VB-Projects hosting raws or using raw clones
1. Copy the following into the Workbook component
```vb
Option Explicit

Private Const HOSTED_RAWS = ""

Private Sub Workbook_Open()
#If CompMan Then
    mCompMan.UpdateRawClones uc_wb:=ThisWorkbook _
                           , uc_hosted:=HOSTED_RAWS
#End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
#If CompMan Then
    mCompMan.ExportChangedComponents ec_wb:=ThisWorkbook _
                                   , ec_hosted:=HOSTED_RAWS
#End If
End Sub
```
2. For a Workbook which hosts _Raw-Components_ specify them in the HOSTED_RAWS constant delimited with commas.

> ++**Be aware:**++ When the update service is initiated from within the Workbook_Open event, the Workbook component of this  VB-Project is the only code which cannot be modified programmatically. When the update service is initiated manually in the immediate window, even the Workbook component's code may be modified. Unfortunately there is no way for the service to check these condition and thus the Workbook component is exempted from any programmatic code modification. This constraint can only be handled by all open code in a dedicated Standard module.

## Usage
### Preconditions
Every service will be denied unless the following preconditions are met:
1. The basic configuration is complete and valid
3. The serviced Workbook resides in a subfolder of the configured _ServicedRootFolder_. When copied to a Location 'outside' the services will be denied even when all other preconditions are met.
4. The serviced Workbook is the only Workbook in its parent folder
5. The CompMan services are not _Paused_
4. The _Conditional Compile Argument_ `CompMan = 1`
5. At least one of the open Workbooks must have referenced the CompMan Addin which results in an opened Addin.

### Pausing, continuing the CompMan Add-in
Pausing and continuing the Addin is possible when the Addin or the development instance of it is open.

  
## Contribution
Contribution of any kind is welcome. It may be likely that one is looking for a Raw/Clone-VB-Project service, described above but yet not implemented. The _Development-Instance_ Workbook is available as public Github repo from where it may be forked, installed and used.


[1]:https://gitcdn.link/repo/warbe-maker/VBA-Components-Management-Services/master/CompManDev.xlsb
[2]:https://GitHub.com/warbe-maker/VBA-Components-Management-Services