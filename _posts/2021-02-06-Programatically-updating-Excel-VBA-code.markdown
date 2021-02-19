---
layout: post
title: "Programatically updating Excel VBA code"
date:   2021-02-05
categories: vba excel code component management
---

## Introduction
Programmatically updating the code of a _VB-Project_ is not straight forward like removing and re-importing a component. Synchronizing all the code between two _VB-Projects_ is an even more ambitious service. Re-started several times I've finally ended up with a set of satisfyingly stable services provided via an Addin-Workbook.


## Challenges
1. There is no safe and stable way for a _VB-Project_ to uodate it's own code other than delegating this service to another _VB-Project_.
2. A component cannot be simply removed and replaced by importing an _Export File_ because the removal of a component is postponed by the system until the running process has ended. However, renaming and removing does the trick because the rename puts the component out of the way for the import.
3. An update service which can be called by any _VB-Project_ (via Application.Run) must be available as an opened Workbook. A Workbook automatically opened for a another one is only possible via a referenced! **Add-in**. The birth of a _Component-Management_ Addin-Workbook which turned out to be much more complex than expected in the first place.
4. Updating individual components developed, maintained and (hopefully appropriately) tested in one _VB-Project_ and used by others I've successfully implemented. Synchronizing all code in a kind of Raw-Clone-Project approach still looks like opening a can of worms and will suffer from some limitations too complicated to be eliminated.

## Disambiguation
The terms below are not only those used in this post but also used with the implementation of the _Component Management_.

| Term             | Meaning                  |
|------------------|------------------------- |
|_Component_       | Generic _VB-Project_ term for a _Class Module_, a  _Data Module_, a _Standard Module_, or a _UserForm_  |
|_Common Component_| A _Component_ which is used by two or more _VB-Projects_ |
| _Raw-Component_ | The instance of a _Common Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw-Host_ Workbook. |
|_Clone-Component_,<br>_Raw-Clone_ | The copy of a _Raw-Component_ in a _VB-Project_ using it |
|_VB-Clone-Project_ | A _VB-Project_ derived from a _VB-Raw-Project_ and productively used (the code is maintained in the _VB-Raw-Project_ |
|_Procedure_     | Any - Public or Private _Property_, _Sub_, or _Funtion_ of a _Component_. See also _Service_.
|_Raw-Host_.     | The Workbook/_VP-Project_ which hosts the _Raw-Component_ |
|_Raw-Project_   | A code-only _VP-Project_ of which all components are regarded _Raw-Components_. A _Raw-Project_ is kind of a template for the productive version of it. In contrast to a classic template it is the life-time raw code base for the productive _Clone-Project_.  The service and process of 'synchronizing' the productive (clone) code with the raw is part of the _Component Management_.|
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
2. For a Workbook which hosts _Raw_Components_ specify them in the HOSTED_RAWS constant. If its more then one, have the component's names delimited with commas.

> ++**Be aware:**++ The Workbook component will be one of which the code cannot be updated by any means because it contains the code executed to perform the update. Thought this will only be relevant for Raw/Clone-VB-Projects which are yet not supported. However, as a consequence only calls to procedures provided with all arguments will remain in the Workbook component code and all the rest will be in a dedicated mWorkbook component.

## Usage
### The preconditions
The export and the update service have the following preconditions:
1. The respective below code snippet is copied to the concerned Workbbok
3. The concerned Workbook is located in any folder within the configured _ServicedRootFolder_
4. The _Conditional Compile Argument_ `CompMan = 1`
5. The Workbook has the Add-in (CompMan) referenced

From the above it follows: When the Workbook is moved/copied to a location outside the _ServicedRootFolder_ which is supposed to be the 'productive' location of the Workbook:
1. The _Conditional Compile Argument_ `CompMan = 0`
1. The Workbook has the Add-in (CompMan) un-referenced

### Pausing, continuing the CompMan Add-in
Pddinusing and again continuing the Add-in is possible in the opened development instance.  When the Add-in is 'paused' the export and the update service will not be executed. Pausing is thus a kind of emergency stop in case the CompMan Add-in seriously fails servicing properly.

## A professional code update process
Proposition: A productive Workbook's code must not be updated but a copy of it. When the update of the copy was successful the productive version will finally be moved to a release stage, updated, finally tested and moved back. Whichever process is chosen should consider: 
- How critical is the Workbook for the business process
- What is the acceptable downtime of the Workbook
- How can the downtime be kept to a minimum
- Planning the code update process like a software release

  
## Contribution
Contribution of any kind is welcome. It may be likely that one is looking for a Raw/Clone-VB-Project service, described above but yet not implemented. The _Development-Instance_ Workbook is available as public Github repo from where it may be forked, installed and used.


[1]:https://gitcdn.link/repo/warbe-maker/VBA-Components-Management-Services/master/CompManDev.xlsb
[2]:https://github.com/