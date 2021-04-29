---
layout: post
title: Programmatically updating or synchronizing VBA code of Excel VB-Project Components
date:   2021-03-22
last_revised: 2021-04-29
categories: vba excel code component management
---
<!--more-->

## Introduction
This post focuses on
 - programmatically updating the code of individual _VB-Project-Components_
 - programmatically synchronizing  _VB-Projects_

The services are implemented as a dedicated Workbook and available either when the Workbook is open or when the Workbook is setup as _Addin-Workbook_. The services cater professional and semi-professional VB-Project development.

## Basic considerations
- A VB-Component developed, maintained and tested in one Workbook and used in many others is regarded a _Common-Component_ which should be updated when changed preferably automated
- A productive _VB-Project_ could be modified with a minimum downtime when a copy is modified and finally synchronized
- There is no safe and stable way to programmatically modify the code of a _VB-Project_  other than delegating this service to another dedicated _VB-Project_.
- A component cannot be simply removed and replaced by importing an _Export-File_ because the removal of a _VB-Component_ is postponed by the system until the running process has ended. However, renaming and removing does the trick because the rename puts the component out of the way for the import
- Service may be available either by means of an open Workbook or via an _Addin-Workbook_, in any case performed via `Application.Run`

## Synchronization specific considerations
- _Document-Modules_ (Workbook and Woeksheet) are updated by transferring the code from an _Export-File_ line by line
- The _Workbook Document-Module_ needs to be distinguished from any _Worksheet Document-Module_ in order to apply specific sheet synchronizations
- A _Worksheet Document-Module_ has a _Name_ and a _CodeName_. When both are renamed/changed the sheet in the source Workbook no longer relates to the corresponding sheet in the target Workbook and thus is regarded a new Worksheet. An assertion that never both are names are changed is explicitly requested to assure disambiguation
- _Worksheets_ may have new or outdated _Shapes_ and _Shapes-Properties_ which should be synchronized.
- _Worksheets_ may come with _Range-Names_ and design changes such like new/removed columns/rows which can only be synchronized when indicated through a _Synchronization-Manifest_.

## Disambiguation

| Term             | Meaning                  |
|------------------|------------------------- |
|_Component_       | Generic _VB&#8209;Project_ term for a _VB-Project-Component_ which may be a _Class Module_ , a  _Data Module_, a _Standard Module_, or a _UserForm_  |
|_Common Component_ | A _VB-Component_ which is hosted in one and commonly used by two or more _VB-Projects_ |
|_Clone&#8209;Component_ <br> | The copy of a _Raw&#8209;Component_ in a _VP&#8209;Project_ using it. _Clone-Components_ are kept up-to-date by an _UpdateRawClones_ service.<br>The term _clone_ is borrowed from GitHub but has a slightly different meaning because the clone is usually not maintained but the _raw_. |
|_Procedure_     | Any - Public or Private _Property_, _Sub_, or _Function_ of a _Component_. See also _Service_.
|_Raw&#8209;Component_ | The instance of a _Common Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw&#8209;Host_ Workbook. The term _raw_ is borrowed from GitHub and indicates the original version of something |
|_Raw&#8209;Host_      | The Workbook/_VB-Project_ which hosts the _Raw-Component_ |
|_Service_       | Generic term for any _Public Property_, _Public Sub_, or _Public Function_ of a _Component_ |
|_VB&#8209;Clone&#8209;Project_ | A _VP-Project_ which is a copy (i.e regarding the VB-Project code a clone) of a corresponding  _VB&#8209;Raw&#8209;Project_. The code of the clone project is kept up-to-date by means of a code synchronization service. |
|_VB-Project_     | In the present case this term is used synonymous with Workbook |
|_Source&#8209;Workbook-/_<br>_Source-VB-Project_   | The temporary copy of productive Workbook which becomes by then the _Target-Workbook/Project for the synchronization.|
| _Workbook-&#8209;Folder_ | A folder dedicated to a Workbook/VB-Project with all its Export-Files and other project specific means. Such a folder is the equivalent of a Github-Repo-Clone (provided Github is used for the project's versioning which is recommendable.|

## Services
### _ExportChangedComponents_
Compares the code of any component in a _VB-Project_ with its last _Export File_ and re-exports it when different. The service is essential for _VB-Projects_ with _Hosted-Components_ used in other Workbooks. The service registers the hosted component as _Raw-Component_ which enables the _UpdateRawClones_ service to recognize and update them as _Clone-Component_.
- The service checks if a _Clone-Component_ has been modified within the _VB-Project_ using it. This may happen when the change of a _Common-Component_ appears appropriate directly in the _VB-Project_ which triggered the change. In this case the service offers a choice for updating the _Raw-Component_ in order to make the modification permanent for all other _VB-Projects_ using the component (testing however will remain a task for the hosting _Workbook/VB-Project_.<br>
It also should be noted that changes not made 'public' will be reverted by the _UpdateRawClones_ service the next time the Workbook is opened.

### _UpdateRawClones_
The service checks each component in a _VB-Project_ whether it is a known/registered _Raw-Component_. If yes, the component is regarded a _Clone-Component_ and updated if the raw had changed. See confirmation dialog below.

![](../Assets/UpdateRawCloneConfirmationDialog.png)
![](/Assets/UpdateRawCloneConfirmationDialog.png)
<br>

Note 1: The service is meant (and tested) only for _Standard-Modules_, _Class-Modules_, and _UserForms_. It may - in theory - be used for Worksheets as well. However, this would mean that one Workbook hosts a Worksheet which is regarded a _Common-Component_ used in other Workbooks. Something pretty unlikely I think. Whenever this case may become true extra testing and possibly some modification with the service may become required.
Note 2: This service must not be confused with a synchronization service. A synchronization service uses one Workbook as the source and synchronizes the corresponding target Workbook. This update service will have different Workbook's Export-Files as the source, depending on where the _Raw-Component_ is hosted.

### _SyncVBProjects_
The service synchronizes a _target-Workbook_ with a _source-Workbook_ whereby the _source-Workbook_ is a temporary copy of the **productive** _target-Workbook_. While the **productive** Workbook remains in use the VB-Project of the _source-Workbook_ can be made without time restraint. When the modification, maintenance, bug-fixing, etc. is finished all changes can be synchronized by a minimized downtime for the **productive** workbook.

#### Synchronization coverage

| Item              | Extent of synchronization |
| ----------------- | ------------------------- |
|_References_       | New, obsolete             |
|_Standard&#8209;Modules_<br>_Class&#8209;Modules_<br>_UserForms_| New, obsolete, code change |
|_Data&#8209;Modules_| _Workbook-Module_: Code change<br>_Worksheet-Module_: New, obsolete, code change, shapes|
|_Shapes_           | New, obsolete, properties (may still be incomplete)            |
|_ActiveX-Controls_ | None. May be added in the future                               |
|_Names_            | New and obsolete will be recognized but (yet) not synchronized.|

The service is (usually) called without arguments and thus displays a dialog for the selection of the source and the target Workbook (which may be already open). The service displays all synchronization issues for confirmation (see example below). In case new or obligatory Worksheets had been detected an explicit assertion is required that never both, the **Name** and the **CodeName** of a sheet is changed. Confirmation dialog example:

![](../Assets/SyncIssuesConfirmation.png)
![](/Assets/SyncIssuesConfirmation.png)

When asserted and confirmed all synchronizations are logged in a file _CompMan.Services.log_ in the target Workbook folder.<br>Example of the synchronization log:
<small>
```
21-03-20 18:14:02 SynchVBProjects by CompMan.xlsb for 'Test_Sync_Target.xlsb': 
21-03-20 18:14:02 -------------------------------------------------------------------
21-03-20 18:14:13 Worksheet       Test_B1(wsSyncTest_B) ............: Name changed to 'Test_B1'.
21-03-20 18:14:14 Worksheet       Test_C_New(wsSyncTest_C_new) .....: Copied from source Workbook.
21-03-20 18:14:25 Name            celUsedInTest_C_New(=Test_A!$C$5) : Link to source sheet removed
21-03-20 18:14:25 Worksheet       Test_Obsolete(wbSyncTest_Obsolete): Obsolete (deleted)
21-03-20 18:17:24 Worksheet       wsSyncTest_C_new .................: Code updated with code from Export-File '.....'
21-03-20 18:17:24 Worksheet       Test_A(wsSyncTest_A1) ............: Order synched!
21-03-20 18:17:24 Worksheet       Test_C_New(wsSyncTest_C_new) .....: Order synched!
21-03-20 18:17:24 Worksheet       Test_B1(wsSyncTest_B) ............: Order synched!
21-03-20 18:17:24 Shape           Button 4 .........................: Copied from source sheet
21-03-20 18:17:24 Shape           Check Box 6 ......................: Copied from source sheet
21-03-20 18:17:24 Shape           Drop Down 5 ......................: Copied from source sheet
21-03-20 18:17:24 Shape           Group Box 3 ......................: Copied from source sheet
21-03-20 18:17:24 Shape           Label 10 .........................: Copied from source sheet
21-03-20 18:17:24 Shape           List Box 8 .......................: Copied from source sheet
21-03-20 18:17:24 Shape           Option Button 9 ..................: Copied from source sheet
21-03-20 18:17:24 Shape           Scroll Bar 11 ....................: Copied from source sheet
21-03-20 18:17:24 Shape           Spinner 7 ........................: Copied from source sheet
21-03-20 18:17:24 Shape           CommandButton1 ...................: Property 'Left' synched
21-03-20 18:17:24                                                     Property 'Top' synched
21-03-20 18:17:24 Shape           List Box 8 .......................: Property 'Height' synched
21-03-20 18:17:24 Shape           CommandButton1 ...................: Property 'Height' synched
21-03-20 18:17:24                                                     Property 'Left' synched
21-03-20 18:17:24                                                     Property 'Top' synched
21-03-20 18:17:24                                                     Property 'Width' synched
21-03-20 18:17:24 Shape           CommandButton2 ...................: Property 'Left' synched
21-03-20 18:17:24                                                     Property 'Top' synched
21-03-20 18:17:25 Standard-Module mNewModule .......................: Component imported from Export-File '.......'
21-03-20 18:17:25 Standard-Module mObsoleteModule ..................: Removed!
21-03-20 18:17:25 UserForm        fObsoleteUserForm ................: Removed!
```
</small>

The service has the following syntax:<br>
`mService.SyncVBProjects target-workbook, source-workbook, backup-folder`<br>
backup-folder is an argument returned by the function which ends with TRUE when the synchronization had been performed (it may have been terminated with the confirmation dialog).

#### Synchronization safety
Each synchronization creates a backup of the _Target-Workbook_ by creating a copy with a .backup extension. In case of a problem this copy just needs to be renamed (better ideas welcome).

## Installation
1. Download and open [CompMan.xlsb][1]
2. Perform _Setup/Renew_ even if you not intend to use CompMan as Addin. Setting up the basic configuration is obligatory even when the [CompMan.xlsb][1] is directly used
3. Follow the instructions to identify a _Serviced-RootFolder_ and a dedicated _CompMan-Addin-Folder_

Once the Add-in is established it will automatically be opened when Excel starts. However, the services may still be paused to avoid execution in case of problems (see Usage below for further required preconditions).

## Usage
### Common preconditions
The update and the export service will be denied unless the following preconditions are met:
1. The basic configuration - confirmed with each Setup/Renew is complete and valid
2. The serviced Workbook resides in a sub-folder of the configured _ServicedRootFolder_
3. The serviced Workbook is the only Workbook in its parent folder
4. The CompMan services are not _Paused_
5. WinMerge is installed

### Common usage requirement
In any Workbook either using the _ExportChangedComponents_ and/or the _UpdateChangedRawClones_ service copy the following in a Standard-Module called _mCompManClient_:
```vb
Option Explicit
' ----------------------------------------------------------------------
' Standard Module mCompManClient, optionally used by any Workbook to:
' - update used 'Common-Components' (hosted, developed, tested,
'   and provided, by another Workbook) with the Workbook_open event
' - export any changed VBComponent with the Workbook_Before_Save event.
'
' W. Rauschenberger, Berlin March 2021
'
' See also Github repo:
' https://github.com/warbe-maker/Excel-VB-Components-Management-Services
' ----------------------------------------------------------------------

Public Sub CompManService(ByVal cm_service As String, _
                          ByVal hosted As String)
' -----------------------------------------------------
' Execution of the CompMan service (cm_service) pre-
' ferably via the CompMan-Addin or when not available
' alternatively via the CompMan.xlsb Workbook.
' -----------------------------------------------------
    Const COMPMAN_BY_ADDIN = "CompMan.xlam!mCompMan."
    Const COMPMAN_BY_DEVLP = "CompMan.xlsb!mCompMan."
    
    On Error Resume Next
    Application.Run COMPMAN_BY_ADDIN & cm_service, ThisWorkbook, hosted
    If Err.Number = 1004 Then
        On Error Resume Next
        Application.Run COMPMAN_BY_DEVLP & cm_service, ThisWorkbook, hosted
        If Err.Number = 1004 Then
            Application.StatusBar = "'" & cm_service & "' neither available by '" & COMPMAN_BY_ADDIN & "' nor by '" & COMPMAN_BY_DEVLP & "'!"
        End If
    End If
End Sub
```

### Using the _ExportChangedComponents_ service
This service is crucial for all Workbooks which either host a commonly used component or which may become the source for a synchronization because both rely on up-to-date Export-Files.

In the concerned Workbook's Workbook-Component copy:
```vb
                                    ' -------------------------------------------------------------
Private Const HOSTED_RAWS = ""      ' Comma delimited names of Common Components hosted, developed,
                                    ' tested, and provided by this Workbook - if any
                                    ' -------------------------------------------------------------
```

and in the concerned Workbook's Workbook_BerforeSave event procedure copy:
```vb
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
End Sub
```

### Using the _UpdateRawClones_ service
In the concerned Workbook's Workbook_Open event procedure copy:
```vb
Private Sub Workbook_Open()
    mCompManClient.CompManService "UpdateRawClones", HOSTED_RAWS
End Sub

```

### Using the synchronization service ( _SyncVBProjects_)
In the _Immediate Window_ enter mService.SyncTargetWithSource. A dialog will open for the selection of the source and the target Workbook. The are selected by their files even when already open. Opening them beforehand may be appropriate in case there are some used _Common-Components_ yet not up-to-date. A VB-Project synchronization will follow the steps:
1. Prepare the **productive** Workbook/VB-Project for using the _ExportChangedComponents_ service
2. Prepare the **productive** Workbbok/VB-Project for using the _UpdateRawClones_ service in case it uses _Common-Components_ hosted elsewhere
3. Copy the Workbook under a different name into a dedicated sub-folder of the configured _Serviced-Root-Folder_.
4. Perform all required changes while the **productive** Workbook remains in use
5. When the required modifications had been finished and successfully tested
6. Move the **productive

### Setup/Renew _CompMan-Addin_
When the [CompMan.xlsb][1] Workbook is opened the services are all available. _Setup/Renew_ offers the option to establish the services as an Addin-Workbook. The steps are logged as follows

Initial setup (in this case Addin-Workbook existed already but was not open):
![](../Assets/CompManSetupResult_not_open.png)
![](/Assets/CompManSetupResult_not_open.png)

Renew (in this case the Addin was already open and thus some more steps were required):
![](../Assets/CompManAddinRenewResult_addin_open.png)
![](/Assets/CompManAddinRenewResult_addin_open.png)

Each Setup/Renew request the confirmation or specification of a _Basic-CompMan-Configuration_ which is a _Service-Root-Folder_ (only Workbooks residing therein are serviced - productive Workbooks are not touched), and a dedicated folder for the Addin-Workbook (additional system files are stored therein as well). The Addin-Workbook folder may be available for development purpose only and hidden from any productively used Workbook. 

![](../Assets/CompManBasicConfigurationDialog.png)
![](/Assets/CompManBasicConfigurationDialog.png)

### Pause/Continue the CompMan-Addin
Use the corresponding command buttons when the [CompMan.xlsb][1] Workbook is open. While paused services will be denied.

  
## Contribution
Contribution of any kind is welcome commenting this post or raising issues with the [GitHub repo][2].


[1]:https://gitcdn.link/repo/warbe-maker/VBA-Components-Management-Services/master/CompMan.xlsb
[2]:https://GitHub.com/warbe-maker/VBA-Components-Management-Services