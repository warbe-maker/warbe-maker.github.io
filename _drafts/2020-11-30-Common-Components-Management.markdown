---
layout: post
title: Excel VB Component Management
subtitle: Keeping Common Components up-to-date in Vb-Projects using them
date:   2020-09-30
categories: vba excel common component management
---
## Abstract
Dealing with components of an Excel VBProject is both, delicate and tricky. Delicate, because it requires "Trust access to the VBA project object model" which is regarded a potential security flaw. It is tricky because it is not possible to remove and re-import a component.

## Services
The key service for others like code synchronization or code update is replacing a component via the import of an Export File - which usually fails because a component cannot be deleted within a code executed. It always is deleted as soon as the running process has ended. The general trick however are the following steps:
1. Rename the to be replaced component and delete it (as mentioned it will not be deleted before the running process has come to an end)
2. Re-import the component via an Export File


The code to do this may look as follows:



## Disambiguation of used terms
| Term | Meaning
|------|--------
| _Component_ | Generic _VB-Project_ term for a UserForm, Standard Module, Class Module or Data Module |
_Common Component_ | A _Component_ which is shared ong two or more VB-Projects |
| _Raw_ | The instance of a _Common Component_ which is regarded the original. In other words the component in a Workbook/VB-project which is dedicated to its development,  maintenance and test. I.e. the Workbook which has the means to ensure the desired quality of the services the component provides |
| _Clone_ | The copy of a _Raw_ component in any Workbook/VP-Project using it |
|_Host_ | The Workbook/VP-Project which hosts the _Raw_ component |
| _Template Project_ | A Workbook/VP-Project of which all components are regarded _Raw_ component. A _Template Project_ is a "code-only-project and does not have any data other than static data |
| _Clone Project_ | A Workbook/VP-Project derived from a _TemplateProject_ |
| _Workbook Folder_ | A folder dedicated to a Workbook/VB-Project with all its Export Files and other project specific means. Such a folder is the equivalent of a Git-Repo-Clone (provided Git is used for the project's versioning which is recommendable |


## Administration of _Common Components_
When something looks like a common reusable procedure it's appropriate to develop, test, and maintain it in a dedicated _Common Components  Workbook_, i.e. a Workbook _hosting_ it. The dedictated folder the Workbook resides in is a perfect clone of a _GitHub repository_.
Each time the Workbook is saved or closed all therein  _hosted Common Components_ are registered as such and exported to the dedicated _Common Component Workbook Folder_ - provided the code has changed. The exported _Common Component_ is the source for the update in an VB-Project using it.
 
#### VB-Projects and GitHub
I've ignored GitHub for very long assuming that it is for true professionals and not the appropriate means for the administration of VB-Projects. Looking for the perfect place to share _Common Components_ with the world - and ideally find/invite contributors keeping them vital I found GitHub as the perfect place.

### Used _Common Components_
When a Workbook is opened all components are checked for being a used _Common Component_ and automatically updated when their code is outdated when compared with the exported .bas, .frm, or .cls file.

### Export hosted and update used _Common Components_
The required methods are provided by a _Common Components Management_ Workbook opened as Addin for VB-Projects which require it - indicated by a Reference to the Addin. This "outsourced" update service surpasses problems which occur when a VB-Project tries to update its own components.

 