---
layout: post
title: Common Components Management
date:   2020-09-30 09:11:20 +0200
categories: vba excel management
---
<small>The matter of this post is available on [GitHub.com](<https://github.com/warbe-maker/Common-Component-Management>)</small>

In this post
[Background](#background)<br>
[Administration of _Common Components_](#administration-of-common-components)<br>
[VB-Projects and GitHub](#vb-projects-and-github)


### Background
_Component_ is the generic _VB-Project_ term for UserForm, Standard Module, Class Module and Data Module. I call a  _Component_ which potentially may be used in (m)any VB-Projects a _Common Component_. The consequence of this _Common_ property is a demand on solid implementation and careful testing. This is facilitated by a dedicated _Common Component Workbook_ which resides in a dedicated _Common Component Workbook Folder_. The later is the 

Example: The module _mErrHndlr_ is _hosted_ in the _Common Component Workbook_ ErrHndlr.xlsm which contains all means required/usefull for development and test.

### Administration of _Common Components_
When something looks like a common reusable procedure it's appropriate to develop, test, and maintain it in a dedicated _Common Components  Workbook_, i.e. a Workbook _hosting_ it. The dedictated folder the Workbook resides in is a perfect clone of a _GitHub repository_.
Each time the Workbook is saved or closed all therein  _hosted Common Components_ are registered as such and exported to the dedicated _Common Component Workbook Folder_ - provided the code has changed. The exported _Common Component_ is the source for the update in an VB-Project using it.
 
#### VB-Projects and GitHub
I've ignored GitHub for very long assuming that it is for true professionals and not the appropriate means for the administration of VB-Projects. Looking for the perfect place to share _Common Components_ with the world - and ideally find/invite contributors keeping them vital I found GitHub as the perfect place.

### Used _Common Components_
When a Workbook is opened all components are checked for being a used _Common Component_ and automatically updated when their code is outdated when compared with the exported .bas, .frm, or .cls file.

### Export hosted and update used _Common Components_
The required methods are provided by a _Common Components Management_ Workbook opened as Addin for VB-Projects which require it - indicated by a Reference to the Addin. This "outsourced" update service surpasses problems which occur when a VB-Project tries to update its own components.

 