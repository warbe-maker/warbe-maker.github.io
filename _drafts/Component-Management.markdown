---
layout: post
title: Common Components Management
date:   2020-09-30 09:11:20 +0200
categories: vba excel management
---
<small>The matter of this post is available on [GitHub.com](<https://github.com/warbe-maker/Common-Component-Management>)</small>
### Hosted _Common Components_
When something looks like a common reusable procedure I stow it in a module and call it a _Common Component_  developed, tested and maintained in a dedicated _Common Components  Workbook_.
Example: The module _mErrHndlr_ is _hosted_ in the _Common Component Workbook_ ErrHndlr.xlsm which also provides all means for testing.
When the Workbook is closed all hosted _Common Components are registered as such and exported into a dedicated Workbook folder - provided the code has changed.

### Used _Common Components_
When a Workbook is opened all components are checked for being a used _Common Component_ and automatically updated when their code is outdated when compared with the exported .bas, .frm, or .cls file.

### Export hosted and update used _Common Components_
The required methods are provided by a _Common Components Management_ Workbook opened as Addin for VB-Projects which require it - indicated by a Reference to the Addin. This "outsourced" update service surpasses problems which occur when a VB-Project tries to update its own components.

 