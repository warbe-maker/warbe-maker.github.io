---
layout: post
title: Common Components Management
date:   2020-09-30 09:11:20 +0200
categories: vba excel management
---
### Hosted _Common Components_
When something looks like a common reusable procedure I stow it in a module which I call a _Common Component_ and this module is _hosted_. in a dedicated _Common Components  Workbook_.
Example: The module _mErrHndlr_ is _hosted_ in the _Common Component Workbook_ ErrHndlr.xlsm. For development and test the Workbook/Project has usually some other modules but only the _mErrHndlr_ is the one _hosted_.

### Used _Common Components_
When a Workbook which uses a _Common Component_ is opened and the component is outdated it is automatically updated. The source for the update/synchronization is the exported .bas, .frm, or .cls file which resides in the _Common Components  Workbook's_ dedicated folder.

The instance which is called by a VB Project which uses a _Common Component_ is a _Common Component Management Workbook_ which runs as _AddIn_. This ensures a stable update process because the AddIn us called by the "user".

 