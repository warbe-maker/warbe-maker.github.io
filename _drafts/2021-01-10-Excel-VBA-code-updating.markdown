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

## The _UpdateOutdatedClones_ service
A component which is developed, maintained and tested in another VB project can be called the _Raw Component_ component. The copy of this component used by another VB project can be called a _Clone Component_.
The service checks for any clone of which the raw has changed and replaces it.

## The _ExportChangedComponents_ service
The service is evoked with the Workbook_Before_Save event, compares the code of all components in a VB project with its last _Export File_, and exports them when different. VB projects which host _Raw Components_ as well as VB projects using them use the service. Raws hosting VB projects register these components as available for other VB projects using them. 

## Installation

## Usage