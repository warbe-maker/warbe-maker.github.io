---
layout: post
title: "Code updates with VBA"
date:   2020-09-30 09:11:20 +0200
categories: vba excel component management
---


How to update the code in a VB project is often asked. It is possible but not as straight forward as one may think or believe.

## Approaches and their applicability
First of all, there is no safe and stable way for a Workbook to modify it's own code. The minimum requirement is an Workbook to which the service is delegated. For a best possible stability this Workbook's service will de-activate the serviced Workbook before any code modification.<br>
Furthermore a component cannot be simply replaced by removing it and (re-)importing an _Export File_ because any removal will take place when the service had finished. However, renaming and removing it does the trick. Once the to be renewed/updated component is renamed it has been put out of the way. Unfortunately this approach is not applicable for a UserForm Module, neither is it applicable for a Data Module.

## The implementation
It almost had become a life's work because I've started the implementation over and over because it never turned out to be really bullet proof. Many suggestion I've found did not fullfil the promise. The below implementation is an extraction of my _Excel VBA Component Management_ and that's why it comes with several components.

## Installation

## Usage