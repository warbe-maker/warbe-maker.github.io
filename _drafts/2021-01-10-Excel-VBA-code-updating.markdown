---
layout: post
title: "Code updates with VBA"
date:   2020-09-30 09:11:20 +0200
categories: vba excel component management
---


How to update the code in a VB project is often asked. The good message is: It is possible! The bad is: It is by far not as straight forward as some think or believe. The truth is, it is pretty tricky and requires different approaches depending on the kind of component/module.

## The hurdles
1. First of all, there is no safe and stable way for a Workbook to modify it's own code other than delegating this job/service is to another Workbook. And even the other Workbook has to de-activate the serviced Workbook before any code modification
2. A component cannot be simply replaced by removing it and (re-)importing an _Export File_ because any removal takes place when the service had finished. However, renaming and removing it does the trick. Once the to be renewed/updated component is renamed it has been put out of the way.

## The implementation
It almost had become a life's work because I've started the implementation over and over because it never turned out to be really bullet proof. Many suggestion I've found did not fullfil the promise. The below implementation is an extraction of my _Excel VBA Component Management_ and that's why it comes with several components.

## Installation

## Usage