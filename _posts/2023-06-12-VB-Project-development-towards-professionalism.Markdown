---
layout: post
title: Excel VB-Project development towards professionalism
subtitle: 
date:   2023-06-12
categories: vba excel component management common components
---
## Excel VB-Project development towards professionalism
It's the heritage of may active professional IT life time and the result of my retirement life: Means substantially supporting - mainly Excel - VB-Project development. A bunch of ***Common Component*** increase efficiency of development and a ***Component-Management*** keep this ***Common Components*** up-to-date in all VB-Projects using them.

### Common Components
A peace of code with properties potentially useful for more than one VB-Project, triggers the birth of a Common Component. A dedicated development and test Workbook hosting it provides the means for:
- **Execution trace**: supports test documentation and helps identifying the source of performance issues
- **Elaborated error handling**: supports debugging and regression testing 
- **log services**: make it easy to write entries to log-files
- etc.

Each of my Common Components work completely autonomous. I.e they are 'single VB-Components not requiring other modules, the use of other Common Components remains optional.

### Components Management (Excel only)
- An **Update outdated Common Components** service keeps used **Common Components** up-to-date by updating any outdated when a Workbook is opened for development/maintenance (but not when opened productively!)
- An **Export of changed VB-Components** service supports versioning - in combination with versioning means like GitHub (what I prefer)
- A **Synchronization VB-Projects** service supports development and maintenance without a downtime for the productive Workbook


