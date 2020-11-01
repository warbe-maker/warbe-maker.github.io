---
layout: post
title: Implementing a Data structure by means of a Collection
subtitle: A UDT by alternative means
date: 2020-10-19 16:00 +0200
categories: vba common
toc: In this post
---

# Why avoiding a Class Module

There a cases where user defined type cannot be used, e.g. when the structure should be stored in a Collection but a class module should be avoided, e.g. to keep the number of to-be-installed components for a common component as few as possible.

# Specification of an example
## Objects
|  Object (Name)  |Properties/Items|
|-----------------|----------------|
|Employee (Mply)  | Name (Nm)      |
|                 | Age (Ag)       |
|                 | Salary (Slry)  |
|Employees (Mplys)| Employee       |

## Methods
Employee.Add
Employee.Delete
Employee.Update
Employee

# Implementation by means of a Collection
A Collection is an object which can store various types of data and any kind of object. I.e. a Collection can store Collections.

## Creating an Employee record
```vbs
Private Function Mply( _
          ByName ky As String, _
          ByName nm As String, _
          ByName ag As Long, _
          ByName slry As Currency) As Collection
    Dim cll As New Collection
    cll.Add ky
    cll.Add nm
    cll.Add at
    cll.Add slry
    Set Mply = cll
End Function
```

## Collecting Employee records
 
```vbs
Private Sub MplyAdd( _
            ByRef mplys As Collection, _
            ByName nm As String, _
            ByName ag As Long, _
            ByName slry As Currency
    If mplys Is Nothing The Set mplys = New Collection
    mplys.Add Mply( _
              ky:=CStr(mplys.Count + 1), _
              nm:=nm, _
              ag:=ag, _
              slry:=slry
End Sub
```

For the index of the elements in the Collection:
```vbs
Enum Employee Index
    iName
    iAge
End Enum
Private cllEmployees As Collection
```
For the elements:
```vbs
Property Get Name(Optional ByName entry As Collection) As String
    Name = entry(iName)
End Property
Property Get Age(Optional ByName entry As Collection) As String
    Name = entry(iAge)
End Property


Property Let Name(Optional ByName entry As Collection, s As String)
    entry.Add s, iName
End Property
Property Let Age(Optional ByName entry As Collection, ByVal l As Long)
    entry.Add l, iAge
End Property

```
To create an entry:
```vbs
Private Function EmpEntry(ByVal name As String, ByVal age As Long) As Collection
    Dim call As New Collection
    Name(cll) = name
    Age(cll) = age
    Set EmpEntry = cll
End Sub

Private Sub EmpAdd(ByVal name As String, ByVal age As Long)
    If cllEmployees Is Nothing The Set cllEmployees = New Collection
    cllEmployees.Add EmpEntry(name, age)
End Sub

Private Sub Employees
    EmpAdd "Smith", 45
    EmpAdd "Miller", 35
End Sub