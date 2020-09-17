---
layout: post
title: Min and Max
subtitle: Minimum and Maximum out of any number of values
---
I've looked for these "bread and butter"  functions and did not find something really convincing. The below are in my _mBasics_ module in any VBA PROJECT almost by default. They do the job for any number of values:
```vbscript
Private Function Max(ParamArray va() As Variant) As Variant
' ---------------------------------------------------------
' Returns the maximum of all values provided (va).
' ---------------------------------------------------------
   Dim v As Variant
   For Each v In va
      If v > Max Then Max = v
   Next v
End Function

Private Function Min(ParamArray va() As Variant) As Variant
' ---------------------------------------------------------
' Returns the minimum of all values provided (va). The
' returned type depends on the provided values. If all are
' of type Integer the returned type is Integer else it is
' Double.
' ---------------------------------------------------------
   Dim v As Variant
   Min = va(LBound(va))
   For Each v In va
      If v < Min Then Min = v
   Next v
End Function
```
Test:
```vbscript
Public Sub Test_Max()
    '~~ Test includes errorneous empty elements
    Debug.Assert mMsg.Max(9, , 100, 355, 2, 9, 50) = 355
    Debug.Assert mMsg.Min(9, 100, 355, 2, 9, , 50) = 2
End Sub
```
