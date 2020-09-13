---
layout: post
title: Min and Max
---
I've looked for these functions and did not find something really convincing. The below do the job for any number of values provided as a ParmArray:
```vbscript
Private Function Max(ParamArray va() As Variant) As Variant' ------------------------------------------------------' Returns the maximum value of all values provided (va).' ------------------------------------------------------ Dim v As Variant Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function For Each v In va If v > Max Then Max = v Next v End Function Private Function Min(ParamArray va() As Variant) As Variant' ------------------------------------------------------' Returns the minimum (smallest) of all provided values.' ------------------------------------------------------ Dim v As Variant Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function For Each v In va If v < Min Then Min = v Next v End Function
```