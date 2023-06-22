---
layout:        post
title:         Managing UserForm Instances
subtitle:      An easy way to have any number of UserForm instances managed with a minimum effort
date:          2021-11-09
modified_date: 2023-06-19
categories:    vba common
---
An easy way to have any number of UserForm instances managed with a minimum effort.
<!--more-->

## The _FormInstance_ function
> - **fProcTest** will have to be replaced by the desired UserForm
> - The function requires a Reference to _Microsoft Scripting RunTime_

```vb
Private Function FormInstance(ByVal fi_key As String, _
                     Optional ByVal fi_unload As Boolean = False) As fProcTest
' -------------------------------------------------------------------------
' Returns an instance of the UserForm fProcTest which is definitely
' identified by anything uniqe for the instance (fi_key). This may be what
' becomes the title (property Caption) or even an object such like a
' Worksheet (if the instance is Worksheet specific). An already existing or
' new created instance is maintained in a static Dictionary with fi_key as
' the key and returned to the caller. When fi_unload is true only a possibly
' already existing Userform identified by fi_key is unloaded.
'
' Requires: Reference to the "Microsoft Scripting Runtime".
' Usage   : The fProcTest has to be replaced by the name of the desired
'           UserForm
' -------------------------------------------------------------------------
    Static Instances As Dictionary    ' Collection of (possibly still active) instances
    
    If Instances Is Nothing Then Set Instances = New Dictionary
    
    If fi_unload Then
        If Instances.Exists(fi_key) Then
            On Error Resume Next
            Unload Instances(fi_key) ' The instance may be already unloaded
            Instances.Remove fi_key
        End If
        Exit Function
    End If
    
    If Not Instances.Exists(fi_key) Then
        '~~ There is no evidence of an already existing instance
        Set FormInstance = New fProcTest
        Instances.Add fi_key, FormInstance
    Else
        '~~ An instance identified by fi_key exists in the Dictionary.
        '~~ It may however have already been unloaded.
        On Error Resume Next
        Set FormInstance = Instances(fi_key)
        Select Case Err.Number
            Case 0
            Case 13
                If Instances.Exists(fi_key) Then
                    '~~ The apparently no longer existing instance is removed from the Dictionarys
                    Instances.Remove fi_key
                End If
                Set FormInstance = New fProcTest
                Instances.Add fi_key, FormInstance
            Case Else
                '~~ Unknown error!
                Err.Raise 1 + vbObjectError, ErrSrc(PROC), "Unknown/unrecognized error!"
        End Select
        On Error GoTo -1
    End If

End Function
```

## Test of the _FormInstance_ function

```vb
Public Sub Test_FormInstance()
' ------------------------------------------------------------------------------
' Creates a number of instances of the UserForm named fProcTest and unloads them
' in the revers order. Application.Wait is used to allow the observation of the
' process.
' Note: The test shows that is not required to have a variable for the instance
'       object. It may however make sense in practise.
' ------------------------------------------------------------------------------
        
    Dim i   As Long
    Dim key As String
    Dim obj As Object ' not required for the function but only to get the UserForm's name
    
    For i = 1 To 5
        key = "Instance-" & i
        '~~ Set obj ... will create the instance. However, this is not not required.
        '~~ It is just used to obtain the UserForms name
        Set obj = FormInstance(fi_key:=key)
        With FormInstance(fi_key:=key)
            .Height = 80
            .Width = 200
            .Caption = key & " of UserForm '" & obj.Name & "'"
            .Show Modeless
            .Top = 30 * i
            .Left = 30 * i
        End With
        Application.Wait Now() + 0.000006
    Next i
    
    For i = 5 To 1 Step -1
        key = "Instance-" & i
        '~~ Unloading the instance this way has two advantages:
        '~~ 1. The instance is removed from the Dictionary
        '~~ 2. No error in case the instance no longer exists
        FormInstance fi_key:=key, fi_unload:=True
        Application.Wait Now() + 0.000006
    Next i
    
End Sub
```
Displays:

![](../Assets/UserForm-Instances.gif)
![](/Assets/UserForm-Instances.gif)



