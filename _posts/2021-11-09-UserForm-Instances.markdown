---
layout: post
title: Managing UserForm Instances
subtitle: An easy way to have any number of UserForm instances managed with a minimum effort
date:          2021-11-09
modified_date: 2021-11-09
categories:    vba common
---
An easy way to have any number of UserForm instances managed with a minimum effort.
<!--more-->

## The _FormInstance_ function
> **Const PROC** and the error handling are my personal standard which of course may or may not be used. Just in case of interest and to complete the matter all the used code is added below as well.

> **fProcTest** will have to be replaced by the desired UserForm.

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
    Const PROC = "FormInstance"
    
    On Error GoTo eh
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

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function
```

## Test of the _FormInstance_ function
The test procedure by-the-way proves that it is not required to have a variable for the instance. I.e. any instance may be addressed "directly". 
```vb
Public Sub Test_FormInstance()
' ------------------------------------------------------------------------------
' Creates a number of instance of the UserForm named fProcTest and unloads them
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

## Used error handling procedures
```vb
Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' This is a kind of universal error message which includes a debugging option.
' It may be copied into any module as a Private Function. When the/my Common
' VBA Error Handling Component (ErH) is installed and the Conditional Compile
' Argument 'CommErHComp = 1' indicates this the error message will be displayed
' by means of the Common VBA Message Component (fMsg, mMsg) which is part of it.
'
' Usage: Example of using this function in any procedure. With the Conditional
'        Compile Argument 'Debugging = 1' the debugging option will be available
'        as follows - want do anything if not.
'
'            Const PROC = "procedure-name"
'            On Error Goto eh
'        '   ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC)
'               Case vbYes: Stop: Resume
'               Case vbNo:  Resume Next
'               Case Else:  Goto xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but for sure will be recognised as 
'        a godsend in case of an error!
'
' Uses:  - AppErr for programmed application errors (Err.Raise AppErr(n), ....).
'          AppErr turns a positive number into a negative one thereby avoiding any
'          conflict with VB Runtime Error. The error message in return will regard a
'          negative error number as an 'Application Error' and will use AppErr to
'          turn it back in its original positive number.
'        - ErrSrc is used to identify the source of an error.
' ------------------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              err_dscrptn & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    
#If Debugging Then
    ErrBttns = vbYesNoCancel
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume error line" & vbLf & _
              "No     = Resume Next (skip error line)" & vbLf & _
              "Cancel = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
#If CommErHComp Then
    '~~ When the Common VBA Error Handling Component (ErH) is installed/used by in the VB-Project
    ErrMsg = mErH.ErrMsg(err_source:=err_source, err_number:=err_no, err_dscrptn:=err_dscrptn, err_line:=err_line)
    '~~ Translate back the elaborated reply buttons mErrH.ErrMsg displays and returns to the simple yes/No/Cancel
    '~~ replies with the VBA MsgBox.
    Select Case ErrMsg
        Case mErH.DebugOptResumeErrorLine:  ErrMsg = vbYes
        Case mErH.DebugOptResumeNext:       ErrMsg = vbNo
        Case Else:                          ErrMsg = vbCancel
    End Select
#Else
    '~~ When the Common VBA Error Handling Component (ErH) is not used/installed there might still be the
    '~~ Common VBA Message Component (Msg) be installed/used
#If CommMsgComp Then
    ErrMsg = mMsg.ErrMsg(err_source:=err_source)
#Else
    '~~ None of the Common Components is installed/used
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
#End If
#End If
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "<module-name>." & sProc
End Function

```

