---
layout: post
title: Add items to a Dictionary "ordered"
date:   2020-09-25 16:00 +0200
categories: vba basic
---
Instead of sorting a dictionary when all items had been added, adding them directly in the desired order is an option. The below procedure adds items to a dictionary in either of the modes enumerated in ```enDctMode```.

The performance may suffer when several hundreds of items are added and for many of them the entry sequence is not the specified one. Testing available in the dedicated _Common Component Workbook_ Basic.xlsm has proven that the procedure works fine even when the key is an object, provided the object has a name property.

````vbscript
Public Enum enDctMode ' Dictionary add/insert modes
    dct_addafter
    dct_addbefore
    dct_ascendingcasesensitive
    dct_ascendingcaseingignored
    dct_descendingcasesensitive
    dct_descendingcaseignored
    dct_sequence
End Enum
```
```vbscript
Public Sub DctAdd(ByRef dct As Dictionary, _
                  ByVal dctkey As Variant, _
                  ByVal dctitem As Variant, _
         Optional ByVal dctmode As enDctMode = dct_sequence, _
         Optional ByVal dcttargetkey As Variant)
' ----------------------------------------------------------------------
' Adds the item (dctitem) to the Dictionary (dct) with the key (dctkey).
' Supports various key sequences, case and case insensitive key as well
' as adding items before or after an existing item.
' - When the key (dctkey) already exists the item is updated when it is
'   numeric or a string, else it is ignored.
' - When the dictionary (dct) is Nothing it is setup on the fly.
' - When dctmode = before or after dcttargetkey is obligatory and an
'   error is raised when not provided.
' - When the item's key is an object any dctmode other then by sequence
'   requires an object with a name property. If not the case an error is
'   raised.

' W. Rauschenberger, Berlin Mar 2020
' -----------------------------------------------------------------
    Const PROC = "DctAdd"
    Dim dctTemp As Dictionary
    Dim vKey    As Variant
    Dim bAdd    As Boolean

    On Error GoTo on_error
    
    If dct Is Nothing Then Set dct = New Dictionary
    
    If Not IsNumeric(dctkey) And TypeName(dctkey) <> "String" Then
        On Error Resume Next
        Debug.Print "Added object with name '" & dctkey.Name & "'"
        If Err.Number <> 0 _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "The key is neither a numeric value nor a string, nor an object with a name property!"
    End If
    
    With dct
        If .Count = 0 Or dctmode = dct_sequence Then ' the very first item is just added
            .Add dctkey, dctitem
            Exit Sub
        End If
        ' ----------------------------------------------------------------------
        ' Let's see whether the new key can be added directly after the last key
        ' ----------------------------------------------------------------------
        If IsNumeric(.Keys()(.Count - 1)) Or TypeName(.Keys()(.Count - 1)) = "String" _
        Then vKey = .Keys()(.Count - 1) _
        Else Set vKey = .Keys()(.Count - 1)
        
        Select Case dctmode
            Case dct_ascendingcasesensitive
                If DctAddKeyValue(dctkey) > DctAddKeyValue(vKey) Then
                    .Add dctkey, dctitem
                    Exit Sub                ' Done!
                End If
            Case dct_ascendingcaseingignored
                If LCase$(dctkey) > LCase$(DctAddKeyValue(vKey)) Then
                    .Add dctkey, dctitem
                    Exit Sub                ' Done!
                End If
            Case dct_descendingcasesensitive
                If DctAddKeyValue(dctkey) < DctAddKeyValue(vKey) Then
                    .Add dctkey, dctitem
                    Exit Sub                ' Done!
                End If
            Case dct_descendingcaseignored
                If LCase$(DctAddKeyValue(dctkey)) < LCase$(DctAddKeyValue(vKey)) Then
                    .Add dctkey, dctitem
                    Exit Sub                ' Done!
                End If
        End Select
    End With

    '~~ -------------------------------------------------------------------------
    '~~ Since the new key could not simply be added to the Dictionary it will be
    '~~ added, somewhere in between, before the very first or after the last key.
    '~~ -------------------------------------------------------------------------
    Set dctTemp = New Dictionary
    bAdd = True
    For Each vKey In dct
        With dctTemp
            If bAdd Then
                If dct.Exists(dctkey) Then
                    '~~ When the item is numeric or a string and the key already exists the item is updated
                    '~~ else ignored
                    If IsNumeric(dctitem) Or TypeName(dctitem) = "String" Then dct.Item(dctkey) = dctitem
                    Exit Sub
                End If
                Select Case dctmode
                    Case dct_ascendingcasesensitive
                        If DctAddKeyValue(vKey) > DctAddKeyValue(dctkey) Then
                            .Add dctkey, dctitem
                            bAdd = False ' Add done
                        End If
                    Case dct_ascendingcaseingignored
                        If LCase$(DctAddKeyValue(vKey)) > LCase$(DctAddKeyValue(dctkey)) Then
                            .Add dctkey, dctitem
                            bAdd = False ' Add done
                        End If
                    Case dct_addbefore
                        If DctAddKeyValue(vKey) = dcttargetkey Then
                            '~~> Add before dcttargetkey key has been reached
                            .Add dctkey, dctitem
                            bAdd = True
                        End If
                    Case dct_descendingcasesensitive
                        If DctAddKeyValue(vKey) < DctAddKeyValue(dctkey) Then
                            .Add dctkey, dctitem
                            bAdd = False ' Add done
                        End If
                    Case dct_descendingcaseignored
                        If LCase$(DctAddKeyValue(vKey)) < LCase$(DctAddKeyValue(dctkey)) Then
                            .Add dctkey, dctitem
                            bAdd = False ' Add done
                        End If
                End Select
            End If
            
            '~~> Transfer the existing item to the temporary dictionary
            .Add vKey, dct.Item(vKey)
            
            If dctmode = dct_addafter And bAdd Then
                If DctAddKeyValue(vKey) = dcttargetkey Then
                    ' ----------------------------------------
                    ' Just add when dctmode indicates add after,
                    ' and the vTraget key has been reached
                    ' ----------------------------------------
                    .Add dctkey, dctitem
                    bAdd = False
                End If
            End If
            
        End With
    Next vKey
    
    '~~> Return the temporary dictionary with the new item added
    Set dct = dctTemp
    Set dctTemp = Nothing

end_proc:
    Exit Sub

on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    MsgBox Prompt:=Err.Description, Title:="VB Runtime Error " & Err.Number & " in " & ErrSrc(PROC)
End Sub

Private Function DctAddKeyValue(ByVal dctkey As Variant) As Variant
' -----------------------------------------------------------------
' When dctkey is numeric or a string it is returned as is
' else when it is an object with a name property the name
' else a vbNullString. -----------------------------------------------------------------
    If IsNumeric(dctkey) Or TypeName(dctkey) = "String" Then
        DctAddKeyValue = dctkey
    Else
        On Error Resume Next
        DctAddKeyValue = dctkey.Name
        If Err.Number <> 0 Then DctAddKeyValue = vbNullString
    End If
End Function

End Sub
```

### Usage example
The _VBComponents_ of ThisWorkbook are added ordered in ascending sequence case sensitive (requires a reference to "") 
```vbscript
Private Sub DctAddExample()

   Dim dct As Dictionary
   Dim cbc As VbComponent
   
   For each vbc in This workbook.VBProject.VBComponents
      ' DctAdd dct, vbc, vbc.name
      ' would be the equivalent of
      ' dct.Add cbc, cbc.Name
      ' The items are added in entry
      DctAdd dct, vbc, vbc.name, dct_ascendingcasesensitive
      ' adds each item in ascending
      ' sequence         
   Next vbc
```