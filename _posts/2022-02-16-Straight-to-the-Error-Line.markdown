---
layout: post
title: Straight to the Error Line
date:     2022-02-16
modified: 2023-06-17
categories: vba common error handling
---
An error message with a debugging option enabling to go straight to the error line without knowing the error line number (as usual there aren't any).
<!--more-->

## Preface
I hardly eve spend more than a second. When an error message is displayed I often not even read it but go straight to the error line. This increases the debugging performance as much as possible.

## What does the trick?
I've seen several ways described on the web and decided for the following which I found most elegant.

### Preparing the module
At first we need a function which provides an unambiguous procedure name by prefixing it with the module's name.
```vb
Private Function ErrSrc(ByVal proc_name As String) As String
    ErrSrc = "mDemo." & proc_name
End Function
```
Next we need and error message function which provides an extra button.
```vb
Private Function ErrMsg(ByVal proc_name As String) As Variant
    ErrMsg = VBA.MsgBox(Prompt:="Error: " & Err.Description & vbLf & vbLf & _
                                "Yes: Resume Error Line" & vbLf & _
                                "No : Terminate procedure" _
                      , Title:="Error " & Err.Number & " in " & proc_name _
                      , Buttons:=vbYesNo + vbCritical)
End Function
```
### Preparing procedures
```vb
Private Sub TestProc()
    Const PROC = "TestProc"
    
    On Error GoTo eh
    '
    TestTestProc    ' this one will raise the error
    '
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub TestTestProc()
    Const PROC = "TestTestProc"
    
    On Error GoTo eh
    Dim wb As Workbook
    
    Debug.Print wb.Name ' will raise a VB Runtime error no 91

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
```
Executing TestProc displays the following message:<br>
![](../Assets/StraightToTheErrorLine.png)<br>
![](/Assets/StraightToTheErrorLine.png)<br>(sorry for the German) that's it! Yes/Ja and twice the F8 keystroke and voila! there we are.

> One may think that's a lot of code lines but I can assure everyone: When an error message like this is displayed the debugging option will be perceived like a godsend.

However, I'd prefer an error message like this:<br>
![](../Assets/StraightToTheErrorLineOptimum.png)<br>
![](/Assets/StraightToTheErrorLineOptimum.png)<br>It almost invites for debugging and that's why I have implemented this kind of error message with all my VB-Projects and all my  _[Common VBA Components][1]_. The implementation only requires:

- 4 procedures I use in each module (AppErr, BoP, EoP, and ErrMsg), the last one replacing the simple ErrMsg function above)
- 2 of my _Common VBA Components_ ([Common VBA Error Services][2], and _Common VBA Message Service_) the latter installed together with the first one all well explained in the README) 
- 2 Conditional Compile Arguments to indicate debugging and the use of the common components.

> Don't shy away when approaching the public ***GitHub repos*** with the links above - as I did for the first time, thinking that that's stuff beyond my level of knowledge. The README provides all the means (clicks) for downloading the components. And the procedures which are copied into your modules are the interfaces. So there is no need to know any further details. 

The [StraightToTheErrorLine.xlsm][4] Workbook has all this included and may be used straight as a demonstration. You will have to change the signature of the VB-Project or the macro settings however (I prefer the first one).

Anything looking strange? Maybe the use of the Conditional Compile Argument? The answer may be:
> All my _Common VBA Components_ are prepared to function as autonomous as possible (download, import, use) while still integrating with my personal 'standard' VB-Project design. See [Conflicts with personal and public _Common Components_][3] for more details.

## Contribution
Any kind of contribution is welcome. I apologize for the fact that logging in to GitHub may be an all but insurmountable obstacle. It is an appropriate means for keeping away spammers.

 [1]:https://warbe-maker.github.io/vba/common/2021/02/19/Common-VBA-Components.html
 [2]:https://github.com/warbe-maker/VBA-Error
 [3]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
 [4]:https://gitcdn.link/cdn/warbe-maker/Straight-to-the-error-line-demo/master/StraightToTheErrorLine.xlsm