---
layout: post
title: An "Alternative VBA MsgBox"
subtitle: Less constraints, more options, better display
date: 2020-10-19 16:00 +0200
categories: vba common
---

In this post

[Service](#service)<br>
[Why just another, alternative MsgBox](#why-just-another-alternative-msgbox)<br>
[Installation](#installation)<br>
[Properties of the _fMsg_ UserForm](#properties-of-the-fmsg-userform)<br>
[Usage](#usage)<br>
&nbsp;&nbsp;&nbsp;[Direct usage of the _fMsg_ UserForm](#directly-using-the-fmsg-userform)<br>
&nbsp;&nbsp;&nbsp;[Usage via a general purpose _Msg_ function](#usage-via-a-general-purpose-msg-function)<br>
&nbsp;&nbsp;&nbsp;[Proportional versus Mono-Spaced](#proportional-versus-mono-spaced)<br>
&nbsp;&nbsp;&nbsp;[Additional properties for advanced usage](#additional-properties-for-advanced-usage)


### Service
A message box which intelligently considers the space required for the displayed elements title, message, and buttons, waiting for the user to click a button, and providing a variant indicating which button the user had  clicked.
![Example of an error message using an additional free text reply button](../Assets/ErrrorMessageWithResumeButton.png)
![Example of an error message using an additional free text reply button](/Assets/ErrrorMessageWithResumeButton.png)

![Example for a text wich spans mor than the specified maximum message window width](../Assets/ExecutionTraceDetailed.png)
![Example for a text wich spans mor than the specified maximum message window width](/Assets/ExecutionTraceDetailed.png)

### Why just another, alternative MsgBox?
The alternative implementation  addresses many of the constraints of the VBA MsgBox - without re-implementing it to 100%.

|The VBA MsgBox|The Alternative|
|--------------|---------------|
| The message width and height is limited and cannot be altered | The&nbsp;maximum&nbsp;width and&nbsp;height&nbsp;is&nbsp;specified as&nbsp;a percentage of the screen&nbsp;size&nbsp; which&nbsp;defaults&nbsp;to: 80%&nbsp;width and  90%&nbsp;height (hardly ever used)|
| When a message exceeds the (hard to tell) size limit it is truncated | When the maximum size is exceeded a vertical and/or a horizontal scroll bar is applied
| The message is displayed with a proportional font | A message may (or part of it may) be displayed mono-spaced |
| Composing a fair designed message is time consuming and it is difficult to come up with a satisfying result | Up&nbsp;to&nbsp;4&nbsp; _Message&nbsp;Sections_ \*) each with an optional _Message Text Label_ and a _Monospaced_ option allow an appealing design without any extra  effort<br>\*) Adding an additional section is just a matter of the design and does not require any code change in the UserForm.  |
| The maximum reply _Buttons_ is 3 | Up to 7 reply _Buttons_ may be displayed in up to 7 reply _Button Rows_ in any order (=49 buttons in total) |
| The caption of the reply _Buttons_ is specified by a [value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) which results in 1 to 3 reply _Buttons_ with corresponding untranslated! native English captions | The caption of the reply _Buttons_ may be specified by the [VB MsgBox values](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) **and** additionally by any multi-line text (see [Syntax of the _buttons_ argument](#syntax-of-the-buttons-argument) |
| Specifying the default button | (yet) not implemented |
| Display of an alert image (?, !, etc.) | (yet) not implemented |

## Installation
1. Download the UserForm  [fMsg.frm](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm) and   [fMsg.frx](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frx)
2. Import _fMsg.frm_
3. In the VBE add a Reference to "Microsoft Scripting Runtime"
5. Copy the following into a standard module or alternatively [download the _mMsg_ module](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mMsg.bas) and import it. It has all the required resources on board:<br>
```
Public Enum StartupPosition             ' --------------------
    Manual                              ' Used to position
    CenterOwner                         ' the message window
    CenterScreen                        ' horizontally and
    WindowsDefault                      ' vertically centered
End Enum                                ' -------------------

Public Type tMsgSection                 ' ---------------------
       sLabel As String                 ' Structure of the
       sText As String                  ' UserForm's message
       bMonspaced As Boolean            ' area which consists
End Type                                ' of 4 message sections
Public Type tMsg                        ' Attention: 4 is a
       section(1 To 4) As tMsgSection   ' design constant!
End Type                                ' ---------------------

```

### Properties of the _fMsg_ UserForm

| Property | Meaning |
|----------|---------|
| _MsgTitle_| Mandatory. String expression. Applied in the message window's handle bar|
| _Msg_     | Optional. User defined type _tMessage_. Structure of the UserForm's message area. May alternatively be used to the below properties _MsgLabel_, _MsgText_, and _MsgMonoSpaced_ be used to pass a complete message.<br>See .... |
| _MsgLabel(n)_ | Optional. String expression with _n__ as a numeric expression 1 to 4. Applied as a descriptive label above a below message text. Not displayed (even when provided) when no corresponding _MsgText_ is provided |
| _MsgText(n)_ | Optional.String expression with _n__ as a numeric expression 1 to 4). Applied as message text of section _n_.|
| _MsgMonospaced(n)_ | Optional. Boolean expression with _n__ as a numeric expression 1 to 4). Defaults to False when omitted. When True, the text in section _n_ is displayed mono-spaced.|
| _MsgButtons_ | Optional. Defaults to vbOkOnly.<br>A MsgBox buttons value,<br>a comma delimited String expression,<br>a Collection,<br>or a dictionary,<br>with each item specifying a displayed command button's caption or a button row break (vbLf, vbCr, or vbCrLf)|
| _ReplyValue_ | Read only. The clicked button's caption string or [value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>). When there is more than one button the form is unloaded when the clicked buttons value is fetched. When there is just one button this value will not be available since the form is immediately unloaded with the button click.|
| _ReplyIndex_ | Read only. The clicked button's index. When there is more than one button the form is unloaded when the clicked button's index is fetched. When there is just one button this value will not be available since the form is immediately unloaded with the button click. |

See [Additional properties for advanced usage](<Implementation.md#public-properties-for-advanced-usage-of-the-message-form>) to create application specific messages.

## Usage
Before start using the message form have a look at the [UserForm's properties](#properties-of-the-fmsg-userform).
Either continue with [Usage step by step](#usage-step-by-step) or go directly to using the prepared [Using the message form via an nterface](#Interfaces).  

### Direct usage of the _fMsg_ UserForm
It's not as comfortable as possible but appropriate to understand its use 
```vbs
Public Sub DemoDirect()
          
    With fMsg
        .MsgTitle = "Any title"
        .MsgText(1) = "Any message"
        .MsgButtons = vbYesNoCancel
        .Setup
        .Show
        Select Case .ReplyValue ' obtaining it unloads the form !
            Case vbYes:     MsgBox "Button ""Yes"" clicked"
            Case vbNo:      MsgBox "Button ""No"" clicked"
            Case vbCancel:  MsgBox "Button ""Cancel"" clicked"
        End Select
   End With
End Sub
```

Displays:
![](../Assets/AlternativeMsgBoxFirstStepMessage.png)
![](/Assets/AlternativeMsgBoxFirstStepMessage.png)

The above example seems not being worth using the alternative. However, when encapsulated in a function which exposes all relevant matter via arguments things look much better. Copy the following may be copied into a standard module or [download the _mMsg_ module](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mMsg.bas) and import it. It has all the required resources on board:
```vbs
Public Function Dsply(ByVal dsply_title As String, _
                      ByRef dsply_message As tMsg, _
             Optional ByVal dsply_buttons As Variant = vbOKOnly, _
             Optional ByVal dsply_returnindex As Boolean = False, _
             Optional ByVal dsply_min_width As Long = 300, _
             Optional ByVal dsply_max_width As Long = 80, _
             Optional ByVal dsply_max_height As Long = 70, _
             Optional ByVal dsply_min_button_width = 30) As Variant
' ------------------------------------------------------------------------------------
' Common VBA Message Display: A service using the Common VBA (alternative) MsgBox.
' See: https://warbe-maker.github.io/vba/common/2020/10/19/Alternative-VBA-MsgBox.html
'
' W. Rauschenberger, Berlin, Nov 2020
' ------------------------------------------------------------------------------------

    With fMsg
        .MaxFormHeightPrcntgOfScreenSize = dsply_max_height ' percentage of screen size
        .MaxFormWidthPrcntgOfScreenSize = dsply_max_width   ' percentage of screen size
        .MinFormWidth = dsply_min_width                     ' defaults to 300 pt. the absolute minimum is 200 pt
        .MinButtonWidth = dsply_min_button_width
        .MsgTitle = dsply_title
        .Msg = dsply_message
        .MsgButtons = dsply_buttons
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form improves the performance significantly  ||
        '|| and avoids any flickering message window with its setup.             ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        .Setup '                                                                 ||
        '+------------------------------------------------------------------------+
        .show
    End With
    
    ' -----------------------------------------------------------------------------
    ' Obtaining the reply value/index is only possible when more than one button is
    ' displayed! When the user had a choice the form is hidden when the button is
    ' pressed and the UserForm is unloade when the return value/index (either of
    ' the two) is obtained!
    ' -----------------------------------------------------------------------------
    If dsply_returnindex Then Dsply = fMsg.ReplyIndex Else Dsply = fMsg.ReplyValue

End Function


```
The _Dsply_ function syntax has these named arguments:

|    Part    | Description|
| ---------- |----------- |
| msg_title  | Obligatory. String expression displayed in the title bar of the dialog box. |
| msg_message| Obligatory, User defined type _tMessage_, no length limit. When the maximum height or width is exceeded a vertical and/or horizontal scrollbars is displayed. Lines may be separated by using a carriage return character (vbCr or Chr(13), a linefeed character (vbLf or Chr(10)), or carriage return - linefeed character combination (vbCrLf or Chr(13) & Chr(10)) between each line.|
| msg_buttons| Optional. Defaults to vbOkOnly when omitted. Variant expression, either a [VB MsgBox value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>), a comma delimited string, a collection of string expressions, or a dictionary of string expressions. In case of a string, a collection, or a dictionary, each item either specifies a button's caption (up to 7) or a reply button row break (vbLf, vbCr, or vbCrLf). |

#### Syntax of the _buttons_ argument
```
msg_buttons:=string|value[, rowbreak][, button2][, rowbreak][, button3][, rowbreak][, button4][, rowbreak][, button5][, rowbreak][, button6][, rowbreak][, button7]
```

**string**, **button2** ... **button7**: captions for the buttons 1 to 7<br>
value: the VB MsgBox argument for 1 to 3 buttons all in one row<br>
rowbreak: vbLf or Chr(10). Indicates that the next button is displayed in the row below

Displaying a message with this function may either look pretty much the same as using the VBA MsgBox:

```vbs

```

or may use the full flexibility of the message form when displaying a message with 3 sections, each with a label and 7 reply buttons ordered in rows 3-3-1

To keep this example simple the button's value/caption text is used as the return value.
```vbs
Public Sub Test_Msg()
' ---------------------------------------------------------
' Displays a message with 3 sections, each with a label and
' 7 reply buttons ordered in rows 3-3-1
' ---------------------------------------------------------
    Const B1 = "Caption Button 1"
    Const B2 = "Caption Button 2"
    Const B3 = "Caption Button 3"
    Const B4 = "Caption Button 4"
    Const B5 = "Caption Button 5"
    Const B6 = "Caption Button 6"
    Const B7 = "Caption Button 7"
    Dim tMsg    As tMessage                         ' structure of the message
    Dim cll     As New Collection                   ' specification of the displayed buttons
    
    cll.Add B1
    cll.Add B2
    cll.Add B3
    cll.Add vbLf ' row break
    cll.Add B4
    cll.Add B5
    cll.Add B6
    cll.Add vbLf ' row break
    cll.Add B7
       
    With tMsg.Section(1)
        .sLabel = "Any label 1"
        .sText = "Any section text 1"
    End With
    With tMsg.Section(2)
        .sLabel = "Any label 2"
        .sText = "Any section 2 text"
        .bMonspaced = True ' Just to demonstrate
    End With
    With tMsg.Section(3)
        .sLabel = "Any label 3"
        .sText = "Any section text 3"
   End With
       
   Select Case Msg(msg_title:="Any title", _
                   msg_message:=tMsg, _
                   msg_buttons:=cll)
        Case B1: Debug.Print "Button with caption """ & B1 & """ clicked"
        Case B2: Debug.Print "Button with caption """ & B2 & """ clicked"
        Case B3: Debug.Print "Button with caption """ & B3 & """ clicked"
        Case B4: Debug.Print "Button with caption """ & B4 & """ clicked"
        Case B5: Debug.Print "Button with caption """ & B5 & """ clicked"
        Case B6: Debug.Print "Button with caption """ & B6 & """ clicked"
        Case B7: Debug.Print "Button with caption """ & B7 & """ clicked"
    End Select
   
End Sub
```

### Proportional versus Mono-Spaced

With
```vbs
With tMsg.Section(n)
        .sLabel = "..."
        .sText = "......."
        .bMonospaced_ = True
End With
```

or when the UserForm is directly used with:
```vbs
   With fMsg
        .ApplTitle = "Any title"
        .ApplText(1) = "Any message"
        .Monospaced(1) = True
        .ApplButtons = vbYesNoCancel
        .Setup
        .Show
   End With
```

the section specific message text  text is ++not++  "wrapped" and thus the width of the _Message Form_ is determined by the longest text line (up to the _Maximum Form Width_ specified). When the maximum width is exceeded a vertical scroll bar is applied.<br>Note: The title and the broadest _Button Row_ May still determine an even broader final _Message Form_.

When Monospaced is omitted it defaults False and the message text's width is determined by the title's length and the buttons' width.

## Additional properties for advanced usage
| Property | Description |
| -------- | ----------- |
|          |             |
|          |             |