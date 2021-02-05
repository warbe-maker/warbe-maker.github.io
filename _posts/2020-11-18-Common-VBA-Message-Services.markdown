---
layout: post
title: Common VBA Message Service
subtitle: An alternative for the VBA MsgBox with less constraints, more options, and a better display
date: 2020-11-17
categories: vba common
---

## Why this alternative to the VBA MsgBox?
The alternative implementation  addresses many of the constraints of the VBA MsgBox - without re-implementing it yet to 100% however.

|The VBA MsgBox|The Common VBA Message Service|
|--------------|------------------------------|
| The message width and height is limited and cannot be altered | The&nbsp;maximum&nbsp;width and&nbsp;height&nbsp;is&nbsp;specified as&nbsp;a percentage of the screen&nbsp;size&nbsp; which&nbsp;defaults&nbsp;to: 80%&nbsp;width and  90%&nbsp;height (hardly ever used)|
| When a message exceeds the (hard to tell) size limit it is truncated | When the maximum size is exceeded a vertical and/or a horizontal scroll bar is applied
| The message is displayed with a proportional font | A message may (or part of it may) be displayed mono-spaced |
| Composing a fair designed message is time consuming and it is difficult to come up with a satisfying result | Up&nbsp;to&nbsp;4&nbsp; _Message&nbsp;Sections_ \*) each with an optional _Message Text Label_ and a _Monospaced_ option allow an appealing design without any extra  effort<br>\*) Adding an additional section is just a matter of the design and does not require any code change in the UserForm.  |
| The maximum reply _Buttons_ is 3 | Up to 7 reply _Buttons_ may be displayed in up to 7 reply _Button Rows_ in any order (=49 buttons in total) |
| The caption of the reply _Buttons_ is specified by a [value][1] which results in 1 to 3 reply _Buttons_ with corresponding untranslated! native English captions | The caption of the reply _Buttons_ may be specified by the [VB MsgBox values][1] **and** additionally by any multi-line text (see [Syntax of the _buttons_ argument](#syntax-of-the-buttons-argument) |
| Specifying the default button | possible |
| Display of an alert image (?, !, etc.) | (yet) not implemented |

## The _Dsply_ and the _Box_ service
- _Dsply_ displays an optionally structured message, i.e one which may consist of 4 sections, each with an optional label
- _Box_ displays a message analogously to the VBA MsgBox _Prompt_ argument, i.e. a single string message of any length (up to 1GB respectively), with the full buttons flexibility however.<br><br>Both services:
- displays up to 49 free configurable return buttons in up to 7 rows
- intelligently considers the space required for the displayed elements: title, message, and buttons
- displays a horizontal and/or vertical scroll-bar when applicable/required
- Waits for the user to click a button, and provides a return variant indicating which button the user had  clicked.

Example of an error message using an additional free text reply button:<br>
![](../Assets/ErrrorMessageWithResumeButton.png)
![](/Assets/ErrrorMessageWithResumeButton.png)

Example for a text wich spans mor than the specified maximum message window width:<br>
![](../Assets/ExecutionTraceDetailed.png)
![](/Assets/ExecutionTraceDetailed.png)<br>

The _Dsply_ and the _Box_ service have these named arguments:

|    Part              | Description                    |
| -------------------- |------------------------------- |
| msg_title            | Obligatory. String expression displayed in the title bar of the dialog box. |
| msg                  | _Box_ service:<br>Optional string expression, when omitted only the button(s) are displayed<br> _Dsply_ service:<br>Obligatory, User defined type _tMsg_, no message length limit. When the argument remains empty, i.e. a type tMsg variable is provided without any content, only the buttons are displayed. Message lines may be separated by using a carriage return character (vbCr or Chr(13), a linefeed character (vbLf or Chr(10)), or carriage return - linefeed character combination (vbCrLf or Chr(13) & Chr(10)) between each line.<br>_Box_ service:<br>Optional, String expression of any length (up to 1 GB), when not provided only the specified buttons are displayed. The message string may consist of any number of lines, separated by means of: vbCr or Chr(13), vbLf or Chr(10), or vbCrLf Chr(13) & Chr(10)).|
| msg_monospaced       | _Box_ service only, Optional, Boolean expression, defaults to False, displays the message monospaced when True, adjusts the message window width to the longest line, displays a horizontal Scroll-Bart when exceeded|
| msg_buttons          | Optional. Variant expression, defaults to vbOkOnly when omitted, either<br>- a string of comma delimited items,<br>- a collection of items as provided by the [Buttons](#the-buttons-service) service,<br>- a dictionary of variant items.<br> Each item may be a<br>- a button's caption string<br>- a valid [VBA MsgBox value][1]<br>- a row break indication (vbLf, vbCr, or vbCrLf).|
| msg_button_default   | Optional, Variant, defaults 1. May be the sequence number or the caption string of the button |
| msg_returnindex      | Optional, Boolean, False when omitted    |
| msg_min_width        | Optional, Long, defaults to 300 pt when omitted, cannot be less than 200 pt |
| msg_max_width        | Optional, Long, defaults to 80% of the screen size when omitted |
| msg_max_height       | Optional, Long, defaults to 70% of the screen size when omitted |
| msg_min_button_width | Optional, Long, defaults to 70 pt when omitted |

## The _Buttons_ service
The _mMsg.Buttons_ service returns a Collection of items provided via a ParamArray argument. Each of the items may either be a string expression, a valid [VBA MsgBox Buttons argument value][1], or a row break indication (vbLf, vbCr, or vbCrLf). When more than 7 buttons items are provided without a row break indicator one is in inserted by the service. Any invalid item is ignored and any button specification which exceeds 7 rows by 7 buttons (= 47 buttons) is ignored.

The _Buttons_ service has this syntax:
`Buttons(item-1[, item-2][, item-3] ...`

The _Buttons_ service has this named argument:
|    Part     | Description                    |
| ----------- |------------------------------- |
| msg_buttons | Obligatory, ParamArray, each item either specifies a button or a row break (vbLf). 

## The _fMsg_ UserForm
The UserForm may be used [directly](#direct-usage-of-the-fmsg-userform)  but with significant less comfort compared with the _Dsply_ and the _Box_ service.

The _fMsg_ service has the following Properties (usually covered by the _Dsply_ and the _Box_ service):

| Property      | Meaning |
|---------------|---------|
| _MsgTitle_    | Mandatory. String expression. Applied in the message window's handle bar|
| _Msg_         | Optional. User defined type _tMsg_. Structure of the UserForm's message area. May alternatively be used to the below properties _MsgLabel_, _MsgText_, and _MsgMonoSpaced_ to pass a complete message.<br>See .... |
| _MsgLabel(n)_ | Optional. String expression with _n__ as a numeric expression 1 to 4. Applied as a descriptive label above a below message text. Not displayed (even when provided) when no corresponding _MsgText_ is provided |
| _MsgText(n)_  | Optional.String expression with _n__ as a numeric expression 1 to 4). Applied as message text of section _n_.|
| _MsgMonospaced(n)_ | Optional. Boolean expression with _n__ as a numeric expression 1 to 4). Defaults to False when omitted. When True, the text in section _n_ is displayed mono-spaced.|
| _MsgButtons_  | Optional. Defaults to vbOkOnly when not provided (see [The Buttons service](#the-buttons-service) and the [_dsply\_buttons_](#the-dsply-buttons-argument) argument.|
| _ReplyValue_  | Read only. The clicked button's caption string or [value][1]. When there is more than one button the form is unloaded when the clicked buttons value is fetched. When there is just one button this value will not be available since the form is immediately unloaded with the button click.|
| _ReplyIndex_  | Read only. The clicked button's index. When there is more than one button the form is unloaded when the clicked button's index is fetched. When there is just one button this value will not be available since the form is immediately unloaded with the button click. |

See [Additional properties for advanced usage](<Implementation.md#public-properties-for-advanced-usage-of-the-message-form>) to create application specific messages.

## Installation
1. Download the UserForm  [fMsg.frm][2] and   [fMsg.frx][3]
1. Import _fMsg.frm_
1. Download and import [mMsg.bas][4]
1. In the VBE add a Reference to "Microsoft Scripting Runtime"


## Usage
### Using the _Box_ service
The code example directly uses the _Box_ service. Not a typical use for the _Box_ service which is meant to be pretty much like the VBA.MsgBox service but used here to show the difference to the _Dsply_ service.
```
    Const BTTN_1 = "Caption Button 1"
    Const BTTN_2 = "Caption Button 2"
    Const BTTN_3 = "Caption Button 3"
    Const BTTN_4 = "Caption Button 4"
    Const BTTN_5 = "Caption Button 5"
    Const BTTN_6 = "Caption Button 6"
    Const BTTN_7 = "Caption Button 7"
    
    Select Case mMsg.Box(msg_title:="Any title" _
                       , msg:="The message"
                       , msg_buttons:=msg_buttons:=mMsg.Buttons(BTTN_1, BTTN_2, BTTN_3, vbLf, BTTN_4, BTTN_5, BTT_6, vbLf, BTTN_7)
        Case BTTN_1 ...
        Case BTTN_1 ...
        Case BTTN_1 ...
        Case BTTN_1 ...
    End Select
```

### Using the _Dsply_ service

```
    Const BTTN_1 = "Caption Button 1"
    Const BTTN_2 = "Caption Button 2"
    Const BTTN_3 = "Caption Button 3"
    Const BTTN_4 = "Caption Button 4"
    Const BTTN_5 = "Caption Button 5"
    Const BTTN_6 = "Caption Button 6"
    Const BTTN_7 = "Caption Button 7"
    Dim sMsg As tMsg
    sMsg.Section(1).sLabel = "Any label 1"
    sMsg.Section(1).sText = "Any section text 1"
    sMsg.Section(2).sLabel = "Any label 2"
    sMsg.Section(2).sText = "Any section text 2"
    sMsg.Section(3).sLabel = "Any label 3"
    sMsg.Section(3).sText = "Any section text 3"
    
    Select Case mMsg.Dsply(msg_title:="Any title" _
                         , msg:=sMsg _
                         , msg_buttons:=mMsg.Buttons(BTTN_1, BTTN_2, BTTN_3, vbLf, BTTN_4, BTTN_5, BTT_6, vbLf, BTTN_7)
        Case BTTN_1 ...
        Case BTTN_1 ...
        Case BTTN_1 ...
        Case BTTN_1 ...
    End Select
```
displays:<br>
![](../Assets/TestDsply.png)
![](/Assets/TestDsply.png)<br>

### Direct usage of the _fMsg_ UserForm
The below example may not look very attractive because the _fMsg_ UserForm exposes an enormous amount of inwards. Checkout what the _[Dsply](#using-the-dsply-service)_ and the _[Box](#using-the-box-service)_ service offers before using the _fMsg_ [UserForm's properties](#the-fmsg-userform) directly.

```vbs
Public Sub DemoUserForm()
          
    With fMsg
        .MsgTitle = "Any title"
        .MsgText(1) = "Any message"
        .MsgButtons = vbYesNoCancel
        .Setup
        .Show
    End With
    Select Case fMsg.ReplyValue ' obtaining it unloads the form !
            Case vbYes:     MsgBox "Button ""Yes"" clicked"
            Case vbNo:      MsgBox "Button ""No"" clicked"
            Case vbCancel:  MsgBox "Button ""Cancel"" clicked"
    End Select
End Sub
```

Displays:<br>
![](../Assets/AlternativeMsgBoxFirstStepMessage.png)
![](/Assets/AlternativeMsgBoxFirstStepMessage.png)<br>


### Proportional versus Mono-Spaced

The effect it has when a text in a section is specified mono-spaced (the default is proportional-spaced) is demonstrated by the second example in the [Services](#services) section above. Because the section specific message text is ++not++ "wrapped"  but The message windows width is adjusted up to the maximum width specified. In case even that's not enough a horizontal scroll-bar is displayed.

[1]:https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings
[2]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm
[3]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frx
[4]:https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mMsg.bas