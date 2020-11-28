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
| The caption of the reply _Buttons_ is specified by a [value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) which results in 1 to 3 reply _Buttons_ with corresponding untranslated! native English captions | The caption of the reply _Buttons_ may be specified by the [VB MsgBox values](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) **and** additionally by any multi-line text (see [Syntax of the _buttons_ argument](#syntax-of-the-buttons-argument) |
| Specifying the default button | (yet) not implemented |
| Display of an alert image (?, !, etc.) | (yet) not implemented |

## The display service _mMsg.Dsply_
The service
- Displays a message which may consist of 4 sections, each with an optional label
- Displays up to 49 free configurable return buttons in up to 7 rows
- Intelligently considers the space required for the displayed elements: title, message, and buttons
- Displays a horizontal and/or vertical scroll-bar when applicable/required
- Waits for the user to click a button, and provides a return variant indicating which button the user had  clicked.
![Example of an error message using an additional free text reply button](../Assets/ErrrorMessageWithResumeButton.png)

![Example of an error message using an additional free text reply button](/Assets/ErrrorMessageWithResumeButton.png)

![Example for a text wich spans mor than the specified maximum message window width](../Assets/ExecutionTraceDetailed.png)
![Example for a text wich spans mor than the specified maximum message window width](/Assets/ExecutionTraceDetailed.png)

The _Dsply_ service has these named arguments:

|    Part                | Description                    |
| ---------------------- |------------------------------- |
| dsply_title            | Obligatory. String expression displayed in the title bar of the dialog box. |
| dsply_msg              | Obligatory, User defined type _tMsg_, no message length limit. When the argument remains empty, i.e. a type tMsg variable is provided without any content, only the buttons are displayed. Message lines may be separated by using a carriage return character (vbCr or Chr(13), a linefeed character (vbLf or Chr(10)), or carriage return - linefeed character combination (vbCrLf or Chr(13) & Chr(10)) between each line.  |
| dsply_msg_string | Optional, String expression, confirms with the _Prompt_ argument of the VBA MsgBox, may be used when only one message string with no label is to be displayed|
| dsply_msg_string_monospaced| Optional, Boolean expression, defaults to False, displays the message monospaced when True|
| dsply_buttons          | Optional. Defaults to vbOkOnly when omitted. Variant expression, either a [VB MsgBox value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>), a comma delimited string, a collection of string expressions, or a dictionary of string expressions. In case of a string, a collection, or a dictionary, each item either specifies a button's caption (up to 7) or a reply button row break (vbLf, vbCr, or vbCrLf). |
| dsply_returnindex      | Optional, Boolean, False when omitted                                |
| dsply_min_width        | Optional, Long, defaults to 300 pt when omitted, cannot be less than 200 pt |
| dsply_max_width        | Optional, Long, defaults to 80% of the screen size when omitted |
| dsply_max_height       | Optional, Long, defaults to 70% of the screen size when omitted |
| dsply_min_button_width | Optional, Long, defaults to 70 pt when omitted   |

## The Box service _mMsg.Box_
The service
- Displays a one-string message (analogous to the VBA MsgBox Prompt argument) of any length
- Displays up to 49 free configurable return buttons in up to 7 rows
- Intelligently considers the space required for the displayed elements: title, message, and buttons
- Displays a horizontal and/or vertical scroll-bar when applicable/required
- Waits for the user to click a button, and provides a return variant indicating which button the user had  clicked.

The _Box_ service has these named arguments:

|    Part                | Description                    |
| ---------------------- |------------------------------- |
| dsply_title            | Obligatory. String expression displayed in the title bar of the dialog box. |
| dsply_msg              | Optional, String expression of any length (up to 1 GB), when not provided only the specified buttons are displayed. The message string may consist of any number of lines, separated by means of: vbCr or Chr(13), vbLf or Chr(10), or vbCrLf Chr(13) & Chr(10)).  |
| dsply_msg_monospaced| Optional, Boolean expression, defaults to False, when True the message is displayed mon-spaced.|
| dsply_buttons          | Optional. Variant expression, defaults to vbOkOnly when omitted. |
| dsply_returnindex      | Optional, Boolean, False when omitted                                |
| dsply_min_width        | Optional, Long, defaults to 300 pt when omitted, cannot be less than 200 pt |
| dsply_max_width        | Optional, Long, defaults to 80% of the screen size when omitted |
| dsply_max_height       | Optional, Long, defaults to 70% of the screen size when omitted |
| dsply_min_button_width | Optional, Long, defaults to 70 pt when omitted   |

## The Buttons service
The _mMsg.Buttons_ service returns a Collection of items provided via a ParamArray argument. each of the items may be:
- a string expression
- a valid [VBA MsgBox Buttons argument value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>)
- a row break indication (vbLf, vbCr, or vbCrLf). 

When more than 7 items are provided without a row break indicator one is in inserted by the service. Any invalid item is ignored and any specification which exceeds 7 rows or 47 buttons is ignored.

## The _dsply\_buttons_ argument

This argument of the Box, the Dsply, and the Buttons service is a variant expression which may be:
- a string of comma delimited items, 
- a collection of variant items as provided by the [Buttons](#the-buttons-service) service, 
- a dictionary of variant items

Each item may be :
- a button's caption string
- a valid [VBA MsgBox value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>)
- a row break indication (vbLf, vbCr, or vbCrLf). 

## The UserForm service _fMsg_
The UserForm may be used [directly](#direct-usage-of-the-fmsg-userform)  but with significant less comfort compared with the _Dsply_ and the _Box_ service.

The UserForm service has the following Properties:

| Property      | Meaning |
|---------------|---------|
| _MsgTitle_    | Mandatory. String expression. Applied in the message window's handle bar|
| _Msg_         | Optional. User defined type _tMsg_. Structure of the UserForm's message area. May alternatively be used to the below properties _MsgLabel_, _MsgText_, and _MsgMonoSpaced_ to pass a complete message.<br>See .... |
| _MsgLabel(n)_ | Optional. String expression with _n__ as a numeric expression 1 to 4. Applied as a descriptive label above a below message text. Not displayed (even when provided) when no corresponding _MsgText_ is provided |
| _MsgText(n)_  | Optional.String expression with _n__ as a numeric expression 1 to 4). Applied as message text of section _n_.|
| _MsgMonospaced(n)_ | Optional. Boolean expression with _n__ as a numeric expression 1 to 4). Defaults to False when omitted. When True, the text in section _n_ is displayed mono-spaced.|
| _MsgButtons_  | Optional. Defaults to vbOkOnly when not provided (see [The Buttons service](#the-buttons-service) and the [_dsply\_buttons_](#the-dsply-buttons-argument) argument.|
| _ReplyValue_  | Read only. The clicked button's caption string or [value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>). When there is more than one button the form is unloaded when the clicked buttons value is fetched. When there is just one button this value will not be available since the form is immediately unloaded with the button click.|
| _ReplyIndex_  | Read only. The clicked button's index. When there is more than one button the form is unloaded when the clicked button's index is fetched. When there is just one button this value will not be available since the form is immediately unloaded with the button click. |

See [Additional properties for advanced usage](<Implementation.md#public-properties-for-advanced-usage-of-the-message-form>) to create application specific messages.

## Installation
1. Download the UserForm  [fMsg.frm](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frm) and   [fMsg.frx](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/fMsg.frx)
2. Import _fMsg.frm_
3. In the VBE add a Reference to "Microsoft Scripting Runtime"
4. Download and import [mMsg.bas](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mMsg.bas)


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

The above example seems not being worth using the alternative. Above all when using the _fMsg_ UserForm directly one may not see the forest for the trees because the UserForm exposes an enormous amount od inwards. The following is an appropriate interface which has only thos arguments which matter. The function may be copied to any standard module. [Downloading the _mMsg_ module](https://gitcdn.link/repo/warbe-maker/VBA-MsgBox-alternative/master/mMsg.bas) and importing it is an option. It has all the required resources on board:

```vbs
Attribute VB_Name = "mMsg"
Option Explicit
' -----------------------------------------------------------------------------------------
' Standard Module mMsg Interface for the Common VBA "Alternative" MsgBox (fMsg UserForm)
'
' Methods: Dsply  Exposes all properties and methods for the display of any kind of message
'
' W. Rauschenberger, Berlin Nov 2020
' -----------------------------------------------------------------------------------------
Public Type tMsgSection                 ' ---------------------
       sLabel As String                 ' Structure of the
       sText As String                  ' UserForm's message
       bMonspaced As Boolean            ' area which consists
End Type                                ' of 4 message sections
Public Type tMsg                        ' Attention: 4 is a
       section(1 To 4) As tMsgSection   ' design constant!
End Type                                ' ---------------------

Public Function Dsply(ByVal dsply_title As String, _
                      ByRef dsply_msg_type As tMsg, _
             Optional ByVal dsply_msg_strng As String = vbNullString, _
             Optional ByVal dsply_msg_strng_monospaced As Boolean = False, _
             Optional ByVal dsply_buttons As Variant = vbOKOnly, _
             Optional ByVal dsply_returnindex As Boolean = False, _
             Optional ByVal dsply_min_width As Long = 300, _
             Optional ByVal dsply_max_width As Long = 80, _
             Optional ByVal dsply_max_height As Long = 70, _
             Optional ByVal dsply_min_button_width = 70) As Variant
' -------------------------------------------------------------------------------------
' Common VBA Message Display: A service using the Common VBA Message Form as an
' alternative MsgBox.
' Note: In case there is only one single string to be displayed the argument
'       dsply_msg_type will remain unused while the messag is provided via the
'       dsply_msg_strng and dsply_msg_strng_monospaced arguments instead.
'
' See: https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Form.html
'
' W. Rauschenberger, Berlin, Nov 2020
' -------------------------------------------------------------------------------------
    Dim i As Long
    
    With fMsg
        .MaxFormHeightPrcntgOfScreenSize = dsply_max_height ' percentage of screen size
        .MaxFormWidthPrcntgOfScreenSize = dsply_max_width   ' percentage of screen size
        .MinFormWidth = dsply_min_width                     ' defaults to 300 pt. the absolute minimum is 200 pt
        .MinButtonWidth = dsply_min_button_width
        .MsgTitle = dsply_title
        If dsply_msg_strng <> vbNullString Then
            '~~ The message os provided as a simple string
            .MsgText(1) = dsply_msg_strng
        Else
            For i = 1 To fMsg.NoOfDesignedMsgSections
                .MsgLabel(i) = dsply_msg_type.section(i).sLabel
                .MsgText(i) = dsply_msg_type.section(i).sText
                .MsgMonoSpaced(i) = dsply_msg_type.section(i).bMonspaced
            Next i
        End If
        
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


Using the Dsply function may look as follows:
```
Public Sub Test_Dsply()
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
    
    ' Preparing for the buttons
    cll.Add B1: cll.Add B2: cll.Add B3: cll.Add vbLf ' 3 buttons in row 1
    cll.Add B4: cll.Add B5: cll.Add B6: cll.Add vbLf ' 3 buttons in row 2
    cll.Add B7                                       ' 1 button in row 3
       
    ' Preparing for the message
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
       
   Select Case Dsply(dsply_title:="Any title", _
                     dsply_message:=tMsg, _
                     dsply_buttons:=cll)
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

displays:
![Test_Dsply](../Assets/TestDsply.png)
![Test_Dsply](/Assets/TestDsply.png)

### Proportional versus Mono-Spaced

The effect it has when a text in a section is specified mono-spaced (the default is proportional-spaced) is demonstrated by the second example in the [Services](#services) section above. Because the section specific message text is ++not++ "wrapped"  but The message windows width is ajusted up to the maximum width specified. In case even that's not enough a horizontal scroll-bar is displayed.