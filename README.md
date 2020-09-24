<div align="center">

## Automatic About Form


</div>

### Description

This code sample allows you to display an About form for your application with a couple of lines of code.
 
### More Info
 
Uses API call ShellAbout


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Troy Blake](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/troy-blake.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/troy-blake-automatic-about-form__1-7062/archive/master.zip)

### API Declarations

```
'Insert in General Declarations
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
```


### Source Code

```
'Insert in Event (Like Button_Click)
Call ShellAbout(Me.hwnd, "- About Box Example", "A small example " & "that uses the ShellAbout Function to create an About Box.", Me.Icon)
```

