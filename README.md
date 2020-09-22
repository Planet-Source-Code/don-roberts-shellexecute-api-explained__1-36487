<div align="center">

## ShellExecute API explained


</div>

### Description

I view many VB coding websites and mailing lists, and a common question is asked: "How do I open (this document or program) from my application?" The answer is the ShellExecute API (I never use Shell())
 
### More Info
 
This example assumes you have the file 'frunlog.txt' in your c:\ drive.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Don Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/don-roberts.md)
**Level**          |Beginner
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/don-roberts-shellexecute-api-explained__1-36487/archive/master.zip)

### API Declarations

```
Paste this into a module:
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function OpenFile(hWnd As Long, strOperation As String, ByVal File As String)
 Dim lRet As Long
 'these constants are all the values you can have for
 'the 'nShowCmd' part of the function
 Const SW_SHOWNORMAL = 1
 Const SW_HIDE As Long = 0
 Const SW_MAXIMIZE As Long = 3
 Const SW_MINIMIZE As Long = 6
 Const SW_RESTORE As Long = 9
 Const SW_SHOW As Long = 5
 Const SW_SHOWDEFAULT As Long = 10
 Const SW_SHOWMAXIMIZED As Long = 3
 Const SW_SHOWMINIMIZED As Long = 2
 Const SW_SHOWMINNOACTIVE As Long = 7
 Const SW_SHOWNA As Long = 8
 Const SW_SHOWNOACTIVATE As Long = 4
 'the 'lpOperation' can have 3 different values:
 '"Open"
 '"Print"
 'and "Explore" all in quotes
 'if you use vbNullString, the default is "Open"
 'the 'lpFile' is of course the name of the file, either
 'and executable or a file with an association
 'the 'lpParameters' part of the function can hold any command line
 'switches that the called program may have. You can't use this when
 'opening regular files. You usually don't need this.
 'lpDirectory is the default directory of the application, here I just used
 'the application's directory.
 lRet = ShellExecute(hWnd, strOperation, File, vbNullString, App.Path, SW_SHOWNORMAL)
End Function
```


### Source Code

```
Open a new project with three command buttons:
Option Explicit
Private Sub Command1_Click()
 OpenFile Me.hWnd, "Open", "C:\frunlog.txt"
End Sub
Private Sub Command2_Click()
 OpenFile Me.hWnd, "Print", "C:\frunlog.txt"
End Sub
Private Sub Command3_Click()
 OpenFile Me.hWnd, "Explore", "C:\"
End Sub
```

