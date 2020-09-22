<div align="center">

## Find File


</div>

### Description

Will locate a file on any type of drive. I use it for lots of things with little modification. Very useful for looping through all your drives, folders, sub-folders, etc. Perfect for finding files, folders, types of drives, etc. Should be "readable" enough for newbies and ideal for experts as well. Uses File System Object (FSO). Works with VB 5 as long as you've installed VB Scripting support. Can be implemented in ASP's with very little effort.
 
### More Info
 
'name of the file you're looking for

'Form code...

Private Sub cmdFindFile_Click()

Dim strFileName As String

Dim strTmp As String

strFileName = InputBox("Enter file name to look for", "Find a file")

If Len(strFileName) = 0 Then 'Hit cancel or didn't enter anything

Else

With cmdFindFile

strTmp = .Caption

.Caption = "Searching..."

.Enabled = False

FindIt (strFileName)

.Caption = strTmp

.Enabled = True

End With

End If

End Sub

'Paste code into a module or form

'If packaging, you'll need to ship scrrun.dll

'MsgBox containing path name

'If you have mapped network drives and don't have permissions to the

'root folders of those drives, you'll get an error. Easily fixable by not looking at those drives or placing more dynamic error handling code in there.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Code Man](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/code-man.md)
**Level**          |Unknown
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/code-man-find-file__1-4068/archive/master.zip)

### API Declarations

```
'None
'Set a Reference to MS Scripting Runtime
```


### Source Code

```
'Source Code for mdlFindFile.bas or put directly into form
Dim strLocation As String
Dim blFoundItFlag As Boolean
'Different Drive Types
'0 = "Unknown"
'1 = "Removable"
'2 = "Fixed"
'3 = "Network"
'4 = "CD-ROM"
'5 = "RAM Disk"
Public Sub FindIt(strFileName As String)
Dim FS As FileSystemObject
Dim Drv As Drive
Dim DrvCol
Dim RootFldr As Folder
Dim strRootPath As String
Dim strFNameToPass As String
blFoundItFlag = False
strFNameToPass = UCase(strFileName) 'will speed processing passing it this way & ensure proper comparison
 Set FS = CreateObject("Scripting.FileSystemObject")
 Set DrvCol = FS.Drives
 For Each Drv In DrvCol
 If blFoundItFlag Then 'Once we found it, don't got through the rest of the drives
 Exit Sub
 Else
 strRootPath = Drv.DriveLetter & ":\"
 If Drv.IsReady Then 'Will prevent errors
 Set RootFldr = FS.GetFolder(strRootPath)
 Call CheckEm(RootFldr, strRootPath, strFNameToPass)
 End If
 End If
 Next
End Sub
Public Sub CheckEm(Fldr As Folder, Path As String, FName As String)
 Dim SubFldr As Folder
 Dim strPath As String
 Dim strFName As String
On Error GoTo ErrHandler
 strPath = Path
 strFName = FName
 For Each SubFldr In Fldr.SubFolders
 For Each Fil In SubFldr.Files
 strLocation = SubFldr.ParentFolder & "\" & SubFldr.Name & "\"
 DoEvents
 'Debug.Print strLocation
 If UCase(Fil.Name) = strFName Then
 strLocation = Replace(strLocation, "\\", "\") 'Some paths have 2 \\ ???
 MsgBox strLocation 'show em where it's at
 blFoundItFlag = True
 Exit Sub
 End If
 Next
 Call CheckEm(SubFldr, strPath, strFName) 'Little recursive action here
 Next
Exit Sub
ErrHandler:
 If MsgBox("Error: " & Err.Number & " " & Err.Description & vbCrLf & _
 "Do you want to continue?", vbYesNo) = vbYes Then
 Resume Next
 Else
 blFoundItFlag = True
 Exit Sub
 End If
End Sub
```

