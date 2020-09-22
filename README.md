<div align="center">

## Random Wav Player


</div>

### Description

I was really bored today, so I made a completely useless little program that selects a random .wav from a specified folder, plays it and then closes. It could be put in the startup folder and used to play a random .wav sound when Windows starts (as a replacement for the Windows sounds), or any other event.
 
### More Info
 
If you use this at Windows Startup, disable the "Start Windows" sound in the Control Panel > Sounds utility. Put this code on a blank form named "Form1" and change the folder that it searches.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andy T\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andy-t.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andy-t-random-wav-player__1-21074/archive/master.zip)





### Source Code

```
Private Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" _
  (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
' If you use this at Windows Startup, disable the "Start Windows" sound in the Control Panel > Sounds utility.
Sub PlaySound()
  Dim fsoFileSystem, fsoFolder, fsoFile, fsoFolderFiles
  Dim strWavs(0 To 50) As String
  Dim intCounter As Integer
  Dim strFileName As String
  intCounter = 0
  Set fsoFileSystem = CreateObject("Scripting.FileSystemObject")
  Set fsoFolder = fsoFileSystem.GetFolder("c:\winnt\media") '<< OR WHATEVER FOLDER YOU WANT
  Set fsoFolderFiles = fsoFolder.Files
  For Each fsoFile In fsoFolderFiles
    If Right(fsoFile.Name, 4) = ".wav" Then
      strWavs(intCounter) = fsoFile.Name
      intCounter = intCounter + 1
    End If
  Next
  strFileName = strWavs(Int(Rnd * intCounter))
  Call sndPlaySound32(fsoFolder & "\" & strFileName, 0)
End Sub
Private Sub Form_Load()
  Form1.Visible = False
  PlaySound
  End
  '(pretty simple, huh?)
End Sub
```

