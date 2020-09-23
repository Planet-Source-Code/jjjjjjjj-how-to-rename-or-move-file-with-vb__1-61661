<div align="center">

## How to RENAME or MOVE file with Vb???


</div>

### Description

After some quick search I realized that very few in PSC knows how to rename file with VB. So here it is...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[JJJJJJJJ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jjjjjjjj.md)
**Level**          |Beginner
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jjjjjjjj-how-to-rename-or-move-file-with-vb__1-61661/archive/master.zip)





### Source Code

```
'------------------------------------------------
' Auther : Jim Jose
' Email  : jimjosev33@yahoo.com
' Purpose : File renaming with vb
'------------------------------------------------
' Very Easy..
' >>>Name [Source] As [Destination]<<<
'------------------------------------------------
Private Sub cmdTest_Click()
 'SampleCode:
 'Rename file in D:\About.txt to D:\Me.txt
 'Use...
 Name "D:\About.txt" As "D:\Me.txt"
 'SampleCode:
 'Movefile file from D:\About.txt to C:\About.txt
 'Use...
 Name "D:\About.txt" As "C:\About.txt"
 'Thats it... So never use FileCopy..
 'with Kill.. to rename files.
 'Jim Jose :-))
End Sub
```

