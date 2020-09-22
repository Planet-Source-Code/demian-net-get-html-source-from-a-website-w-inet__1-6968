<div align="center">

## Get HTML Source From a Website w\\ Inet


</div>

### Description

Just like it says: Get HTML Source From a Website!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Demian Net](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/demian-net.md)
**Level**          |Advanced
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/demian-net-get-html-source-from-a-website-w-inet__1-6968/archive/master.zip)





### Source Code

```
'Load Richtx32.ocx
 'Load msinet.ocx
 'Make a RichTextBox1
 'Make an Inet1
 'Make a plain textbox names URL
 'Make a command1
 Private Sub Command1_Click()
 On Error Resume Next
   Dim txt As String
   Dim b() As Byte
   Command1.Enabled = False
   b() = Inet1.OpenURL(URL.Text, 1)
   txt = ""
   For t = 0 To UBound(b) - 1
     txt = txt + Chr(b(t))
   Next
   RichTextBox1.Text = txt
   Command1.Enabled = True
 Exit Sub
 End Sub
```

