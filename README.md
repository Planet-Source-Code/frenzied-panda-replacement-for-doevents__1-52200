<div align="center">

## Replacement for DoEvents


</div>

### Description

Stop All these DoEvents cr*p post. This is the way it windows does it and done in c++.

Slight update from my last one. zip is included.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2004-03-07 11:53:40
**By**             |[Frenzied\-Panda](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/frenzied-panda.md)
**Level**          |Beginner
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Replacemen171717372004\.zip](https://github.com/Planet-Source-Code/frenzied-panda-replacement-for-doevents__1-52200/archive/master.zip)





### Source Code

```
While b
 'It is better if you get the hwnd value beforehand so you
 'do not have to ask the vb dll for it each and every time
 DoMyEvents (Me.hWnd) 'Endless loop
 'If you want other window to be processed
 'then you need to DoMyEvents () with their HWND
 'That is why you cannot pause the ide window
 'when your in this loop.
Wend
```

