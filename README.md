<div align="center">

## Computer Information


</div>

### Description

This simply uses the built-in Environ() function to return useful information about your computer. I've found it useful in several applications that I've made. I'm sure almost everyone has learned this, but if you haven't, it's definately something you should.
 
### More Info
 
These are the results I got in this exact order.

ALLUSERSPROFILE

APPDATA

CLIENTNAME

CommonProgramFiles

COMPUTERNAME

ComSpec

HOMEDRIVE

HOMEPATH

include

LOGONSERVER

MSDevDir

NUMBER_OF_PROCESSORS

OS

Path

PATHEXT

PROCESSOR_ARCHITECTURE

PROCESSOR_IDENTIFIER

PROCESSOR_LEVEL

PROCESSOR_REVISION

ProgramFiles

SESSIONNAME

SystemDrive

SystemRoot

TEMP

TMP

USERDOMAIN

USERNAME

USERPROFILE

windir


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel W\. Elkins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-w-elkins.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-w-elkins-computer-information__1-44616/archive/master.zip)





### Source Code

```
Option Explicit
Private Sub Form_Load()
Dim A As Integer
For A = 1 to 30
Debug.Print Environ$(A)
DoEvents
Next A
End Sub
```

