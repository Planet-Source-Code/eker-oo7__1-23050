<div align="center">

## OO7


</div>

### Description

This submission subclasses a folder via a com object I created.

The COM object uses a callback to a function defined in a module , through an undocumented API which hooks into the event notification

system of the shell. In this a case I created a small form to allow you to see it in action. Click Command1 To register the window for notification of events triggered in the C:\temp folder or any folder you want(just change the code), Then copy a file into the folder,or rename a file in the folder, or delete one- whatever. You will get a msgbox alerting your application window that an event has occurred. It also logs to the NT log or a textfile, and detects the current operating system (all via the windows API). I use this in place of timer based routines that just wait poll folder for the existence files. Like the SMTP Service.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-05-08 23:57:50
**By**             |[Eker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eker.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[OO719418592001\.zip](https://github.com/Planet-Source-Code/eker-oo7__1-23050/archive/master.zip)

### API Declarations

See File..Too many to list.





