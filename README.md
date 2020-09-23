<div align="center">

## Resource File Example


</div>

### Description

This Resource File Example extracts files from resource files at/before runtime to ensure your program will run regardless of weather the user has the needed files (inculding VB RTF).

Now i'm sure i'm not the only person to notice that quite alot of the comments found at PSC look along the lines of this...

"Hey wheres the Blah.dll"

"There was no Blah.ocx in the zip"

"You didn't inculde the .ocx/.dll in the zip"

Now there is one main reason why they don't inculde them and that is because of PSC rule of not inculding .ocx or .dll files, but that doesn't mean to say you carn't inculde these files in your project and upload them to PSC.

In the archive is the "Microsoft Winsock Control 6.0" but it's inside a .res file made by Visual Basics which is inside the project, my project will exetract files from these .res files and return them to there orignal form and even register it if need be this can be performed before a form is loaded so that it is insured your project will run without fault.

This is very simple code and doesn't need much explaining i also understand this can lead to a new way to insert virii into projects but if you are that dumb to fall for it then sorry to say you deserved it :P

As you will see (by reading the source code) my project does no call and functions from the .ocx inculded in this project it only extracts it into a directory and registers it using Regsvr32 so there is 0% risk of virii infection
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2004-09-27 05:59:02
**By**             |[Rye](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rye.md)
**Level**          |Beginner
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Res\_\(Resou1797959272004\.zip](https://github.com/Planet-Source-Code/rye-resource-file-example__1-56372/archive/master.zip)








