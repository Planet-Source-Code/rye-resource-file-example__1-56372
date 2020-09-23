Attribute VB_Name = "MGeneral"
Option Explicit

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Dim OCX() As Byte

Sub Main()

    'Here you would check, extract and register
    'any components your projects uses
    
    FMain.Show 'now thats the components are installed your projects will run without fault

End Sub

Public Function FileExists(FileName As String) As Boolean

    FileExists = (Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) <> "")

End Function

Public Function Register(FullPath As String, Optional Slient As Boolean)

    If Slient = True Then
        Shell ("Regsvr32 " & FullPath & " /s") 'Slient register no feedback is given to the user you just hope it worked ;)
    Else
        Shell ("Regsvr32 " & FullPath) 'Unslient register this will confront the user with a dialog when Regsvr32 has finished registering the file
    End If

End Function

Public Function Unregister(FullPath As String, Optional Slient As Boolean)

    If Slient = True Then
        Shell ("Regsvr32 " & FullPath & " /u /s") 'Slient unregister no feedback is given to the user you just hope it worked ;)
    Else
        Shell ("Regsvr32 " & FullPath & " /u") 'Unslient unregister this will confront the user with a dialog when Regsvr32 has finished registering the file
    End If

End Function

Public Function ExtractResource(ResType As String, ResID As Long, FullOutputPath As String)

On Error GoTo Error

    If Not FileExists(FullOutputPath) Then 'Check if the file exists
        OCX = LoadResData(ResID, ResType) 'Loads the .OCX from the resource file
        Open FullOutputPath For Binary As 1 'Opens the output file so we can insert out .OCX
            Put #1, , OCX 'Inserts the .OCX
        Close #1 'Closes the file
    Else
        MsgBox "File already exists, no need to extract it", vbInformation, "File Already Exists"
    End If
    
    Exit Function
    
Error:

    MsgBox "Encountered and error somewhere make sure you have wrtie access to the output path or the file is not already in use", vbCritical, " Extraction Error"
    
End Function

Public Function SystemDirectory() As String
    
    Dim RSTR As String
    Dim RLEN As Long

    RSTR = String(255, 0)
    RLEN = GetSystemDirectory(RSTR, Len(RSTR))
    If RLEN < Len(RSTR) Then
        RSTR = Left(RSTR, RLEN)
        If Right(RSTR, 1) = "\" Then
            SystemDirectory = Left(RSTR, Len(RSTR) - 1)
        Else
            SystemDirectory = RSTR
        End If
    Else
        SystemDirectory = ""
    End If
    
End Function
