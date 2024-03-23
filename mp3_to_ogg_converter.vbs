Option Explicit

Dim objArgs, droppedItem, droppedFile, fileType, shell, fso
Dim scriptFolder, ffmpegPath, oggFileName

Set objArgs = WScript.Arguments
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
If objArgs.Count > 0 Then
    For Each droppedItem In objArgs
        droppedFile = droppedItem
        fileType = GetMimeType(droppedFile)        
        If InStr(fileType, "audio/mpeg") > 0 Then
            ffmpegPath = scriptFolder & "\ffmpeg.exe"            
            If FileExists(ffmpegPath) Then
                oggFileName = Left(droppedFile, Len(droppedFile) - 4) & ".ogg" 
                Dim command
                command = """" & ffmpegPath & """ -i """ & droppedFile & """ -c:a libvorbis """ & oggFileName & """"
                shell.Run command, 0, True
            Else
                MsgBox "FFmpeg executable not found in the script's folder.", vbExclamation
            End If
        End If
    Next
End If
Function GetMimeType(filePath)
    Dim fileExt    
    fileExt = fso.GetExtensionName(filePath)    
    If LCase(fileExt) = "mp3" Then
        GetMimeType = "audio/mpeg"
    Else
        GetMimeType = ""
    End If
End Function

Function FileExists(filePath)
    If fso.FileExists(filePath) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
