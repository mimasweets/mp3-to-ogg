Option Explicit
Dim objArgs, droppedItem, droppedFile, fileType, shell, fso
Dim ffmpegPath, oggFileName
Set objArgs = WScript.Arguments
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If objArgs.Count > 0 Then
    For Each droppedItem In objArgs
        droppedFile = droppedItem
        fileType = GetMimeType(droppedFile)
        If InStr(fileType, "audio/mpeg") > 0 Then
		
            ' Specify the full path to the FFmpeg executable
            ffmpegPath = "C:\ffmpeg\bin\ffmpeg.exe"
            
            If FileExists(ffmpegPath) Then
                oggFileName = Left(droppedFile, Len(droppedFile) - 4) & ".ogg" 
                Dim command
                command = """" & ffmpegPath & """ -i """ & droppedFile & """ -c:a libvorbis """ & oggFileName & """"
                shell.Run command, 0, True
            Else
                MsgBox "FFmpeg executable not found.", vbExclamation
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
