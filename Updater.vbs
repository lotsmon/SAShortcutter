Dim WshShell1, WshEnvirUsr, FSO, MSG, MSGTitle
MSGTitle = "SASUpdater"

set WshShell1 = WScript.CreateObject("WScript.Shell")
set WshEnvirUsr = WshShell1.Environment("USER")
set FSO = CreateObject("Scripting.FileSystemObject")
Dir = FSO.GetAbsolutePathName(".") & "\Starter.vbs"

Function Main()
    if FSO.FileExists(Dir) Then
        if WshEnvirUsr("SASLNK") = Dir Then
            MSG = MsgBox("     Already updated     ", vbOKOnly, MSGTitle)
        Else
            WshEnvirUsr("SASLNK") = Dir
        MSG = MsgBox("     Successfully updated     ", vbOKOnly, MSGTitle)
        End If
    Else
        MSG = MsgBox("     The 'Starter.vbs' file was not found     ", vbRetryCancel, MSGTitle)
        if MSG = vbRetry Then
            Main
        End If
    End If

End Function

Main
