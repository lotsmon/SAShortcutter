Dim Login, Password
Dim oShell : Set oShell = CreateObject("WScript.Shell")

Dim Path : Path = "C:\Program Files\Steam\steam.exe"

Function Main()
    if WScript.Arguments.Count = 2 then
        Login = WScript.Arguments(0)
        Password = WScript.Arguments(1)
    end if

    if(Login <> "") Then
        if(Password <> "") Then
            oShell.Run "taskkill /im Steam.exe /F", 0, True
            quoted_path = CHR(34) & Path & CHR(34)
            oShell.Run quoted_path & "-login " & Login & " " & Password, 0, True
        End If
    End If
    Set oShell = Nothing
End Function 

Main
