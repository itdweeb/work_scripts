Dim strComputer

' Check that all arguments required have been passed.
If Wscript.Arguments.Count < 1 Then
    Wscript.Echo "Arguments <Host> required. For example:" & vbCrLf _
    & "cscript vbping.vbs savdaldc01"
    Wscript.Quit(0)
End If

strComputer = Wscript.Arguments(0)

if Ping(strComputer) = True then
    Wscript.Echo "Host " & strComputer & " contacted"
Else
    Wscript.Echo "Host " & strComputer & " could not be contacted"
end if

Function Ping(strComputer)

    dim objPing, objRetStatus

    set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
      ("select * from Win32_PingStatus where address = '" & strComputer & "'")

    for each objRetStatus in objPing
        if IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode<>0 then
    Ping = False
            'WScript.Echo "Status code is " & objRetStatus.StatusCode
        else
            Ping = True
            'Wscript.Echo "Bytes = " & vbTab & objRetStatus.BufferSize
            'Wscript.Echo "Time (ms) = " & vbTab & objRetStatus.ResponseTime
            'Wscript.Echo "TTL (s) = " & vbTab & objRetStatus.ResponseTimeToLive
        end if
    next
End Function 
