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
