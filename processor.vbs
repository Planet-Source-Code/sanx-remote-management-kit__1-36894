OPTION EXPLICIT

Private Sub GetProcInfo(strServer)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, objProcSet, objProc
    Dim strWBEMClass

    Call HTMLHeaders

    strWBEMClass = "Win32_Processor"

    'Establish a connection with the server.
    If blnConnect("root\cimv2" , _
                   strUserName , _
                   strPassword , _
                   strServer   , _
                   objService  ) Then
        Call document.write("")
        Call document.write("Please check the server name, " _
                        & "credentials and WBEM Core.")
        Exit Sub
    End If

    'Get the set
    Set objProcSet = objService.InstancesOf(strWBEMClass)
        If blnErrorOccurred("Could not obtain " & _
                   strWBEMClass & " instance.") Then
        Exit Sub
    End If

    If objProcSet.Count = 0 Then
        Call WriteLine("No processor information is available.")    
        Exit Sub
    End If

    Call WriteLine("<DIV STYLE='font-family: tahoma, arial, sans-serif; font-size: 10pt'>" & _
				"<TABLE><TR><TH COLSPAN=2 BGCOLOR='#D0D0FF'>Processor information for server: " & _
                strServer & "</TH></TR>")

    For Each objProc In objProcSet
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Name</TD><TD BGCOLOR='#D0FFFF'>" & _
             objProc.Name & "</TD></TR>")
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Current Voltage</TD><TD BGCOLOR='#D0FFFF'>" & _
             (objProc.CurrentVoltage)/10 & "</TD></TR>")
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Device ID</TD><TD BGCOLOR='#D0FFFF'>" & _
             objProc.DeviceID & "</TD></TR>")
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Status</TD><TD BGCOLOR='#D0FFFF'>" & _
             objProc.CpuStatus & "</TD></TR>")
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Data Width</TD><TD BGCOLOR='#D0FFFF'>" & _
             objProc.DataWidth & "</TD></TR>")
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Current Clock Speed</TD><TD BGCOLOR='#D0FFFF'>" & _
             objProc.CurrentClockSpeed & "</TD></TR>")
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>L2 Cache Size</TD><TD BGCOLOR='#D0FFFF'>" & _
             objProc.L2CacheSize & "</TD></TR>")
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Level</TD><TD BGCOLOR='#D0FFFF'>" & _
             objProc.Level & "</TD></TR>")
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>External Clock</TD><TD BGCOLOR='#D0FFFF'>" & _
             objProc.ExtClock & "</TD></TR>")
    Next
	
	WriteLine "</TABLE></DIV>"

End Sub
