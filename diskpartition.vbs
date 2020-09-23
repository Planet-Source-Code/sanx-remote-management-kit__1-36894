OPTION EXPLICIT

Private Sub GetDskPartInf(strServer)


    ON ERROR RESUME NEXT

    Dim objFileSystem, objOutputFile, objService, objDSKPSet, objDSKP
    Dim strQuery, strMessage
    Dim strWBEMClass

    Call HTMLHeaders

    strWBEMClass = "Win32_DiskPartition"

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
    Set objDSKPSet = objService.InstancesOf(strWBEMClass)
        If blnErrorOccurred("Could not obtain " & strWBEMClass & _
            " instance.") Then
        Exit Sub
    End If

    If objDSKPSet.Count = 0 Then
        Call WriteLine("No disk partition information is available.")    
        Exit Sub
    End If

    Call WriteLine("<DIV STYLE='font-family: tahoma, arial, sans-serif; font-size: 10pt'>" & _
				"<TABLE><TR><TH COLSPAN=2 BGCOLOR='#D0D0FF'>Partition Information for server: " & _ 
				UCase(strServer) & "</TH></TR><TR><TD>&nbsp;</TD></TR>" & vbCRLF)

    For Each objDSKP In objDSKPSet
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Name</TD><TD BGCOLOR='#D0FFFF'>" & objDSKP.Name & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Device ID</TD><TD BGCOLOR='#D0FFFF'>" & objDSKP.DeviceID & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Size</TD><TD BGCOLOR='#D0FFFF'>" & strInsertCommas(objDSKP.Size) & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Boot Partition</TD><TD BGCOLOR='#D0FFFF'>" & objDSKP.BootPartition & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Disk Index</TD><TD BGCOLOR='#D0FFFF'>" & objDSKP.DiskIndex & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Primary Partition</TD><TD BGCOLOR='#D0FFFF'>"  & objDSKP.PrimaryPartition & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Starting Offset</TD><TD BGCOLOR='#D0FFFF'>" & objDSKP.StartingOffset & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Type</TD><TD BGCOLOR='#D0FFFF'>" & objDSKP.Type & "</TD></TR><TR><TD>&nbsp;</TD></TR>" & vbCRLF)
    Next
	Call WriteLine("</TABLE></DIV>" & vbCRLF)

End Sub
