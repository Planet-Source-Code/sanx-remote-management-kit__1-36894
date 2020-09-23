OPTION EXPLICIT

Private Sub LogMemInfo(strServer)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, objMem, objWshNet
    Dim strQuery, strMessage, strCat

    Call HTMLHeaders

    'Establish a connection with the server.
    If blnConnect("root\cimv2" , _
                   strUserName , _
                   strPassword , _
                   strServer   , _
                   objService  ) Then
        Call Wscript.Echo("")
        Call Wscript.Echo("Please check the server name, " _
                        & "credentials and WBEM Core.")
        Exit Sub
    End If

    'Get the logical memory configuration
    Set objMem = objService.Get("Win32_LogicalMemoryConfiguration=""" _
                 & "LogicalMemoryConfiguration""")
    If Err.Number Then
      Wscript.Echo "Error 0x" & CStr(Hex(Err.Number)) & _
                   " occurred getting the memory configuration."
      If Err.Description <> "" Then
          Wscript.Echo "Error description: " & Err.Description & "."
      End If
      Err.Clear
      Exit Sub
    End If

    Call WriteLine("<DIV STYLE='font-family: tahoma, arial, sans-serif; font-size: 10pt'>" & _
				"<TABLE><TR><TH COLSPAN=2 BGCOLOR='#D0D0FF'>Logical Memory " & _
				"Configuration of server: " & UCase(strServer) & _
				"</TH></TR><TR><TD>&nbsp;</TD></TR>" & vbCRLF)

    Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Total Physical</TD><TD BGCOLOR='#D0FFFF'>" & strInsertCommas(objMem.TotalPhysicalMemory) & "</TD></TR>" & vbCRLF)
	Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Total Virtual</TD><TD BGCOLOR='#D0FFFF'>" & strInsertCommas(objMem.TotalVirtualMemory) & "</TD></TR>" & vbCRLF)
	Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Pagefile Space</TD><TD BGCOLOR='#D0FFFF'>" & strInsertCommas(objMem.TotalPagefileSpace) & "</TD></TR>" & vbCRLF)
	Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Available Virtual</TD><TD BGCOLOR='#D0FFFF'>" & strInsertCommas(objMem.AvailableVirtualMemory) & "</TD></TR>" & vbCRLF)
	Call WriteLine("</TABLE></DIV>" & vbCRLF)
	
End Sub
