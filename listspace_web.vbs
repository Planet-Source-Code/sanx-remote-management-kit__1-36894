OPTION EXPLICIT

    'Define constants
    CONST CONST_ERROR                   = 0
    CONST CONST_WSCRIPT                 = 1
    CONST CONST_CSCRIPT                 = 2
    CONST CONST_SHOW_USAGE              = 3
    CONST CONST_PROCEED                 = 4

    'Declare variables
    Dim intOpMode, i
    Dim strServer, strUserName, strPassword, strOutputFile

Sub ListFreeSpace(strServer)

    ON ERROR RESUME NEXT
    Dim objFileSystem, objService, objEnumerator, objInstance
    Dim strQuery, strMessage
    Dim lngSpace

    Call HTMLHeaders

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

    strQuery = "Select FreeSpace, DeviceID From Win32_LogicalDisk"

    Set objEnumerator = objService.ExecQuery(strQuery,,0)
    If Err.Number Then
        document.write("Error 0x" & CStr(Hex(Err.Number)) & _
                     " occurred during the query.")
        If Err.Description <> "" Then
            document.write("Error description: " & Err.Description & ".")
        End If
        Err.Clear
        Exit Sub
    End If

	WriteLine "<DIV STYLE='font-family: tahoma, arial, sans-serif; font-size: 10pt'>" & _
				"<TABLE><TR><TH COLSPAN=2 BGCOLOR='#D0D0FF'>Disk Information for server: " & _ 
				UCase(strServer) & "</TH></TR>"
	
    For Each objInstance in objEnumerator
        If Not (objInstance is nothing) Then
            strMessage = "<TR><TD WIDTH='75px' BGCOLOR='#D0FFD0'>" & objInstance.DeviceID & "</TD><TD BGCOLOR='#D0FFFF'>"
            lngSpace = objInstance.FreeSpace
            If lngSpace <> 0 Then
                strMessage = strMessage & _
                  strInsertCommas(lngSpace)& " bytes free</TD></TR>"
            Else
                strMessage = strMessage & _
                  "not available</TD></TR>"
            End If
            WriteLine strMessage
        End If
        If Err.Number Then
            Err.Clear
        End If
    Next
	
	WriteLine "</TABLE></DIV>"

End Sub