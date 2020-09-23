OPTION EXPLICIT
	
	CONST CONST_EXECUTE                 = 4
    CONST CONST_KILL                    = 5

    'Declare variables
    Dim intProcessID
    Dim strCommand

Sub ExecuteCmd(strServer, strCommand)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, objInstance 
    Dim strQuery, strMessage
    Dim intProcessId, intStatus

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

    strMessage = ""
    intProcessId = 0

    Set objInstance = objService.Get("Win32_Process")
    If blnErrorOccurred(" occurred getting a " & _
                        " Win32_Process class object.") Then Exit Sub

    If objInstance is nothing Then Exit Sub

    intStatus = objInstance.Create(strCommand, null, null, intProcessId)
    If blnErrorOccurred(" occurred in creating process " & _
                          strCommand & ".") Then Exit Sub

    If intStatus = 0 Then
        If intProcessId < 0 Then
            '4294967296 is 0x100000000.
            intProcessId = intProcessId + 4294967296
        End If
        strMessage = "Succeeded in executing " &  strCommand & "." & vbCRLF
        strMessage = strMessage & "The process id is " & intProcessId & "."
    Else
        strMessage = "Failed to execute " & strCommand & "." & vbCRLF
        strMessage = strMessage & "Status = " & intStatus
    End If
    WriteLine strMessage

End Sub

Sub Kill(strServer, intProcessID)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, objInstance
    Dim strWBEMClass, strMessage
    Dim intStatus

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

    'Now executes the method.
    If strServer = "" Then
        strWBEMClass = "Win32_Process.Handle=" & intProcessId
    Else
        strWBEMClass = "\\" & strServer & "\root\cimv2:Win32_Process.Handle=" _
                       & intProcessId
    End If

    Set objInstance = objService.Get(strWBEMClass)
    If blnErrorOccurred(" occurred in getting process " & strWBEMClass & ".") _
                          Then Exit Sub

    intStatus = objInstance.Terminate

    If intStatus = 0 Then
        strMessage = "Process " & intProcessId & " has been killed."
    Else
        strMessage = "Failed to kill process " & intProcessId & "."
    End If

    WriteLine strMessage

End Sub
