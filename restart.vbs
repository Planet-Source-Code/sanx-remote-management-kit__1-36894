OPTION EXPLICIT

    'Shutdown Method Constants
    CONST CONST_SHUTDOWN                = 1
    CONST CONST_LOGOFF                  = 0
    CONST CONST_POWEROFF                = 8
    CONST CONST_REBOOT                  = 2
    CONST CONST_FORCE_REBOOT            = 6
    CONST CONST_FORCE_POWEROFF          = 12
    CONST CONST_FORCE_LOGOFF            = 4
    CONST CONST_FORCE_SHUTDOWN          = 5
     
    'Declare variables
    Dim blnLogoff, blnPowerOff, blnReBoot, blnShutDown
    Dim blnForce
    Dim intTimer
    Dim UserArray(3)
    Dim MyCount

Sub Reboot(strServer)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, objEnumerator, objInstance
    Dim strQuery, strMessage
    Dim intStatus
    ReDim strID(0), strName(0)

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

    strID(0) = ""
    strName(0) = ""
    strMessage = ""
    strQuery = "Select * From Win32_OperatingSystem"

    Set objEnumerator = objService.ExecQuery(strQuery,,0)
    If Err.Number Then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred during the query."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Err.Clear
        Exit Sub
    End If

    i = 0
    For Each objInstance in objEnumerator
        If blnForce Then
            intStatus = objInstance.Win32ShutDown(CONST_FORCE_REBOOT)
        Else
            intStatus = objInstance.Win32ShutDown(CONST_REBOOT)
        End If

        IF intStatus = 0 Then
            strMessage = "Reboot of machine: " & strServer & " successful."
        Else
            strMessage = "Failed to reboot machine: " & strServer & "."
        End If
        Call WriteLine(strMessage,objOutputFile)
    Next

End Sub


'********************************************************************
'*
'* Sub LogOff()
'*
'* Purpose: Logs off the user currently logged onto a machine.
'*
'* Input:   strServer           a machine name
'*          strOutputFile       an output file name
'*          strUserName         the current user's name
'*          strPassword         the current user's password
'*          blnForce            specifies whether to force the logoff
'*          intTimer            specifies the amount of time to preform the function
'*
'* Output:  Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************
Private Sub LogOff(strServer)


    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, objEnumerator, objInstance
    Dim strQuery, strMessage
    Dim intStatus
    ReDim strID(0), strName(0)

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

    strID(0) = ""
    strName(0) = ""
    strMessage = ""
    strQuery = "Select * From Win32_OperatingSystem"

    Set objEnumerator = objService.ExecQuery(strQuery,,0)
    If Err.Number Then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred during the query."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Err.Clear
        Exit Sub
    End If

    i = 0
    For Each objInstance in objEnumerator
        If blnForce Then
            intStatus = objInstance.Win32ShutDown(CONST_FORCE_LOGOFF)
        Else
            intStatus = objInstance.Win32ShutDown(CONST_LOGOFF)
        End If

        IF intStatus = 0 Then
            strMessage = "Logging off the current user on machine " & _
                         strServer & "..."
        Else
            strMessage = "Failed to log off the current user from machine " _
                & strServer & "."
        End If
        Call WriteLine(strMessage,objOutputFile)
    Next

End Sub


'********************************************************************
'*
'* Sub PowerOff()
'*
'* Purpose: Powers off a machine.
'*
'* Input:   strServer           a machine name
'*          strOutputFile       an output file name
'*          strUserName         the current user's name
'*          strPassword         the current user's password
'*          blnForce            specifies whether to force the logoff
'*          intTimer            specifies the amount of time to perform the function
'*
'* Output:  Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************
Private Sub PowerOff(strServer)


    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, objEnumerator, objInstance
    Dim strQuery, strMessage
    Dim intStatus
    ReDim strID(0), strName(0)

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

    strID(0) = ""
    strName(0) = ""
    strMessage = ""
    strQuery = "Select * From Win32_OperatingSystem"

    Set objEnumerator = objService.ExecQuery(strQuery,,0)
    If Err.Number Then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred during the query."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Err.Clear
        Exit Sub
    End If

    i = 0
    For Each objInstance in objEnumerator
        If blnForce Then
            intStatus = objInstance.Win32ShutDown(CONST_FORCE_POWEROFF)
        Else
            intStatus = objInstance.Win32ShutDown(CONST_POWEROFF)
        End If

        IF intStatus = 0 Then
            strMessage = "Power off of machine: " & strServer & " successful."
        Else
            strMessage = "Failed to power off machine: " & strServer & "."
        End If
        Call WriteLine(strMessage,objOutputFile)
    Next

End Sub


'********************************************************************
'*
'* Sub Shutdown()
'*
'* Purpose: Shutsdown a machine.
'*
'* Input:   strServer           a machine name
'*          strOutputFile       an output file name
'*          strUserName         the current user's name
'*          strPassword         the current user's password
'*          blnForce            specifies whether to force the logoff
'*          intTimer            specifies the amount of time to perform the function
'*
'* Output:  Results are either printed on screen or saved in strOutputFile.
'*
'********************************************************************
Private Sub Shutdown(strServer)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, objEnumerator, objInstance
    Dim strQuery, strMessage
    Dim intStatus
    ReDim strID(0), strName(0)

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

    strID(0) = ""
    strName(0) = ""
    strMessage = ""
    strQuery = "Select * From Win32_OperatingSystem"

    Set objEnumerator = objService.ExecQuery(strQuery,,0)
    If Err.Number Then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred during the query."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Err.Clear
        Exit Sub
    End If

    i = 0
    For Each objInstance in objEnumerator
        If blnForce Then
            intStatus = objInstance.Win32ShutDown(CONST_FORCE_SHUTDOWN)
        Else
            intStatus = objInstance.Win32ShutDown(CONST_SHUTDOWN)
        End If

        IF intStatus = 0 Then
            strMessage = "Shut down of machine: " & strServer & "successful."
        Else
            strMessage = "Failed to shut down machine: " & strServer & "."
        End If
        Call WriteLine(strMessage,objOutputFile)
    Next

End Sub
