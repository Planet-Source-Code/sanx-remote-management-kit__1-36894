OPTION EXPLICIT

    CONST CONST_DELETE                  = "DELETE"
    CONST CONST_CREATE                  = "CREATE"
    
    'Declare variables
    Dim strShareName, strSharePath
    Dim strShareComment, strShareType


Private Sub Share(strServer,strShareName,strSharePath,strShareType,strShareComment,strTaskCommand)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, strQuery

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

	'Now execute the method.
	Call ExecuteShare(objService, strTaskCommand, _
         strShareName,strSharePath,strShareType,strShareComment,strServer)

End Sub


Private Sub ExecuteShare(objService, strTaskCommand, _
            strShareName, strSharePath,strShareType,strShareComment,strServer)

    ON ERROR RESUME NEXT
	
    Dim intType, intShareType, i, intStatus, strMessage
    Dim objEnumerator, objInstance
    ReDim strName(0), strDescription(0), strPath(0),strType(0), intOrder(0)

    intShareType = 0
    strMessage = ""
    strName(0) = ""
    strPath(0) = ""
    strDescription(0) = ""
    strType(0) = ""
    intOrder(0) = 0

    Select Case strTaskCommand
        Case CONST_CREATE
            Set objInstance = objService.Get("Win32_Share")
            If Err.Number Then
                Print "Error 0x" & CStr(Hex(Err.Number)) & _
                  " occurred in getting " & " a share object."
                If Err.Description <> "" Then
                     Print "Error description: " & Err.Description & "."
                End If
                Err.Clear
                Exit Sub
            End If

            If objInstance is nothing Then
                Exit Sub
            Else
                Select Case strShareType
                    Case "Disk"
                        intShareType = 0
                    Case "PrinterQ"
                        intShareType = 1
                    Case "Device"
                        intShareType = 2
                    Case "IPC"
                        intShareType = 3
                    Case "Disk$"
                        intShareType = -2147483648
                    Case "PrinterQ$"
                        intShareType = -2147483647
                    Case "Device$"
                        intShareType = -2147483646
                    Case "IPC$"
                        intShareType = -2147483645
                End Select

                intStatus = objInstance.Create(strSharePath, strShareName, _
                    intShareType, null, strShareComment, null, null)
                If intStatus = 0 Then
                    strMessage = "Succeeded in creating share " & _
                      strShareName & "."
                Else
                    strMessage = "Failed to create share " & strShareName & "."
                    strMessage = strMessage & vbCRLF & "Status = " & _
                      intStatus & "."
                End If

                WriteLine strMessage
                i = i + 1
            End If
        Case CONST_DELETE
            Set objInstance = objService.Get("Win32_Share='" & strShareName _
              & "'")
            If Err.Number Then
                Print "Error 0x" & CStr(Hex(Err.Number)) & _
                  " occurred in getting share " _
                    & strShareName & "."
                If Err.Description <> "" Then
                    Print "Error description: " & Err.Description & "."
                End If
                Err.Clear
                Exit Sub
            End If

            If objInstance is nothing Then
                Exit Sub
            Else
                intStatus = objInstance.Delete()
                If intStatus = 0 Then
                    strMessage = "Succeeded in deleting share " & _
                    strShareName & "."
                Else
                    strMessage = "Failed to delete share " & strShareName & "."
                    strMessage = strMessage & vbCRLF & "Status = " & _
                    intStatus & "."
                End If
                WriteLine strMessage
                i = i + 1
            End If
        Case CONST_LIST
            Set objEnumerator = objService.ExecQuery (_
                "Select Name,Description,Path,Type From Win32_Share",,0)
            If Err.Number Then
                Print "Error 0x" & CStr(Hex(Err.Number)) & _
                  " occurred during the query."
                If Err.Description <> "" Then
                    Print "Error description: " & Err.Description & "."
                End If
                Err.Clear
                Exit Sub
            End If
			WriteLine "<DIV STYLE='font-family: tahoma; font-size: 10pt'>" & _
					"<TABLE CELLPADDING='2'><TR><TH COLSPAN=4 BGCOLOR='#D0D0FF'>" & _
					"Share information for server: " & UCase(strServer) & "</TH></TR>"
            Call WriteLine("<TR><TD COLSPAN=4 BGCOLOR='#D0D0FF'>There are " & objEnumerator.Count & _
              " shares.</TD></TR>")
            Call WriteLine("<TR><TD>Name</TD><TD>Type</TD><TD>Description</TD><TD>Path</TD></TR>")
            For Each objInstance in objEnumerator
                I = I + 1
                If objInstance is nothing Then
                    Exit Sub
                End If
                Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>" & objInstance.Name & "</TD>")
                Select Case objInstance.Type
                    Case 0
                        Call WriteLine("<TD BGCOLOR='#D0FFFF'>Disk</TD>")
                    Case 1
                        Call WriteLine("<TD BGCOLOR='#D0FFFF'>Printer Queue</TD>")
                    Case 2
                        Call WriteLine("<TD BGCOLOR='#D0FFFF'>Device</TD>")
                    Case 3
                        Call WriteLine("<TD BGCOLOR='#D0FFFF'>IPC</TD>")
                    Case -2147483648
                        Call WriteLine("<TD BGCOLOR='#D0FFFF'>Hidden Disk</TD>")
                    Case -2147483647
                        Call WriteLine("<TD BGCOLOR='#D0FFFF'>Hidden Printer Queue</TD>")
                    Case -2147483646
                        Call WriteLine("<TD BGCOLOR='#D0FFFF'>Hidden Device</TD>")
                    Case -2147483645
                        Call WriteLine("<TD BGCOLOR='#D0FFFF'>Hidden IPC</TD>")
                    Case else
                        Call WriteLine("<TD BGCOLOR='#D0FFFF'>Unknown</TD>")
                        strType (i) = "Unknown"
                End Select
                Call WriteLine("<TD BGCOLOR='#D0FFD0'>" & objInstance.Description & "</TD>")
                Call WriteLine("<TD BGCOLOR='#D0FFFF'>" & objInstance.Path & "</TD></TR>" & vbCRLF)
            Next

            If i > 0 Then

            Else
                strMessage = "No share is found."
                WriteLine strMessage
           End If
    End Select
	WriteLine "</TABLE></DIV>"
	
End Sub
