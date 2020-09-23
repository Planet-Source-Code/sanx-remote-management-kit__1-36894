OPTION EXPLICIT

    CONST CONST_LIST                    = "LIST"
    CONST CONST_START                   = "START"
    CONST CONST_STOP                    = "STOP"
    CONST CONST_INSTALL                 = "INSTALL"
    CONST CONST_REMOVE                  = "REMOVE"
    CONST CONST_DEPENDS                 = "DEPENDS"
    CONST CONST_MODE                    = "MODE"


    'Declare variables
    Dim strTaskCommand, strServiceName, strExecName
    Dim strStartMode, strDisplayName

 
Sub Service(strTaskCommand, _ 
                    strServiceName,    _
                    strExecName,       _
                    strDisplayName,    _
                    strStartMode,      _
                    strServer)

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
    Call ExecuteMethod(objService,        _
	                   strTaskCommand, _
	                   strServiceName,    _
	                   strExecName,       _
    				   strDisplayName,    _
                       strStartMode,      _
					   strServer)

End Sub

Private Sub ExecuteMethod(objService,        _
                          strTaskCommand, _
                          strServiceName,    _
						  strExecName,       _
                          strDisplayName,    _
                          strStartMode,      _
                 		  strServer)

    ON ERROR RESUME NEXT

	Dim objEnumerator, objInstance, strMessage, intStatus, objReference
    ReDim strName(0), strDisplayName(0), strState(0), intOrder(0)

	'Initialize local variables
    strMessage        = ""
    strName(0)        = ""
    strDisplayName(0) = ""
    strState(0)       = ""
    intOrder(0)       = 0
	
    Select Case strTaskCommand
        Case CONST_START
            Set objInstance = objService.Get("Win32_Service='" &_
                strServiceName & "'")
            If Err.Number Then
                call document.write( "Error 0x" & CStr(Hex(Err.Number)) & _
                    " occurred in getting " & _
                      "service " & strServiceName & ".")
                If Err.Description <> "" Then
                    call document.write( "Error description: " & Err.Description & ".")
                End If
                Err.Clear
                Exit Sub
            End If
            If objInstance is nothing Then
                Exit Sub
            Else
                intStatus = objInstance.StartService()
                If intStatus = 0 Then
                    strMessage = "Succeeded in starting service " & _
                        strServiceName & "."
                Else
                    strMessage = "Failed to start service " & _
                        strServiceName & "."
                End If
                WriteLine strMessage
            End If

        Case CONST_STOP
            Set objInstance = objService.Get("Win32_Service='" & _
                 strServiceName&"'")
            If Err.Number Then
                call document.write( "Error 0x" & CStr(Hex(Err.Number)) & _
                    " occurred in getting " & _
                      "service " & strServiceName & ".")
                Err.Clear
                Exit Sub
            End If
            If objInstance is nothing Then
                Exit Sub
            Else
                intStatus = objInstance.StopService()
                If intStatus = 0 Then
                    strMessage = "Succeeded in stopping service " & _
                        strServiceName & "."
                Else
                    strMessage = "Failed to stop service " & _
                        strServiceName & "."
                End If
                WriteLine strMessage
            End If

        Case CONST_MODE

            Set objInstance = objService.Get("Win32_Service='" & _
                strServiceName & "'")
            If Err.Number Then
                call document.write( "Error 0x" & CStr(Hex(Err.Number)) & _
                    " occurred in getting " & _
                    "service " & strServiceName & ".")
                Err.Clear
                Exit Sub
            End If
            If objInstance is nothing Then
                Exit Sub
            Else
                intStatus = objInstance.ChangeStartMode(strStartMode)
                If intStatus = 0 Then
                    strMessage = "Succeeded in changing start mode of the " _
                        & "service " & strServiceName & "."
                Else
                    strMessage = "Failed to change the start mode of the" _
                        & " service " & strServiceName & "."
                End If
                WriteLine strMessage
            End If

        Case CONST_INSTALL
            Set objInstance = objService.Get("Win32_Service")
            If Err.Number Then
                call document.write( "Error 0x" & CStr(Hex(Err.Number)) & _
                " occurred in getting " & _
                      "service " & strServiceName & ".")
                If Err.Description <> "" Then
                    call document.write( "Error description: " & Err.Description & ".")
                End If
                Err.Clear
                Exit Sub
            End If
            If objInstance is Nothing Then
                Exit Sub
            Else

                If IsEmpty(strDisplayName) then strDisplayName = strServiceName

                intStatus = objInstance.Create(strServiceName, strDisplayName(0), strExecName)
                If intStatus = 0 Then
                    strMessage = "Succeeded in creating service " & _
                        strServiceName & "."
                Else
                    strMessage = "Failed to create service " & _
                        strServiceName & "."
                End If
                WriteLine strMessage
            End If

        Case CONST_REMOVE
            Set objInstance = objService.Get("Win32_Service='" & _
                strServiceName & "'")
            If Err.Number Then
                call document.write( "Error 0x" & CStr(Hex(Err.Number)) & _
                    " occurred in getting " & _
                    "service " & strServiceName & ".")
                If Err.Description <> "" Then
                    call document.write( "Error description: " & Err.Description & ".")
                End If
                Err.Clear
                Exit Sub
            End If
            If objInstance is Nothing Then
                Exit Sub
            Else
                intStatus = objInstance.Delete()
                If intStatus = 0 Then
                    strMessage = "Succeeded in deleting service " & _
                        strServiceName & "."
                Else
                    strMessage = "Failed to delete service " & _
                        strServiceName & "."
                End If
                WriteLine strMessage
            End If

        Case CONST_DEPENDS
            Set objInstance = objService.Get("Win32_Service='" & _
                strServiceName&"'")
            If Err.Number Then
                call document.write( "Error 0x" & CStr(Hex(Err.Number)) & _
                    " occurred in getting " & _
                      "service " & strServiceName & ".")
                If Err.Description <> "" Then
                    call document.write( "Error description: " & Err.Description & ".")
                End If
                Err.Clear
                Exit Sub
            End If
            If objInstance is Nothing Then
                Exit Sub
            Else
                Set objEnumerator = _
                    objInstance.References_("Win32_DependentService")

                If Err.Number Then
                    call document.write( "Error 0x" & CStr(Hex(Err.Number)) & _
                        " occurred in getting " & _
                        "reference set.")
                    If Err.Description <> "" Then
                        call document.write( "Error description: " & _
                            Err.Description & ".")
                    End If
                    Err.Clear
                    Exit Sub
                End If
			
                If objEnumerator.Count = 0 then
                    document.write "No dependents listed"
			    Else
                    i = 0
                    For Each objReference in objEnumerator
                        If objInstance is nothing Then
                            Exit Sub
                        Else
                            ReDim Preserve strName(i)
                            ReDim strDisplayName(i), strState(i), intOrder(i)
                            strName(i) = _
                                objService.Get(objReference.Dependent).Name
                            strDisplayName(i) = _
                                objService.Get _
                                    (objReference.Dependent).DisplayName
                            strState(i) = _
                                objService.Get(objReference.Dependent).State
                            intOrder(i) = i
                            i = i + 1
                        End If
                        If Err.Number Then
                            Err.Clear
                        End If
                    Next

                   'Display the header
                    WriteLine "<DIV STYLE='font-family: tahoma; font-size: 10pt'>" & _
					"<TABLE CELLPADDING='2'><TR><TH COLSPAN=3 BGCOLOR='#D0D0FF'>" & _
					"Dependents of Service: " & strServiceName & "</TH></TR>"
               	 	strMessage = "<TR><TD BGCOLOR='#D0D0FF'>Name</TD>"
               	 	strMessage = strMessage & "<TD BGCOLOR='#D0D0FF'>State</TD>"
               	 	strMessage = strMessage & "<TD BGCOLOR='#D0D0FF'>Display Name</TD></TR>"
                	WriteLine strMessage
                    Call SortArray(strName, True, intOrder, 0)
                    Call ReArrangeArray(strDisplayName, intOrder)
                    Call ReArrangeArray(strState, intOrder)
                    For i = 0 To UBound(strName)
                        strMessage = "<TR><TD BGCOLOR='#D0FFD0'>" & strName(i) & _
					"</TD><TD BGCOLOR='#D0FFFF'>"
                    strMessage = strMessage & strState(i) & "</TD><TD BGCOLOR='#FFFFd0'>"
                    strMessage = strMessage & strDisplayName(i) & "</TD></TR>" & vbCRLF
                    WriteLine strMessage
                    Next
					WriteLine "</TABLE></DIV>"
                End If
            End If

        Case CONST_LIST
            Set objEnumerator = objService.ExecQuery ( _
                "Select Name,DisplayName,State From Win32_Service")
            If Err.Number Then
                call document.write( "Error 0x" & CStr(Hex(Err.Number)) & _
                    " occurred during the query.")
                If Err.Description <> "" Then
                    call document.write( "Error description: " & Err.Description & ".")
                End If
                Err.Clear
                Exit Sub
            End If
            i = 0
            For Each objInstance in objEnumerator
                If objInstance is nothing Then
                    Exit Sub
                Else
                    ReDim Preserve strName(i), strDisplayName(i)
                    ReDim Preserve strState(i), intOrder(i)
                    strName(i) = objInstance.Name
                    strDisplayName(i) = objInstance.DisplayName
                    strState(i) = objInstance.State
                    intOrder(i) = i
                    i = i + 1
                End If
                If Err.Number Then
                    Err.Clear
                End If
            Next
			
            If i > 0 Then
                'Display the header
				WriteLine "<DIV STYLE='font-family: tahoma; font-size: 10pt'>" & _
				"<TABLE CELLPADDING='2'><TR><TH COLSPAN=3 BGCOLOR='#D0D0FF'>" & _
				"Services installed on server: " & UCase(strServer) & "</TH></TR>"
                strMessage = "<TR><TD BGCOLOR='#D0D0FF'>Name</TD>"
                strMessage = strMessage & "<TD BGCOLOR='#D0D0FF'>State</TD>"
                strMessage = strMessage & "<TD BGCOLOR='#D0D0FF'>Display Name</TD></TR>"
                WriteLine strMessage
                Call SortArray(strName, True, intOrder, 0)
                Call ReArrangeArray(strDisplayName, intOrder)
                Call ReArrangeArray(strState, intOrder)
                For i = 0 To UBound(strName)
                    strMessage = "<TR><TD BGCOLOR='#D0FFD0'>" & strName(i) & _
					"</TD><TD BGCOLOR='#D0FFFF'>"
                    strMessage = strMessage & strState(i) & "</TD><TD BGCOLOR='#FFFFd0'>"
                    strMessage = strMessage & strDisplayName(i) & "</TD></TR>" & vbCRLF
                    WriteLine strMessage
                Next
            Else
                document.write "Service not found!"
            End If
			WriteLine "</TABLE></DIV>"
			
    End Select

End Sub

Private Sub Swap(ByRef strA, ByRef strB)

    Dim strTemp

    strTemp = strA
    strA = strB
    strB = strTemp

End Sub

Private Sub ReArrangeArray(strArray, intOrder)

    ON ERROR RESUME NEXT

    Dim intUBound, i, strTempArray()

    If Not (IsArray(strArray) and IsArray(intOrder)) Then
        call document.write( "At least one of the arguments is not an array")
        Exit Sub
    End If

    intUBound = UBound(strArray)

    If intUBound <> UBound(intOrder) Then
        call document.write( "The upper bound of these two arrays do not match!")
        Exit Sub
    End If

    ReDim strTempArray(intUBound)

    For i = 0 To intUBound
        strTempArray(i) = strArray(intOrder(i))
        If Err.Number Then
            call document.write( "Error 0x" & CStr(Hex(Err.Number)) & " occurred in " _
                      & "rearranging an array.")
            If Err.Description <> "" Then
                call document.write( "Error description: " & Err.Description & ".")
            End If
            Err.Clear
            Exit Sub
        End If
    Next

    For i = 0 To intUBound
        strArray(i) = strTempArray(i)
    Next

End Sub
