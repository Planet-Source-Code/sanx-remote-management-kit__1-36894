OPTION EXPLICIT

    Dim intSortProperty, intWidth, intSortOrder
    ReDim strProperties(2), intWidths(2)

Sub ListJobs(strServer)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, strQuery, strMessage
    Dim objEnumerator, objInstance
    Dim k, i, j, intUBound

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

    'Set the query string.
    strQuery = "Select processid, name, executablepath From Win32_Process"

    'Now execute the query.

    intUBound = UBound(strProperties)
    'Need to use redim so the last dimension can be resized
    ReDim strResults(intUBound, 0), intOrder(0), strArray(0)

    Set objEnumerator = objService.ExecQuery(strQuery,,0)
    If Err.Number Then
        Print "Error 0x" & CStr(Hex(Err.Number)) & " occurred during the query."
        If Err.Description <> "" Then
            Print "Error description: " & Err.Description & "."
        End If
        Err.Clear
        Exit Sub
    End If

    'Properties to get
    strProperties(0) = "processid"
    strProperties(1) = "name"
    strProperties(2) = "executablepath"
    intWidths(1) = 15
    intWidths(0) = 15
    intWidths(2) = 40

    'Read properties of processes into arrays.
    i = 0
    For Each objInstance in objEnumerator
        If objInstance is nothing Then
            Exit For
        End If
        ReDim Preserve strResults(intUBound, i), intOrder(i), strArray(i)
        For j = 0 To intUBound
            Select Case LCase(strProperties(j)) 
                Case "processid" 
                    strResults(j, i) = objInstance.properties_(strProperties(j))
                    If strResults(j, i) < 0 Then
                        '4294967296 is 0x100000000.
                        strResults(j, i) = CStr(strResults(j, i) + 4294967296)
                    End If
                Case "owner"
                    Dim strDomain, strUser
                    Call objInstance.GetOwner(strUser, strDomain)
                    strResults(j, i) = strDomain & "\" & strUser
                Case Else
                    strResults(j, i) = CStr _
                        (objInstance.properties_(strProperties(j)))
            End Select
            If Err.Number Then
                Err.Clear
                strResults(j, i) = "(null)"
            End If
        Next
        intOrder(i) = i
        'Copy the property values to be sorted.
        strArray(i) = strResults(0, i)
        i = i + 1
        If Err.Number Then
            Err.Clear
        End If
    Next

    'Check the data type of the property to be sorted
    k = CDbl(strArray(0))
    If Err.Number Then      'not a number
        Err.Clear
    End If

    If i > 0 Then
        'Print the header
        WriteLine "<P><P></P></P>" & _
			"<DIV STYLE='font-family: tahoma, arial, sans-serif; font-size: 10pt'>" & _
			"<TABLE><TR><TH COLSPAN=3 BGCOLOR='#D0D0FF'>Process Information for server: " & _ 
			strServer & "</TH></TR>"
		WriteLine "<TR>"
        For j = 0 To intUBound
            strMessage = strMessage & "<TD BGCOLOR='#D0D0FF'>" & UCase(strProperties(j)) & "</TD>"
        Next
        WriteLine strMessage & "</TR>" & vbCRLF

        'Sort strArray
        Call SortArray(strArray, 1, intOrder, 0)

            For j = 0 To intUBound
                'First copy results to strArray and change the order of elements
                For k = 0 To i-1    'i is number of instances retrieved.
                    strArray(k) = strResults(j, intOrder(k))
                Next
                'Now copy results back to strResults.
                For k = 0 To i-1    'i is number of instances retrieved.
                    strResults(j, k) = strArray(k)
                Next
            Next

        For k = 0 To i-1
            strMessage = "<TR>"
            For j = 0 To intUBound
                strMessage = strMessage & "<TD>" & strResults(j, k) & "</TD>"
            Next
			strMessage = strMessage & "</TR>"
            WriteLine strMessage
        Next
    End If
	
	WriteLine "</TABLE></DIV>"

End Sub


