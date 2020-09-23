Function strInsertCommas(intValue)

    Dim IntPlace

    strInsertCommas = ""
    intPlace = 0 
    For I = len(intValue) to 1 Step - 1
        strInsertCommas = Mid(intValue,I,1) & strInsertCommas
        intPlace = intPlace + 1
        If intPlace = 3 then
            strInsertCommas = "," & strInsertCommas
            intPlace = 0
        End If
    Next
    If Left(strInsertCommas,1) = "," then
        strInsertCommas = Right(strInsertCommas, len(strInsertCommas) -1)
    End If

End Function


Private Function strPackString( ByVal strString, _
                                ByVal intWidth,  _
                                ByVal blnAfter,  _
                                ByVal blnTruncate)

    ON ERROR RESUME NEXT

    intWidth      = CInt(intWidth)
    blnAfter      = CBool(blnAfter)
    blnTruncate   = CBool(blnTruncate)

    If Err.Number Then
        Call document.write ("Argument type is incorrect!")
        Err.Clear
        Wscript.Quit
    End If

    If IsNull(strString) Then
        strPackString = "null" & Space(intWidth-4)
        Exit Function
    End If

    strString = CStr(strString)
    If Err.Number Then
        Call document.write ("Argument type is incorrect!")
        Err.Clear
        Wscript.Quit
    End If

    If intWidth > Len(strString) Then
        If blnAfter Then
            strPackString = strString & Space(intWidth-Len(strString))
        Else
            strPackString = Space(intWidth-Len(strString)) & strString & " "
        End If
    Else
        If blnTruncate Then
            strPackString = Left(strString, intWidth-1) & " "
        Else
            strPackString = strString & " "
        End If
    End If

End Function


Private Function blnGetArg ( ByVal StrVarName,   _
                             ByRef strVar,       _
                             ByRef intArgIter) 

    blnGetArg = False 'failure, changed to True upon successful completion

    If Len(Wscript.Arguments(intArgIter)) > 2 then
        If Mid(Wscript.Arguments(intArgIter),3,1) = ":" then
            If Len(Wscript.Arguments(intArgIter)) > 3 then
                strVar = Right(Wscript.Arguments(intArgIter), _
                         Len(Wscript.Arguments(intArgIter)) - 3)
                blnGetArg = True
                Exit Function
            Else
                intArgIter = intArgIter + 1
                If intArgIter > (Wscript.Arguments.Count - 1) Then
                    Call document.write( "Invalid " & StrVarName & ".")
                    Call document.write( "Please check the input and try again.")
                    Exit Function
                End If

                strVar = Wscript.Arguments.Item(intArgIter)
                If Err.Number Then
                    Call document.write( "Invalid " & StrVarName & ".")
                    Call document.write( "Please check the input and try again.")
                    Exit Function
                End If

                If InStr(strVar, "/") Then
                    Call document.write( "Invalid " & StrVarName)
                    Call document.write( "Please check the input and try again.")
                    Exit Function
                End If

                blnGetArg = True 'success
            End If
        Else
            strVar = Right(Wscript.Arguments(intArgIter), _
                     Len(Wscript.Arguments(intArgIter)) - 2)
            blnGetArg = True 'success
            Exit Function
        End If
    Else
        intArgIter = intArgIter + 1
        If intArgIter > (Wscript.Arguments.Count - 1) Then
            Call document.write( "Invalid " & StrVarName & ".")
            Call document.write( "Please check the input and try again.")
            Exit Function
        End If

        strVar = Wscript.Arguments.Item(intArgIter)
        If Err.Number Then
            Call document.write( "Invalid " & StrVarName & ".")
            Call document.write( "Please check the input and try again.")
            Exit Function
        End If

        If InStr(strVar, "/") Then
            Call document.write( "Invalid " & StrVarName)
            Call document.write( "Please check the input and try again.")
            Exit Function
        End If
        blnGetArg = True 'success
    End If
End Function


Private Function blnConnect(ByVal strNameSpace, _
                            ByVal strUserName,  _
                            ByVal strPassword,  _
                            ByRef strServer,    _
                            ByRef objService)

    ON ERROR RESUME NEXT

    Dim objLocator, objWshNet

    blnConnect = False     'There is no error.

    'Create Locator object to connect to remote CIM object manager
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    If Err.Number then
        Call document.write( "Error 0x" & CStr(Hex(Err.Number)) & _
                           " occurred in creating a locator object." )
        If Err.Description <> "" Then
            Call document.write( "Error description: " & Err.Description & "." )
        End If
        Err.Clear
        blnConnect = True     'An error occurred
        Exit Function
    End If

    'Connect to the namespace which is either local or remote
    Set objService = objLocator.ConnectServer (strServer, strNameSpace, _
       strUserName, strPassword)
    ObjService.Security_.impersonationlevel = 3
    If Err.Number then
        Call document.write( "Error 0x" & CStr(Hex(Err.Number)) & _
                           " occurred in connecting to server " _
           & strServer & ".")
        If Err.Description <> "" Then
            Call document.write( "Error description: " & Err.Description & "." )
        End If
        Err.Clear
        blnConnect = True     'An error occurred
    End If

    'Get the current server's name if left unspecified
    If IsEmpty(strServer) Then
        Set objWshNet = CreateObject("Wscript.Network")
    strServer     = objWshNet.ComputerName
    End If

End Function


Sub WriteLine(ByVal strMessage)

        Call document.write( strMessage )

End Sub


Private Function blnErrorOccurred (ByVal strIn)

    If Err.Number Then
        Call document.write( "Error 0x" & CStr(Hex(Err.Number)) & ": " & strIn)
        If Err.Description <> "" Then
            Call document.write( "Error description: " & Err.Description)
        End If
        Err.Clear
        blnErrorOccurred = True
    Else
        blnErrorOccurred = False
    End If

End Function


Private Sub SortArray(strArray, blnOrder, strArray2, blnCase)

    ON ERROR RESUME NEXT

    Dim i, j, intUbound

    If IsArray(strArray) Then
        intUbound = UBound(strArray)
    Else
        Print "Argument is not an array!"
        Exit Sub
    End If

    blnOrder = CBool(blnOrder)
    blnCase = CBool(blnCase)
    If Err.Number Then
        Print "Argument is not a boolean!"
        Exit Sub
    End If

    i = 0
    Do Until i > intUbound-1
        j = i + 1
        Do Until j > intUbound
            If blnCase Then     'Case sensitive
                If (strArray(i) > strArray(j)) and blnOrder Then
                    Swap strArray(i), strArray(j)   'swaps element i and j
                    Swap strArray2(i), strArray2(j)
                ElseIf (strArray(i) < strArray(j)) and Not blnOrder Then
                    Swap strArray(i), strArray(j)   'swaps element i and j
                    Swap strArray2(i), strArray2(j)
                ElseIf strArray(i) = strArray(j) Then
                    'Move element j to next to i
                    If j > i + 1 Then
                        Swap strArray(i+1), strArray(j)
                        Swap strArray2(i+1), strArray2(j)
                    End If
                End If
            Else                 'Not case sensitive
                If (LCase(strArray(i)) > LCase(strArray(j))) and blnOrder Then
                    Swap strArray(i), strArray(j)   'swaps element i and j
                    Swap strArray2(i), strArray2(j)
                ElseIf (LCase(strArray(i)) < LCase(strArray(j))) and _
                        Not blnOrder Then
                    Swap strArray(i), strArray(j)   'swaps element i and j
                    Swap strArray2(i), strArray2(j)
                ElseIf LCase(strArray(i)) = LCase(strArray(j)) Then
                    'Move element j to next to i
                    If j > i + 1 Then
                        Swap strArray(i+1), strArray(j)
                        Swap strArray2(i+1), strArray2(j)
                    End If
                End If
            End If
            j = j + 1
        Loop
        i = i + 1
    Loop

End Sub

Private Sub Swap(ByRef strA, ByRef strB)

    Dim strTemp

    strTemp = strA
    strA = strB
    strB = strTemp
	strb = 4/0
End Sub

Private Sub HTMLHeaders()

	ON ERROR RESUME NEXT

	Dim strMessage

	strMessage = "<HTML><HEAD><TITLE>Sanx's Computer Management Interface :: " & _
			"Query Results</TITLE>" & vbCRLF
	strMessage = strMessage & "<STYLE>BODY {font-family: tahoma;font-size: 10pt;}" & _
			"TD {font-family: tahoma;font-size: 10pt;}" & _
			"TH {font-family: tahoma;font-size: 12pt;font-weight: bold;}" & _
			"</STYLE></HEAD>" & vbCRLF
	strMessage = strMessage & "<BODY BGCOLOR='white'><CENTER>"

	WriteLine strMessage

End Sub