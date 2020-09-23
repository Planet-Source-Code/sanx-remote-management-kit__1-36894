OPTION EXPLICIT

Private Sub ListPrinters(strServer)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, objPrinterSet, objPrinter
    Dim strWBEMClass, strPrinterAttributes
    strPrinterAttributes = Array("Queued"             , _
                                 "Direct"             , _
                                 "Default"            , _
                                 "Shared"             , _ 
                                 "Network"            , _
                                 "Hidden"             , _
                                 "Local"              , _
                                 "Enable DevQ"        , _
                                 "Keep Printed Jobs"  , _
                                 "Do Complete First"  , _
                                 "Work Offline"       , _
                                 "Enable BIDI"        , _
                                 "Raw Only"             )

    Call HTMLHeaders

    strWBEMClass = "Win32_Printer"

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

    'Get the set
    Set objPrinterSet = objService.InstancesOf(strWBEMClass)
    If blnErrorOccurred("Could not obtain " & strWBEMClass & " instance.") Then
        Exit Sub
    End If

    If objPrinterSet.Count = 0 Then
        Call WriteLine("No printers information is available.", _
                        objOutputFile)    
        Exit Sub
    End If

	Call WriteLine("<DIV STYLE='font-family: tahoma, arial, sans-serif; font-size: 10pt'>" & _
				"<TABLE><TR><TH COLSPAN=2 BGCOLOR='#D0D0FF'>Printers installed on server: " & _ 
				UCase(strServer) & "</TH></TR>" & vbCRLF)
    Call WriteLine("<TR><TD>&nbsp;</TD></TR>" & vbCRLF)

    For Each objPrinter In objPrinterSet
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Name</TD><TD BGCOLOR='#D0FFFF'>" & objPrinter.Caption & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Location</TD><TD BGCOLOR='#D0FFFF'>" & objPrinter.Location & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Description</TD><TD BGCOLOR='#D0FFFF'>" & objPrinter.Description & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Driver</TD><TD BGCOLOR='#D0FFFF'>" & objPrinter.DriverName & "</TD></TR>" & vbCRLF)
        Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Error State</TD><TD BGCOLOR='#D0FFFF'>" & objPrinter.DetectedErrorState & "</TD></TR>" & vbCRLF)
        For I = 1 to len(objPrinter.Attributes)
            If I = 1 then
                Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>Attributes</TD><TD BGCOLOR='#D0FFFF'>" & _
				strPrinterAttributes & (Mid(objPrinter.Attributes,I,1)) & "</TD></TR>" & vbCRLF)
            Else
                Call WriteLine("<TR><TD BGCOLOR='#D0FFD0'>" & strPrinterAttributes _
                                & "</TD><TD BGCOLOR='#D0FFFF'>" & _
								(Mid(objPrinter.Attributes,I,1)) & "</TD></TR>" & vbCRLF)            
            End If
        Next
		Call WriteLine("<TR><TD>&nbsp;</TD></TR>" & vbCRLF)
    Next
	Call WriteLine("</TABLE></DIV>" & vbCRLF)

End Sub
