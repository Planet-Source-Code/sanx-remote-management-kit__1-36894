Private Sub Drives(strServer)

    ON ERROR RESUME NEXT

    Dim objFileSystem, objService, objset, objInst
    Dim strLine

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

    'Get the first instance
    Set objSet = objService.InstancesOf("Win32_DiskDrive")
    If blnErrorOccurred ("obtaining the Win32_DiskDrive.") Then Exit Sub

	Call WriteLine("<DIV STYLE='font-family: tahoma, arial, sans-serif; font-size: 10pt'>" & _
				"<TABLE><TR><TH COLSPAN=2 BGCOLOR='#D0D0FF'>Disk Drives on server: " & _ 
				UCase(strServer) & "</TH></TR><TR><TD>&nbsp;</TD></TR>" & vbCRLF)
	
    For Each objInst In objSet
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>" & objInst.Caption & _
		"</TD><TD BGCOLOR='#D0FFFF'>" & objInst.Description  & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Status</TD><TD BGCOLOR='#D0FFFF'>" & objInst.Status & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Media Loaded</TD><TD BGCOLOR='#D0FFFF'>" & _
		strYesOrNo(objInst.MediaLoaded) & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Partitions</TD><TD BGCOLOR='#D0FFFF'>" & _
		CStr(objInst.Partitions) & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>System Name</TD><TD BGCOLOR='#D0FFFF'>" & _
		objInst.SystemName & "</TD></TR>" & vbCRLF

        if objInst.InterfaceType = "SCSI" Then
            WriteLine "<TR><TD BGCOLOR='#D0FFD0'>SCSIBus</TD><TD BGCOLOR='#D0FFFF'>" & _
			CStr(objInst.SCSIBus) & "</TD></TR>" & vbCRLF
            WriteLine "<TR><TD BGCOLOR='#D0FFD0'>SCSILogicalUnit</TD><TD BGCOLOR='#D0FFFF'>" & _
			CStr(objInst.SCSILogicalUnit) & "</TD></TR>" & vbCRLF
            WriteLine "<TR><TD BGCOLOR='#D0FFD0'>SCSIPort</TD><TD BGCOLOR='#D0FFFF'>" & _
			CStr(objInst.SCSIPort) & "</TD></TR>" & vbCRLF
            WriteLine "<TR><TD BGCOLOR='#D0FFD0'>SCSITargetId</TD><TD BGCOLOR='#D0FFFF'>" & _
			CStr(objInst.SCSITargetId) & "</TD></TR>" & vbCRLF
        End If

        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Manufacturer/Model</TD><TD BGCOLOR='#D0FFFF'>" & _
             objInst.Manufacturer & " " & objInst.Model & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Size</TD><TD BGCOLOR='#D0FFFF'>"& _
             strInsertCommas(objInst.Size) & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Total Cylinders</TD><TD BGCOLOR='#D0FFFF'>" & _
             strInsertCommas(objInst.TotalCylinders) & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Total Heads</TD><TD BGCOLOR='#D0FFFF'>" & _
             strInsertCommas(CStr(objInst.TotalHeads)) & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Total Sectors</TD><TD BGCOLOR='#D0FFFF'>" & _
             strInsertCommas(objInst.TotalSectors) & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Total Tracks</TD><TD BGCOLOR='#D0FFFF'>" & _
             strInsertCommas(objInst.TotalTracks) & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Sectors Per Track</TD><TD BGCOLOR='#D0FFFF'>" & _
             strInsertCommas(objInst.SectorsPerTrack) & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Tracks Per Cylinder</TD><TD BGCOLOR='#D0FFFF'>" & _
             strInsertCommas(CStr(objInst.TracksPerCylinder)) & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Bytes Per Sector</TD><TD BGCOLOR='#D0FFFF'>" & _
             strInsertCommas(objInst.BytesPerSector) & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Name</TD><TD BGCOLOR='#D0FFFF'>" & _
             objInst.Name & "</TD></TR>" & vbCRLF
        WriteLine "<TR><TD BGCOLOR='#D0FFD0'>Creation Class Name</TD><TD BGCOLOR='#D0FFFF'>" & _
             objInst.CreationClassName & "</TD></TR><TR><TD>&nbsp;</TD></TR>" & vbCRLF
    Next
	Call WriteLine("</TABLE></DIV>" & vbCRLF)

End Sub
