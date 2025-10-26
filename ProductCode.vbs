strComputer = "." 
Set objFSO=CreateObject("Scripting.FileSystemObject")
outFile="ProductCode.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_BaseBoard",,48) 
For Each objItem in colItems 
    Wscript.Echo "SystemID " & objItem.Product
    objFile.Write "SystemID: " & objItem.Product & vbCrLf
    objFile.Write "Serial Number: " & objItem.SerialNumber & vbCrLf
    objFile.Write "SKU: " & objItem.Description & vbCrLf
Next

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colSMBIOS = objWMIService.ExecQuery ("Select * from Win32_ComputerSystemProduct") 
For Each objSMBIOS in colSMBIOS
    objFile.Write "Product Name: " & objSMBIOS.Name & vbCrLf
    objFile.Write "UUID: " & objSMBIOS.UUID & vbCrLf
    objFile.Write "SKU Number: " & objSMBIOS.SKUNumber & vbCrLf
Next

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colSMBIOS = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem") 
For Each objSMBIOS in colSMBIOS
    objFile.Write "SKU Number: " & objSMBIOS.ChassisSKUNumber & vbCrLf
Next

objFile.Close
