Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim CurrentDirectory
Dim Count
Count = 0
CurrentDirectory = objFSO.GetAbsolutePathName(".")
Set oFolder = objFSO.GetFolder(CurrentDirectory)

For Each oFile in oFolder.Files
	If Right(oFile.Name, 3) = "xml" Then
		Set objExcel = CreateObject("Excel.Application")
		Set objWorkbook = objExcel.Workbooks.Open(CurrentDirectory & "\" & oFile.Name)
		objWorkbook.SaveAs CurrentDirectory & "\" & Left(oFile.Name, Len(oFile.Name)-4) & ".csv", 6
		objWorkbook.Close
		Count = Count + 1
	End If
Next

MsgBox "Finished! Converted " & Count & " files"

