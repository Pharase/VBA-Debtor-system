Attribute VB_Name = "Module3"
Sub SaveWorkbookWithDialog(Name As String, ws As Workbook, password As String)
    Dim filePath As Variant
    
    ' Prompt user to select a location and enter a file name
    filePath = Application.GetSaveAsFilename(InitialFileName:=Name, FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    
    ' Check if the user selected a file
    If filePath <> "False" Then
        ' Save the workbook at the selected location
        ws.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook, password:=password
        MsgBox "File saved successfully at: " & filePath
    Else
        MsgBox "Save operation canceled."
    End If
End Sub

