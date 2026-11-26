Attribute VB_Name = "Module2"
Sub Create_report_payment()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wt As ListObject
    Dim xlApp As Object
    Dim ACC_report_path As String
    Dim password As String
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim tbl As ListObject
    Dim SourceRange As Range
    Dim targetRange As Range
    Dim startRow As Long
    Dim row_count As Long
    Dim lr As Long
    Dim SelectedFile As String
    Dim savePath As String
    Dim today As String
    Dim folderPath As String
    Dim latestFile As String
    Dim latestDate As Date
    Dim i As Long

    folderPath = "C:\card_data\payment\raw_email_attached" ' <-- Adjust this path as needed
    
    ' Value in this worksheet table
    Set targetWb = ThisWorkbook
    Set targetWs = targetWb.Sheets("Payment_Term_History")
    Set tbl = targetWs.ListObjects("Payment")
    
    ' Try to find the latest Excel file in the folder
    If Dir(folderPath, vbDirectory) <> "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set folder = fso.GetFolder(folderPath)
    
        For Each file In folder.Files
            ' Check file extension and naming pattern
            If (LCase(Right(file.Name, 5)) = ".xlsx" Or LCase(Right(file.Name, 4)) = ".xls") _
               And LCase(file.Name) Like "confirm_payment_for_load_*.xlsx" Then
               
                fileDate = file.DateLastModified
                If fileDate > latestDate Then
                    latestDate = fileDate
                    latestFile = file.Path
                End If
            End If
        Next file
    End If
    
    ' Use found file or let user select
    If latestFile <> "" Then
        SelectedFile = latestFile
        ACC_report_path = SelectedFile
    Else
        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "Select the Excel file manually"
            .Filters.Clear
            .Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsm"
            If .Show = -1 Then
                SelectedFile = .SelectedItems(1)
                ACC_report_path = SelectedFile
            Else
                MsgBox "No file selected. Exiting."
                Exit Sub
            End If
        End With
    End If
    
    ' Set Excel instance
    Set xlApp = CreateObject("Excel.Application")
    
    ' Open the selected workbook
    On Error Resume Next
    Set wb = xlApp.Workbooks.Open(ACC_report_path)
    On Error GoTo 0
    
    ' Check if the workbook was successfully opened
    If wb Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
        Exit Sub
    End If
    
    ' Set the worksheet
    Set ws = wb.Sheets(targetWs.Range("AA24").Value)
    
    ' Define the starting row where you want to write data
    startRow = 2 ' Change to the row where you want to start writing
    
    ' Get the last row of data in column D
    row_count = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    
    ' Copy ID code
    Set SourceRange = ws.Range("D2:D" & row_count)
    Set targetRange = targetWs.Cells(startRow, 4)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    
    ' Copy Payment
    Set SourceRange = ws.Range("G2:G" & row_count)
    Set targetRange = targetWs.Cells(startRow, 13)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    
    ' Copy Date
    Set SourceRange = ws.Range("I2:I" & row_count)
    Set targetRange = targetWs.Cells(startRow, 15)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    Set SourceRange = ws.Range("H2:H" & row_count)
    Set targetRange = targetWs.Cells(startRow, 16)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    
    ' Copy Payment Remark
    Set SourceRange = ws.Range("J2:K" & row_count)
    Set targetRange = targetWs.Cells(startRow, 17)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    
    ' Copy Payment Detail
    Set SourceRange = ws.Range("N2:N" & row_count)
    Set targetRange = targetWs.Cells(startRow, 21)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    Set SourceRange = ws.Range("O2:O" & row_count)
    Set targetRange = targetWs.Cells(startRow, 20)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    Set SourceRange = ws.Range("P2:P" & row_count)
    Set targetRange = targetWs.Cells(startRow, 9)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    Set SourceRange = ws.Range("R2:R" & row_count)
    Set targetRange = targetWs.Cells(startRow, 10)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    Set SourceRange = ws.Range("S2:S" & row_count)
    Set targetRange = targetWs.Cells(startRow, 23)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
        
    ' Search for "HC 1/" in column Q and add new row if found
    Dim sheet_count As Integer
    sheet_count = 3
    For i = 2 To row_count
        If targetWs.Cells(i, 18).Value Like "*HC*" Then
            Dim temp_text As String
            temp_text = targetWs.Cells(i, 20).Value
            
            ' Extract the numbers and check if they are equal
            Dim parts As Variant
            Dim number1 As Integer
            Dim number2 As Integer
            
            parts = Split(temp_text, " ")
            If UBound(parts) >= 1 Then
                Dim hcPart As String
                hcPart = parts(1)
                
                parts = Split(hcPart, "/")
                If UBound(parts) >= 1 Then
                    number1 = CInt(parts(0))
                    number2 = CInt(parts(1))
                    
                    If number1 = number2 Then
                        ' Insert new row and copy data
                        targetWs.Cells(i, 20).EntireRow.Insert Shift:=xlDown
                        targetWs.Cells(i, 4).Value = targetWs.Rows(i + 1).Cells(4).Value ' Copy ID code
                        targetWs.Cells(i, 13).Value = ws.Cells(sheet_count, 12).Value ' Copy Discount
                        targetWs.Cells(i, 15).Value = Format(DateAdd("d", -1, CDate(targetWs.Rows(i + 1).Cells(15).Value)), "mm/dd/yyyy") ' Copy Date
                        targetWs.Cells(i, 16).Value = targetWs.Cells(i, 15).Value
                        targetWs.Cells(i, 18).Value = targetWs.Rows(i + 1).Cells(18).Value ' Copy Payment Remark
                        targetWs.Cells(i, 20).Value = "Add Transaction"
                        i = i + 1
                    End If
                End If
            End If
        End If
        sheet_count = sheet_count + 1
                
        If targetWs.Cells(i, "D").Value <> vbbullstring And targetWs.Cells(i, 18).Value = vbNullString Then
            targetWs.Cells(i, 9).Value = "Sale"
            targetWs.Cells(i, 10).Value = "Sale"
            targetWs.Cells(i, 18).Value = "Repo/Sold"
        End If
        
    Next i
    
    ' Get the number of rows in the table
    lr = tbl.ListRows.Count
    
    ' Reapply table formatting to ensure formulas are expanded
    tbl.Resize tbl.Range.Resize(tbl.Range.Rows.Count, tbl.Range.Columns.Count)
    tbl.ShowTotals = False ' Reset Totals row
    
    ' Close the source workbook
    wb.Close SaveChanges:=False
    
    ' Prompt the user to save the new workbook
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Save Report As"
        .FilterIndex = 1
        .InitialFileName = "Payment_report_" & Format(Date, "DD-MM-YYYY") & ".xlsx"
        If .Show = -1 Then
            savePath = .SelectedItems(1)
            ' Save the workbook as a new file
            targetWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
        Else
            MsgBox "File save canceled.", vbInformation
            Exit Sub
        End If
    End With
    
    xlApp.Quit
    Set xlApp = Nothing
    
    MsgBox "Create Report Done"

End Sub

