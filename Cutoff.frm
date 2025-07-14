VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "DataCutOff"
   ClientHeight    =   4032
   ClientLeft      =   -828
   ClientTop       =   -3336
   ClientWidth     =   2256
   OleObjectBlob   =   "Cutoff.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filename As String
Dim folderPath As String
Dim foundFile As String
Dim searchText As String
Dim MainPath As String
Dim MacroPath As String
Dim xlMain As Object
Dim xlMainS As Object
Dim xlApp As Object
Dim xlWb As Object
Dim xlWs As Object
Dim xlMacro As Object
Dim xlMacroS As Object
Dim wb As Object
Dim ws As Object
Dim wt As ListObject
Dim lastRow As Long
Dim lastRowTB As Long
Dim startRow As Integer
Dim v_pr As String
Dim v_st As String
Dim v_port As String
Dim v_case As String
Dim variable1 As String
Dim variable2 As String
Dim selectedFile As String
Dim selectedOAFile As String
Dim targetWb As Workbook
Dim targetWs As Worksheet
Dim tbl As ListObject
Dim tbMac As ListObject
Dim TempName As String
Dim temp_dis As ReturnValue
Dim T_code As ReturnValue
Dim err_code As ReturnValue
Dim lr As Long
Dim mr As Long
Dim i As Long
Dim l As Long
Dim copy_i As Long
Dim response As VbMsgBoxResult
Dim test_mode As Boolean
Dim core_id As String
Dim User_name As String
Private Sub CheckBox1_Click()
    test_mode = Not test_mode
    
    If test_mode Then
        MsgBox "Test_mode On"
    Else
        MsgBox "Test_mode Off"
    End If
End Sub
Private Sub CommandButton1_Click()

    ' Value in this worksheet table
    Set targetWb = ThisWorkbook
    Set targetWs = targetWb.Sheets("Hold_Cutting")
    Set tbl = targetWs.ListObjects("Payment_list")

        ' Allow the user to select a file
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
            TextBox1.Value = Left(selectedFile, InStrRev(selectedFile, Application.PathSeparator)) & Mid(selectedFile, InStrRev(selectedFile, Application.PathSeparator) + 1)
        Else
            Exit Sub ' The user canceled the file picker
        End If
    End With

    ' Find the position of "C:\Users\" in the path
    Dim startPos As Integer
    startPos = InStr(selectedFile, "C:\Users\")

    If startPos > 0 Then
        ' Extract the substring after "C:\Users\"
        User_name = Mid(selectedFile, startPos + Len("C:\Users\"))

        ' Find the next backslash to determine the end of the username
        Dim endPos As Integer
        endPos = InStr(User_name, "\")

        If endPos > 0 Then
            ' Extract the username
            User_name = Left(User_name, endPos - 1)
            ThisWorkbook.Sheets("Hold_Cutting").Range("ak17").Value = User_name
            ' Display the extracted username
            MsgBox "Username: " & User_name
        End If
    End If
    
    'set excel instance
    Set xlApp = CreateObject("Excel.Application")
    'foundFile = FindExcelFile("C:\Users\" & User_name & "\Hylife Group\PAM-Data - Documents\Data For QMC\Payment Term\Cutoff-beta\Payment report\", "Data_Status")
    On Error Resume Next
    
    If test_mode Then
        MainPath = "C:\Pam\Tools\Cut-off\Testing env\_Data_for_Cut_off_Table.xlsx"
        MacroPath = "C:\Pam\Tools\Cut-off\Cut_off\Cut-Off-Database_marco.xlsm"
    Else
        MainPath = TextBox2.Value
        MacroPath = TextBox3.Value
    End If
    On Error GoTo 0
    
    Set xlMain = xlApp.Workbooks.Open(MainPath)
    Set xlMainS = xlMain.Sheets("DATA")
    
    Set xlMacro = xlApp.Workbooks.Open(ThisWorkbook.Sheets("Hold_Cutting").Range("an13").Value)
    Set xlMacroS = xlMacro.Sheets("payment_mode_status")
    
    ' Open the selected workbook
    On Error Resume Next
    Set wb = xlApp.Workbooks.Open(selectedFile, password:=ThisWorkbook.Sheets("Hold_Cutting").Range("ak15").Value)
    Set ws = wb.Sheets("Payment_Term_History")
    Set wt = ws.ListObjects("Payment")
    On Error GoTo 0
    
    'ws.Unprotect Password:=vbNullString
    
    ' Specify the target workbook (ThisWorkbook) and worksheet within the target workbook
    Set targetWb = ThisWorkbook
    Set targetWs = targetWb.Sheets("Hold_Cutting") ' Change to the name of your existing worksheet
    Set tbl = targetWs.ListObjects("Payment_list") ' Replace "YourTableName" with your table name

    ' Define the starting row where you want to write data
    startRow = 6 ' Change to the row where you want to start writing
    
    ' Clear all data within the table
    tbl.DataBodyRange.ClearContents
    tbl.DataBodyRange.ClearFormats
    tbl.ListColumns(1).DataBodyRange.ClearFormats
    tbl.ListColumns(4).DataBodyRange.ClearFormats
    tbl.ListColumns(10).DataBodyRange.ClearFormats
    
    Dim row_count As Long
    row_count = ws.Cells(Rows.Count, 2).End(xlUp).Row 'wt.ListColumns(1).Range.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    'wt.ListColumns(4).DataBodyRange.Cells(wt.ListRows.Count, 1).End(xlUp).Row
    
    ' Copy data from the selected file to the target worksheet
    'ID code
    Set SourceRange = ws.Range("D2:D" & row_count)
    Set targetRange = targetWs.Cells(startRow, 2)
    
    Dim lead_zero As Integer
    lead_zero = 0
    ' Loop through each cell in the source range
    For Each cell In SourceRange
        If Left(cell.Value, 1) = "0" Then
            ' Format the corresponding target cell as text if it has a leading zero
            targetWs.Cells(lead_zero + startRow, 2).NumberFormat = "@"
        End If
        lead_zero = lead_zero + 1
    Next cell
    
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value

    'Old code
    Set SourceRange = ws.Range("E2:E" & row_count)
    Set targetRange = targetWs.Cells(startRow, 3)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    
    ' Copy Name
    Set SourceRange = ws.Range("C2:C" & row_count)
    Set targetRange = targetWs.Cells(startRow, 4)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    
    ' Copy product c_stat c_mode document
    Set SourceRange = ws.Range("F2:J" & row_count)
    Set targetRange = targetWs.Cells(startRow, 6)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    
    ' Copy Port
    Set SourceRange = ws.Range("P2:P" & row_count)
    Set targetRange = targetWs.Cells(startRow, 5)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    
    ' Copy Payment info
    Set SourceRange = ws.Range("K2:M" & row_count)
    Set targetRange = targetWs.Cells(startRow, 11)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value
    targetWs.Range("M6:M" & row_count + startRow).NumberFormat = "dd-mm-yy"
    
    ' Copy Note
    Set SourceRange = ws.Range("Q2:Q" & row_count)
    Set targetRange = targetWs.Cells(startRow, 14)
    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value

    'Old copy method
    'ws.Range("F" & 2 & ":I" & ws.Cells(row_count, "D").End(xlUp).Row).Copy
    'ws.Range("P" & 2 & ":P" & row_count).Copy
    'targetWs.Cells(startRow, 14).PasteSpecial xlPasteValues
    
    'ws.Protect Password:=vbNullString
    
    'MacroPath = TextBox2.Value
    
    ' Macro path
    'Set xlApp = CreateObject("Excel.Application")
    'Set xlMacro = xlApp.Workbooks.Open(MacroPath) 'Password:=ThisWorkbook.Sheets("Hold_Cutting").Range("ad15").Value')
    'Set xlMacroS = xlMacro.Sheets("DATA")
    'Set tbMac = xlMacroS.ListObjects("data_table")
    
    lr = tbl.ListRows.Count ' Get the number of rows in the table
    'mr = tbMac.ListRows.Count ' Get the number of rows in the macro table

    ' Loop through all rows in the target sheet's table
    'For i = 1 To lr
        'Dim searchValue As String
        'searchValue = tbl.ListColumns("Code").DataBodyRange(i).Value
        ' Find the matching row in the macro table (assuming you have a unique identifier column, e.g., "O" column)
        'Set foundCell = tbMac.ListColumns(3).DataBodyRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        'If searchValue = vbNullString Then
            'Exit For
        'End If
        
        'If Not foundCell Is Nothing Then
            'l = foundCell.Row ' Row in the macro table where the match was found
        'copy_i = i + startRow - 1 'Row of value in Program sheet
            
            'add payment code from filename "qmc" = pm
        'If (InStr(1, Dir(selectedFile), "qmc", vbTextCompare) > 0) And (targetWs.Range("K" & copy_i).Value <> vbNullString) Then
            'targetWs.Cells(copy_i, "J").Value = "pm"
        'End If
            
            'Call status and mode origin
            'xlMacroS.Range("H" & l & ":I" & l).Copy
            'targetWs.Range("G" & copy_i).PasteSpecial xlPasteValues
        'End If
    'Next i
    
    wb.Close SaveChanges:=True
    xlMain.Close SaveChanges:=True
    xlMacro.Close SaveChanges:=True
    xlApp.Quit
    Set xlApp = Nothing
    
    MsgBox "Load Payment Done"

End Sub
Private Sub CommandButton2_Click()
    
    On Error Resume Next
    If test_mode Then
        MainPath = "C:\Pam\Tools\Cut-off\Testing env\_Data_for_Cut_off_Table.xlsx"
        MacroPath = "C:\Pam\Tools\Cut-off\Cut_off\Cut-Off-Database_marco.xlsm"
    Else
        MainPath = TextBox2.Value
        MacroPath = ThisWorkbook.Sheets("Hold_Cutting").Range("AM13").Value 'TextBox3.Value
    End If
    On Error GoTo 0
    
    folderPath = Left(MacroPath, InStrRev(MacroPath, "\"))
    TempName = Right(MacroPath, Len(MacroPath) - InStrRev(MacroPath, "\")) 'ThisWorkbook.Sheets("Hold_Cutting").Range("AM13").Value
    
    Set xlApp = CreateObject("Excel.Application")
            
    Set xlMain = xlApp.Workbooks.Open(MainPath)
    Set xlMainS = xlMain.Sheets("DATA")
    
    ' Value in this worksheet table
    Set targetWb = ThisWorkbook
    Set targetWs = targetWb.Sheets("Hold_Cutting")
    Set tbl = targetWs.ListObjects("Payment_list")
    
    lastRow = tbl.ListColumns(2).DataBodyRange.Cells(tbl.ListRows.Count, 1).End(xlUp).Row - tbl.HeaderRowRange.Row
    For i = 1 To ThisWorkbook.Sheets("Hold_Cutting").Range("AK18").Value
    
        ' Check CountPay before run update
        If tbl.ListColumns(29).DataBodyRange.Cells(i, 1).Value = vbNullString Then
            
            searchText = tbl.ListColumns(2).DataBodyRange.Cells(i, 1).Value
            
            Dim success As Boolean
            Dim newFilePath As String
            
            T_code = LogicTemplate(tbl.ListColumns(10).DataBodyRange.Cells(i, 1).Value)
            foundFile = FindExcelFile(folderPath, searchText)
        
            If foundFile <> vbNullString Then
                'MsgBox "File found: " & foundFile
                
                'response = MsgBox("Do you want to update values?", vbQuestion + vbYesNo, "Update Confirmation")
            
                'If response = vbYes Then
            
                    'Attempt to open the file with the provided password 4142300652
                    Set xlWb = xlApp.Workbooks.Open(foundFile)
                    Set xlWs = xlWb.Sheets("Transaction")
                    
                    If xlWb.Sheets("Payment Term History").Cells(1, "A").Value <> ThisWorkbook.Sheets("Hold_Cutting").Range("ak13").Value Then
                        xlWb.Sheets("Payment Term History").Cells(1, "A").Value = ThisWorkbook.Sheets("Hold_Cutting").Range("ak13").Value
                    End If
    
                    ' Find the latest blank row in xlWs
                    lastRowTB = 2 + xlWb.Sheets("Transaction").Range("Y2").Value
                    
                    If lastRowTB > 30 Then
                        MsgBox "Expanding " & xlWb.Sheets("Payment Term History").Range("I3").Value & "/ 38 rows Left"
                        Dim insert_count As Integer
                        insert_count = 8
                        Do
                            insert_count = insert_count + 1
                            xlWb.Sheets("Payment Term History").Range("A" & insert_count & ":EC" & insert_count).Copy
                        Loop While xlWb.Sheets("Payment Term History").Range("W" & insert_count + 1).Value <> vbNullString
                        xlWb.Sheets("Payment Term History").Cells(insert_count + 1, 1).EntireRow.Insert Shift:=xlDown
                        
                    End If
                    
                    If xlWs.Cells(lastRowTB - 1, "N").Value = tbl.ListColumns(13).DataBodyRange.Cells(i, 1).Value Then
                        response_next = MsgBox("This may be a dupplicate Transaction, still want to continue?", vbQuestion + vbYesNo, "Update Confirmation")
                        If response_next = vbNo Then
                            xlApp.Quit
                            Set xlApp = Nothing
                        End If
                    End If
                    
                    xlWs.Cells(lastRowTB, "A").Value = lastRowTB - 1
                    
                    'lead zero fix
                    If Left(xlWb.Sheets("Payment Term History").Cells(4, 2).Value, 1) = "0" Then
                        xlWs.Cells(lastRowTB, "B").NumberFormat = "@"
                        xlWs.Cells(lastRowTB, "B").Value = "'" & xlWb.Sheets("Payment Term History").Cells(4, 2).Value
                    Else
                        xlWs.Cells(lastRowTB, "B").Value = tbl.ListColumns(2).DataBodyRange.Cells(i, 1).Value
                    End If
                    
                    xlWs.Cells(lastRowTB, "C").Value = tbl.ListColumns(3).DataBodyRange.Cells(i, 1).Value
                    xlWs.Cells(lastRowTB, "D").Value = tbl.ListColumns(5).DataBodyRange.Cells(i, 1).Value
                    xlWs.Cells(lastRowTB, "E").Value = tbl.ListColumns(6).DataBodyRange.Cells(i, 1).Value
                    xlWs.Cells(lastRowTB, "F").Value = tbl.ListColumns(7).DataBodyRange.Cells(i, 1).Value
                    xlWs.Cells(lastRowTB, "G").Value = tbl.ListColumns(8).DataBodyRange.Cells(i, 1).Value
                    xlWs.Cells(lastRowTB, "J").Value = tbl.ListColumns(9).DataBodyRange.Cells(i, 1).Value
                    
                    xlWs.Cells(lastRowTB, "M").Value = CDate(ThisWorkbook.Sheets("Hold_Cutting").Range("AK11").Value)
                    xlWs.Cells(lastRowTB, "N").Value = tbl.ListColumns(13).DataBodyRange.Cells(i, 1).Value
                    
                    'Logic Payment Function
                    If T_code.Result <> vbNullString Then
                        case_select = tbl.ListColumns(10).DataBodyRange.Cells(i, 1).Value
                        
                        If case_select Like "sold_*" And InStr(1, tbl.ListColumns(8).DataBodyRange.Cells(i, 1).Value, "cm") > 0 Then
                            xlWs.Cells(lastRowTB, "L").Value = "pm"
                            Select Case case_select
                                Case "sold_d"
                                    xlWs.Cells(lastRowTB, "P").Value = (tbl.ListColumns(11).DataBodyRange.Cells(i, 1).Value * 100 / 107) - tbl.ListColumns(12).DataBodyRange.Cells(i, 1).Value
                                    xlWs.Cells(lastRowTB, "T").Value = "Direct"
                                Case "sold_a"
                                    xlWs.Cells(lastRowTB, "P").Value = tbl.ListColumns(11).DataBodyRange.Cells(i, 1).Value - tbl.ListColumns(12).DataBodyRange.Cells(i, 1).Value
                                    xlWs.Cells(lastRowTB, "T").Value = "Auction"
                            End Select
    
                        Else
                            Select Case True
                                Case case_select = "pm"
                                    xlWs.Cells(lastRowTB, T_code.Result).Value = tbl.ListColumns(11).DataBodyRange.Cells(i, 1).Value
                                    xlWs.Cells(lastRowTB, "L").Value = "pm"
                                Case case_select Like "*_d"
                                    If Left(case_select, Len(case_select) - 2) = "pm" Then
                                        xlWs.Cells(lastRowTB, T_code.Result).Value = (tbl.ListColumns(11).DataBodyRange.Cells(i, 1).Value * 100 / 107) - tbl.ListColumns(12).DataBodyRange.Cells(i, 1).Value
                                    Else
                                        xlWs.Cells(lastRowTB, T_code.Result).Value = tbl.ListColumns(11).DataBodyRange.Cells(i, 1).Value
                                        xlWs.Cells(lastRowTB, "S").Value = tbl.ListColumns(12).DataBodyRange.Cells(i, 1).Value
                                    End If
                                    xlWs.Cells(lastRowTB, "T").Value = "Direct"
                                    xlWs.Cells(lastRowTB, "L").Value = Left(case_select, Len(case_select) - 2)
                                    
                                Case case_select Like "*_a"
                                    If Left(case_select, Len(case_select) - 2) = "pm" Then
                                        xlWs.Cells(lastRowTB, T_code.Result).Value = tbl.ListColumns(11).DataBodyRange.Cells(i, 1).Value - tbl.ListColumns(12).DataBodyRange.Cells(i, 1).Value
                                    Else
                                        xlWs.Cells(lastRowTB, T_code.Result).Value = (tbl.ListColumns(11).DataBodyRange.Cells(i, 1).Value * 107 / 100)
                                        xlWs.Cells(lastRowTB, "S").Value = tbl.ListColumns(12).DataBodyRange.Cells(i, 1).Value
                                    End If
                                    xlWs.Cells(lastRowTB, "T").Value = "Auction"
                                    xlWs.Cells(lastRowTB, "L").Value = Left(case_select, Len(case_select) - 2)
                                    
                                Case Else
                                    xlWs.Cells(lastRowTB, T_code.Result).Value = tbl.ListColumns(11).DataBodyRange.Cells(i, 1).Value
                                    xlWs.Cells(lastRowTB, "L").Value = tbl.ListColumns(10).DataBodyRange.Cells(i, 1).Value
                            End Select
                        End If
                    End If
                    
                    xlWb.Close SaveChanges:=True
                    'xlMain.Close SaveChanges:=True
                    'xlApp.Quit
                    'Set xlApp = Nothing
            
                    err_code = RevealPayment(foundFile, MainPath, T_code.Result, searchText, xlApp)
                    
                    Select Case err_code.Result
                        Case "Miss"
                            tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Database Missing"
                        Case "Value"
                            tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Database Error"
                        Case "Path"
                            tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Path error"
                        Case "Error"
                            tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Checking Request"
                        Case "Eq_pm"
                            tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Payment Error"
                        Case Else
                            tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Update Payment"
                    End Select
            
                    
            
            Else
                    'MsgBox "File not found."
        
                    'response = MsgBox("Do you want to Create new profile PAMID: " & searchText & " ?", vbQuestion + vbYesNo, "Update Confirmation")
                
                    'If response = vbYes Then
                            foundFile = folderPath & TempName
                            newFilePath = folderPath & "StatementCard_" & searchText
                            
                            Dim resultCollection As Collection
                            Set resultCollection = UpdateExcelFile(foundFile, searchText, i, T_code.Result, newFilePath, MainPath, xlApp)
                            
                            Dim result1 As Boolean
                            Dim result2 As String
                            result1 = resultCollection(1)
                            result2 = resultCollection(2)
                            
                            success = result1
                            err_code = RevealPayment(result2, MainPath, T_code.Result, searchText, xlApp)
                            
                            If rollback = True Or success = False Then
                                tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Path error"
                                Exit Sub
                            End If
                        
                            Select Case err_code.Result
                                Case "Miss"
                                    tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Database Missing"
                                Case "Value"
                                    tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Database Error"
                                Case "Path"
                                    tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Path error"
                                Case "Error"
                                    tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Checking Request"
                                Case "Eq_pm"
                                    tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Payment Error"
                                Case Else
                                    tbl.ListColumns(34).DataBodyRange.Cells(i, 1).Value = "Create new card"
                            End Select
                    
            End If
            
            tbl.ListColumns(4).DataBodyRange.Cells(i, 1).Interior.Color = RGB(150, 200, 0)

            TextBox4.Text = "Updating : " & i & " / " & ThisWorkbook.Sheets("Hold_Cutting").Range("AK18").Value
            
        End If
    Next i
    
    xlApp.Quit
    Set xlApp = Nothing
    
    MsgBox "Card Done"
    
End Sub
Private Sub CommandButton3_Click()

    MsgBox "Select Data Template File"
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
            ThisWorkbook.Sheets("Hold_Cutting").Range("AM13").Value = Left(selectedFile, InStrRev(selectedFile, Application.PathSeparator)) & Mid(selectedFile, InStrRev(selectedFile, Application.PathSeparator) + 1)
            'Right(selectedFile, Len(selectedFile) - InStrRev(selectedFile, Application.PathSeparator))
        ' The user canceled the file picker
        End If
    End With
    
    MsgBox "Select Data Cutoff File"
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
            TextBox2.Value = Left(selectedFile, InStrRev(selectedFile, Application.PathSeparator)) & Mid(selectedFile, InStrRev(selectedFile, Application.PathSeparator) + 1)
        ' The user canceled the file picker
        End If
    End With
    
    MsgBox "Select Data Status File"
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
            ThisWorkbook.Sheets("Hold_Cutting").Range("AN13").Value = Left(selectedFile, InStrRev(selectedFile, Application.PathSeparator)) & Mid(selectedFile, InStrRev(selectedFile, Application.PathSeparator) + 1)
        ' The user canceled the file picker
        End If
    End With
    
    MsgBox "Select Daily Report File"
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
            TextBox3.Value = Left(selectedFile, InStrRev(selectedFile, Application.PathSeparator)) & Mid(selectedFile, InStrRev(selectedFile, Application.PathSeparator) + 1)
        
        Else
            Exit Sub ' The user canceled the file picker
        End If
    End With

End Sub

Private Sub CommandButton4_Click()
    
    On Error Resume Next
    If test_mode Then
        folderPath = "C:\Pam\Tools\Cut-off\Testing env\"
        MainPath = "C:\Pam\Tools\Cut-off\Testing env\_Data_for_Cut_off_Table.xlsx"
        MacroPath = "C:\Pam\Tools\Cut-off\Cut_off\Cut-Off-Database_marco.xlsm"
    Else
        MainPath = ThisWorkbook.Sheets("Hold_Cutting").Range("AM13").Value 'TextBox2.Value
        MacroPath = TextBox3.Value
        folderPath = Left(MainPath, InStrRev(MainPath, "\"))
    End If
    On Error GoTo 0
    
    'Set xlApp = CreateObject("Excel.Application")
    
    ' Value in this worksheet table
    Set targetWb = ThisWorkbook
    Set targetWs = targetWb.Sheets("Hold_Cutting")
    Set tbl = targetWs.ListObjects("Payment_list")
    
    ' Macro path
    Set xlMacro = Workbooks.Open(MacroPath, password:=ThisWorkbook.Sheets("Hold_Cutting").Range("AK15").Value)
    Set xlMacroS = xlMacro.Sheets("Daily_Report")
    Set tbMac = xlMacroS.ListObjects("Data_daily")
    
    lr = ThisWorkbook.Sheets("Hold_Cutting").Range("AK18").Value 'tbl.ListRows.Count ' Get the number of rows in the table
    mr = tbMac.ListRows.Count ' Get the number of rows in the macro table
    Application.ScreenUpdating = True
    ' Loop through all rows in the target sheet's table
    For i = 1 To lr
        Dim searchValue As String
        searchValue = tbl.ListColumns("Code").DataBodyRange(i).Value
        ' Find the matching row in the macro table (assuming you have a unique identifier column, e.g., "O" column)
        Set foundCell = tbMac.ListColumns(2).DataBodyRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            l = foundCell.Row ' Row in the macro table where the match was found
            copy_i = i + 5
            
            If (targetWs.Cells(copy_i, "AC") <> vbNullString) And (targetWs.Cells(copy_i, "A") = vbNullString) Then

                'status mode
                targetWs.Range("G" & copy_i & ":H" & copy_i).Copy
                xlMacroS.Cells(l, 11).PasteSpecial xlPasteValues
                'count payment
                'targetWs.Cells(copy_i, 28).Copy
                'xlMacroS.Cells(l, 27).PasteSpecial xlPasteValues
                'eff date
                xlMacroS.Cells(l, 29).Value = targetWs.Cells(copy_i, 13).Value
                
                If (targetWs.Cells(copy_i, "J") = "vd") Then
                    If xlMacroS.Cells(l, 30) < targetWs.Cells(copy_i, 30).Value Then
                        'Balance and vat
                        targetWs.Range("O" & copy_i & ":AC" & copy_i).Copy
                        xlMacroS.Cells(l, 13).PasteSpecial xlPasteValues
                        'total payment vat
                        targetWs.Range("AD" & copy_i & ":AG" & copy_i).Copy
                        xlMacroS.Cells(l, 30).PasteSpecial xlPasteValues
                    Else
                        'Balance and vat
                        targetWs.Range("O" & copy_i & ":AB" & copy_i).Copy
                        xlMacroS.Cells(l, 13).PasteSpecial xlPasteValues
                    End If
 
                ElseIf (targetWs.Cells(copy_i, "J") = "dis") Then
                    If xlMacroS.Cells(l, 30) < targetWs.Cells(copy_i, 30).Value Then
                        'Balance and vat
                        targetWs.Range("O" & copy_i & ":AC" & copy_i).Copy
                        xlMacroS.Cells(l, 13).PasteSpecial xlPasteValues
                        'total payment vat
                        targetWs.Range("AD" & copy_i & ":AG" & copy_i).Copy
                        xlMacroS.Cells(l, 30).PasteSpecial xlPasteValues
                    Else
                        'Balance and vat
                        targetWs.Range("O" & copy_i & ":AB" & copy_i).Copy
                        xlMacroS.Cells(l, 13).PasteSpecial xlPasteValues
                    End If
                        
                Else
                    'Balance and vat
                    targetWs.Range("O" & copy_i & ":AC" & copy_i).Copy
                    xlMacroS.Cells(l, 13).PasteSpecial xlPasteValues
                    
                    'Amount
                    targetWs.Cells(copy_i, 11).Copy
                    xlMacroS.Cells(l, 28).PasteSpecial xlPasteValues

                    'xlMacroS.Cells(l, 27).Value = targetWs.Cells(copy_i, 24).Value
                    'total payment vat
                    targetWs.Range("AD" & copy_i & ":AG" & copy_i).Copy
                    xlMacroS.Cells(l, 30).PasteSpecial xlPasteValues
                End If
                'DocID
                'targetWs.Cells(copy_i, 9).Copy
                'xlMacroS.Cells(l, 37).PasteSpecial xlPasteValues
                
                'payment_discount
                temp_dis = Payment_code(targetWs.Cells(copy_i, 10).Value)
                If xlMacroS.Cells(l, 34) <> temp_dis.Result Then
                    Select Case targetWs.Cells(copy_i, 10)
                        Case "dis"
                            If targetWs.Cells(copy_i, 11).Value <> vbNullString Then
                                xlMacroS.Cells(l, 34) = "Discount"
                                xlMacroS.Cells(l, 35).Value = targetWs.Cells(copy_i, 11).Value
                                xlMacroS.Cells(l, 36).Value = targetWs.Cells(copy_i, 11).Value * (targetWs.Cells(copy_i, 25).Value / (1 - targetWs.Cells(copy_i, 11).Value))
                                xlMacroS.Cells(l, 37).Value = targetWs.Cells(copy_i, 11).Value * (targetWs.Cells(copy_i, 27).Value / (1 - targetWs.Cells(copy_i, 11).Value))
                            End If
                        Case "add"
                            xlMacroS.Cells(l, 34) = "Normal"
                        Case Else
                    End Select
                End If
                    
                targetWs.Cells(copy_i, "A").Value = "Done"
                targetWs.Cells(copy_i, "A").Interior.Color = RGB(255, 255, 0)
            End If
            'Dim userResponse As VbMsgBoxResult
            'userResponse = MsgBox("Do you want to continue updating values?", vbYesNo)
    
            'If userResponse = vbNo Then
                'Exit Sub
                'xlApp.Quit
            'End If
            TextBox4.Text = "Updating : " & i & " / " & lr
        End If
    Next i
    
    'Dim currentDate As Date
    xlMacro.Sheets("OA").Cells(10, "F").Value = Date
    
    ' Save the workbook with password protection
    'SaveWorkbookWithPassword "DailyReport_" & ThisWorkbook.Sheets("Hold_Cutting").Range("AK11").Value & ".xlsx", xlMacro, ThisWorkbook.Sheets("Hold_Cutting").Range("AK15").Value
    SaveWorkbookWithDialog "DailyReport_" & ThisWorkbook.Sheets("Hold_Cutting").Range("AK11").Value & ".xlsx", xlMacro, ThisWorkbook.Sheets("Hold_Cutting").Range("AK15").Value
    
    MsgBox ("Update Report Done")
   
    xlMacro.Close SaveChanges:=False
    
    ThisWorkbook.Sheets("Hold_Cutting").Range("AL13").Value = Left(MacroPath, InStrRev(MacroPath, "\")) & "DailyReport_" & ThisWorkbook.Sheets("Hold_Cutting").Range("AK11").Value & ".xlsx"
    'xlApp.Quit
    'Set xlApp = Nothing
    'Application.ScreenUpdating = True
    Call SendEmail
    
End Sub
Sub SaveWorkbookWithPassword(ByVal filename As String, ByVal wb As Workbook, ByVal password As String)
    ' Save workbook with password protection
    wb.SaveAs filename:=filename, password:=password
End Sub
Private Sub CommandButton5_Click()

    Dim iter_rows As Integer
    Dim port_fix As String
    
    On Error Resume Next
    If test_mode Then
        MainPath = " "
        MacroPath = " "
    Else
        MainPath = TextBox2.Value
        MacroPath = ThisWorkbook.Sheets("Hold_Cutting").Range("AM13").Value
    End If
    On Error GoTo 0
    
    folderPath = Left(MacroPath, InStrRev(MacroPath, "\"))
    
    Set xlApp = CreateObject("Excel.Application")
            
    Set xlMain = xlApp.Workbooks.Open(MainPath)
    Set xlMainS = xlMain.Sheets("DATA")
    
    ' Value in this worksheet table
    'Set targetWb = ThisWorkbook
    'Set targetWs = targetWb.Sheets("UpdateName")
    'Set tbl = targetWs.ListObjects("UpdateName")
    
    ' Count the total number of files in the folder
    filename = Dir(folderPath & "\StatementCard_*.xlsx")
    Do While filename <> ""
        totalFiles = totalFiles + 1
        filename = Dir
    Loop
    
    ' Loop through each file in the folder
    filename = Dir(folderPath & "\StatementCard_*.xlsx")
    Do While filename <> ""
        ' Open the Excel file
        Set xlWb = xlApp.Workbooks.Open(folderPath & "\" & filename)
        Set xlWs = xlWb.Sheets("Payment Term History")
        
        ' Update the value of a specific cell / Formula
        xlWs.Cells(1, "A").Value = MainPath
        xlWs.Cells(3, "F").Formula = "=COUNTIF(INDIRECT(""L7:L"" & I3+6), ""pm"")"
        xlWs.Cells(3, "I").Formula = "=Transaction!$Y$2+1"
        
        'Fix Port name
        iter_rows = xlWb.Sheets("Transaction").Cells(2, "Y").Value + 1
        port_fix = xlWs.Cells(7, "D").Value
        For i = 2 To iter_rows
            xlWb.Sheets("Transaction").Cells(i, "D").Value = port_fix
        Next i
        
        ' Get the new filename
        Dim newFileName As String
        newFileName = "StatementCard_" & xlWs.Cells(4, "B").Value & "_" & xlWs.Cells(7, "DC").Value & ".xlsx"
        
        ' Save and close the file
        xlWb.Close SaveChanges:=True
        
        'change file name
        If newFileName <> "" Then
            Name folderPath & "\" & filename As folderPath & "\" & newFileName
        End If
        
        processedFiles = processedFiles + 1
        TextBox4.Text = "Updating Name : " & processedFiles & " / " & totalFiles
        
        ' Get the next file in the folder
        filename = Dir
    Loop
    
    MsgBox "Update Card Path"
    
End Sub

Private Sub CommandButton6_Click()
    On Error Resume Next
    If test_mode Then
        MainPath = " "
        MacroPath = " "
    Else
        MainPath = TextBox2.Value
        MacroPath = ThisWorkbook.Sheets("Hold_Cutting").Range("AM13").Value
    End If
    On Error GoTo 0
    
    folderPath = Left(MacroPath, InStrRev(MacroPath, "\"))
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlMain = xlApp.Workbooks.Open(MainPath)
    Set xlMainS = xlMain.Sheets("DATA")
    
    ' Value in this worksheet table
    Set targetWb = ThisWorkbook
    Set targetWs = targetWb.Sheets("Hold_Cutting")
    Set tbl = targetWs.ListObjects("Payment_list")
    
    ' Count the total number of files in the folder
    filename = Dir(folderPath & "\StatementCard_*.xlsx")
    Do While filename <> ""
        totalFiles = totalFiles + 1
        filename = Dir
    Loop
    
    tbl.DataBodyRange.ClearContents
    tbl.ListColumns(1).DataBodyRange.ClearFormats
    tbl.ListColumns(4).DataBodyRange.ClearFormats
    tbl.ListColumns(10).DataBodyRange.ClearFormats
    
    ' Loop through each file in the folder
    processedFiles = 1
    filename = Dir(folderPath & "\StatementCard_*.xlsx")
    Do While filename <> ""
        ' Open the Excel file
        Set xlWb = xlApp.Workbooks.Open(folderPath & "\" & filename)
        Set xlWs = xlWb.Sheets("Payment Term History")
        
        lastRowTB = 6 + xlWs.Cells(3, "I").Value
        ' Update the value of a specific cell to sheet
        'ID name
        tbl.ListColumns(2).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(4, "B").Value
        tbl.ListColumns(3).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "C").Value
        tbl.ListColumns(4).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(7, "A").Value
        'Status mode Port
        tbl.ListColumns(5).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "D").Value
        tbl.ListColumns(6).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "E").Value
        tbl.ListColumns(7).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "F").Value
        tbl.ListColumns(8).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "G").Value
        'amount eff_date
        tbl.ListColumns(10).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "L").Value
        tbl.ListColumns(12).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "S").Value
        
        If (xlWs.Cells(lastRowTB, "L").Value = "pm") And (xlWs.Cells(lastRowTB, "P").Value > 0) Then
            tbl.ListColumns(11).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "P").Value
        ElseIf xlWs.Cells(lastRowTB, "R").Value > 0 Then
            tbl.ListColumns(11).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "R").Value
        Else
            
        End If
        
        Select Case xlWs.Cells(lastRowTB, "T").Value
            Case "Auction"
                tbl.ListColumns(10).DataBodyRange.Cells(processedFiles, 1).Value = tbl.ListColumns(10).DataBodyRange.Cells(processedFiles, 1).Value & "_a"
            Case "Direct"
                tbl.ListColumns(10).DataBodyRange.Cells(processedFiles, 1).Value = tbl.ListColumns(10).DataBodyRange.Cells(processedFiles, 1).Value & "_d"
            Case Else
        End Select
        
        tbl.ListColumns(13).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "N").Value
        xlWb.Close SaveChanges:=True
        
        foundFile = folderPath & "\" & filename
        T_code = LogicTemplate(tbl.ListColumns(10).DataBodyRange.Cells(processedFiles, 1).Value)
        err_code = RevealPayment(foundFile, MainPath, T_code.Result, tbl.ListColumns(2).DataBodyRange.Cells(processedFiles, 1).Value, xlApp)
                    
        Select Case err_code.Result
            Case "Miss"
                tbl.ListColumns(34).DataBodyRange.Cells(processedFiles, 1).Value = "Database Missing"
            Case "Value"
                tbl.ListColumns(34).DataBodyRange.Cells(processedFiles, 1).Value = "Database Error"
            Case "Path"
                tbl.ListColumns(34).DataBodyRange.Cells(processedFiles, 1).Value = "Path error"
            Case "Error"
                tbl.ListColumns(34).DataBodyRange.Cells(processedFiles, 1).Value = "Checking Request"
            Case "Eq_pm"
                tbl.ListColumns(34).DataBodyRange.Cells(processedFiles, 1).Value = "Payment Error"
            Case Else
                tbl.ListColumns(34).DataBodyRange.Cells(processedFiles, 1).Value = "Update Payment"
        End Select
        
        
        processedFiles = processedFiles + 1
        TextBox4.Text = "Recall : " & processedFiles - 1 & " / " & totalFiles
        
        ' Get the next file in the folder
        filename = Dir
    Loop
    
    MsgBox "Recall Done"
End Sub

Private Sub CommandButton7_Click()
    ' Count the total number of files in the folder
    filename = Dir("C:\Pam_card\summary\file" & "\StatementCard_*.xlsx")
    Do While filename <> ""
        totalFiles = totalFiles + 1
        filename = Dir
    Loop
    
    If totalFiles = 0 Then
        MsgBox ("Bring Statement Card to Pam_card -> summary -> file")
        Exit Sub
    Else
        TextBox4.Text = "Summary total  " & totalFiles & " File"
        Call SummaryCard

    End If
    
End Sub

Private Sub TextBox1_Change()
    ' Update the value in cell when TextBox1 changes
    ThisWorkbook.Sheets("Hold_Cutting").Range("aj13").Value = TextBox1.Value
End Sub
Private Sub TextBox2_Change()
    ThisWorkbook.Sheets("Hold_Cutting").Range("ak13").Value = TextBox2.Value
End Sub
Private Sub TextBox3_Change()
    ThisWorkbook.Sheets("Hold_Cutting").Range("al13").Value = TextBox3.Value
End Sub
Private Sub TextBox1_AfterUpdate()
    ' Check if TextBox1 is now filled
    If TextBox1.Value <> vbNullString Then
        ' Enable buttons or perform actions as needed
        Me.CommandButton2.Enabled = True
        Me.CommandButton4.Enabled = True
    End If
End Sub

Private Sub TextBox2_AfterUpdate()
    ' Check if TextBox2 is now filled
    If TextBox2.Value <> vbNullString Then
        ' Enable buttons or perform actions as needed
        Me.CommandButton2.Enabled = True
        Me.CommandButton4.Enabled = True
    End If
End Sub

Private Sub TextBox3_AfterUpdate()
    ' Check if TextBox3 is now filled
    If TextBox3.Value <> vbNullString Then
        ' Enable buttons or perform actions as needed
        Me.CommandButton2.Enabled = True
        Me.CommandButton4.Enabled = True
    End If
End Sub
Private Sub UserForm_Initialize()
    ' Initialize the TextBox with the value from cell on Sheet1
    TextBox1.Value = ThisWorkbook.Sheets("Hold_Cutting").Range("aj13").Value
    TextBox2.Value = ThisWorkbook.Sheets("Hold_Cutting").Range("ak13").Value
    TextBox3.Value = ThisWorkbook.Sheets("Hold_Cutting").Range("al13").Value
    User_name = ThisWorkbook.Sheets("Hold_Cutting").Range("ak17").Value
    
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    If Not xlApp Is Nothing Then
        wbCount = xlApp.Workbooks.Count
        If wbCount > 1 Then
            ' Close all open workbooks except the current one
            For Each wb In xlApp.Workbooks
                If Not wb.Name = ThisWorkbook.Name Then
                    wb.Close False
                End If
            Next wb
        End If
        Set xlApp = Nothing
    End If
    
    Set tbl = ThisWorkbook.Sheets("Hold_Cutting").ListObjects("Payment_list")
    'RowCount = ThisWorkbook.Sheets("Hold_Cutting").ListColumns("Name").DataBodyRange.Cells(ThisWorkbook.Sheets("Hold_Cutting").ListColumns("Name").DataBodyRange.Rows.Count).Row
    ListBox1.MultiSelect = fmMultiSelectMulti
    
    Dim colIndex As Integer
    colIndex = tbl.ListColumns("Code").Index ' Replace "ColumnName" with the actual column name

    Dim cell As Range
    For Each cell In tbl.ListColumns(colIndex).DataBodyRange
        ' Skip cells with blank values
        If Not IsEmpty(cell.Value) Then
            ListBox1.AddItem cell.Value
        End If
    Next cell
    
        ' Check if TextBox1 is blank
    If TextBox1.Value = vbNullString Then
        MsgBox "Please add payment file.", vbExclamation
        ' Disable AnotherButton
        Me.CommandButton2.Enabled = False
        Me.CommandButton4.Enabled = False
    End If
    
    If TextBox2.Value = vbNullString Then
        MsgBox "Please add Data file", vbExclamation
        ' Disable AnotherButton
        Me.CommandButton2.Enabled = False
        Me.CommandButton4.Enabled = False
    End If
    
    If TextBox3.Value = vbNullString Then
        MsgBox "Please add Report file.", vbExclamation
        ' Disable AnotherButton
        Me.CommandButton2.Enabled = False
        Me.CommandButton4.Enabled = False
    End If
    'MsgBox tbl.ListColumns(2).Range(tbl.ListColumns(2).Range.Rows.Count, 1).End(xlUp).Row
    'MsgBox tbl.ListColumns(2).DataBodyRange.Cells(tbl.ListRows.Count, 1).End(xlUp).Row
    'MsgBox tbl.ListColumns(2).DataBodyRange.Cells(tbl.ListRows.Count, 1).End(xlUp).Row - tbl.HeaderRowRange.Row + 1
    
    'Dir(selectedFile, vbDirectory)
    
End Sub
'Update Function Code Here
Function LogicTemplate(C_Code As String) As ReturnValue
    Dim Result As ReturnValue
    
    Select Case True
        Case C_Code Like "pm*"
            Result.Result = "P"
        Case C_Code Like "sold*"
            Result.Result = "R"
        Case C_Code = "esue", C_Code = "sue"
            Result.Result = "Q"
        Case C_Code = "dis"
            Result.Result = "U"
        Case C_Code = "vd"
            Result.Result = "Q"
        Case C_Code = ""
            Result.Result = vbNullString
            ' Handle other cases if needed
    End Select
    
    LogicTemplate = Result
End Function
'Update Payment Code Here
Function Payment_code(D_Code As String) As ReturnValue
    Dim result1 As ReturnValue
    
    Select Case D_Code
        Case "dis"
            result1.Result = "Discount"
        Case Else
            result1.Result = "Normal"
    End Select
    
    Payment_code = result1
End Function

Function UpdateExcelFile(foundFile As String, pam_code As Variant, Nplace As Variant, p_code As String, newFilePath As String, MainPath As String, xlApp As Object) As Collection
    Dim results As New Collection
    Dim xlWb As Object
    Dim xlWs As Object
    Dim xlMain As Object
    Dim xlMainS As Object
    Dim lastRow As Long
    
    Set tbl = ThisWorkbook.Sheets("Hold_Cutting").ListObjects("Payment_list")
    
    'On Error Resume Next
    'Set xlApp = CreateObject("Excel.Application")
    'On Error GoTo 0
    
    'Set xlMain = xlApp.Workbooks.Open(MainPath)
    'Set xlMainS = xlMain.Sheets("DATA")
    
    Set xlWb = xlApp.Workbooks.Open(foundFile)
    Set xlWs = xlWb.Sheets("Payment Term History")
    Set xlWt = xlWb.Sheets("Transaction")
    
    'If xlWs.Cells(1, "A").Value <> ThisWorkbook.Sheets("Hold_Cutting").Range("ah13").Value Then
        'xlWs.Cells(1, "A").Value = ThisWorkbook.Sheets("Hold_Cutting").Range("ah13").Value
    'End If
    
    ' Find the latest blank row in xlWs
    xlWs.Cells(1, 1).Value = ThisWorkbook.Sheets("Hold_Cutting").Range("AK13").Value
    xlWs.Cells(4, 2).Value = pam_code
    xlWs.Cells(7, "M").Value = CDate(ThisWorkbook.Sheets("Hold_Cutting").Range("AK11").Value)
    
    'Error checking to exit loop
    If xlWs.Cells(1, "A").Value <> TextBox2.Value Then
        'fix path
        xlWs.Cells(1, "A").Value = TextBox2.Value
    End If
    
    xlWt.Cells(2, "A").Value = 1
    xlWt.Cells(2, "B").Value = tbl.ListColumns(2).DataBodyRange.Cells(Nplace, 1).Value
    xlWt.Cells(2, "C").Value = tbl.ListColumns(3).DataBodyRange.Cells(Nplace, 1).Value
    xlWt.Cells(2, "D").Value = tbl.ListColumns(5).DataBodyRange.Cells(Nplace, 1).Value
    xlWt.Cells(2, "E").Value = tbl.ListColumns(6).DataBodyRange.Cells(Nplace, 1).Value
    xlWt.Cells(2, "F").Value = tbl.ListColumns(7).DataBodyRange.Cells(Nplace, 1).Value
    xlWt.Cells(2, "G").Value = tbl.ListColumns(8).DataBodyRange.Cells(Nplace, 1).Value
    xlWt.Cells(2, "J").Value = tbl.ListColumns(9).DataBodyRange.Cells(Nplace, 1).Value
    
    xlWt.Cells(2, "M").Value = CDate(ThisWorkbook.Sheets("Hold_Cutting").Range("AK11").Value)
    xlWt.Cells(2, "N").Value = tbl.ListColumns(13).DataBodyRange.Cells(Nplace, 1).Value
    xlWt.Cells(2, "S").Value = tbl.ListColumns(12).DataBodyRange.Cells(Nplace, 1).Value
    
    'Logic Payment Function
    If p_code <> vbNullString Then
        case_select = tbl.ListColumns(10).DataBodyRange.Cells(Nplace, 1).Value
                        
        If case_select Like "sold_*" And InStr(1, tbl.ListColumns(8).DataBodyRange.Cells(Nplace, 1).Value, "cm") > 0 Then
            xlWt.Cells(2, "L").Value = "pm"
            Select Case case_select
                Case "sold_d"
                    xlWt.Cells(2, "P").Value = (tbl.ListColumns(11).DataBodyRange.Cells(Nplace, 1).Value * 100 / 107) - tbl.ListColumns(12).DataBodyRange.Cells(Nplace, 1).Value
                    xlWt.Cells(2, "T").Value = "Direct"
                Case "sold_a"
                    xlWt.Cells(2, "P").Value = tbl.ListColumns(11).DataBodyRange.Cells(Nplace, 1).Value - tbl.ListColumns(12).DataBodyRange.Cells(Nplace, 1).Value
                    xlWt.Cells(2, "T").Value = "Auction"
            End Select
    
        Else
            Select Case True
                Case case_select = "pm"
                    xlWt.Cells(2, p_code).Value = tbl.ListColumns(11).DataBodyRange.Cells(Nplace, 1).Value
                    xlWt.Cells(2, "L").Value = "pm"
                Case case_select Like "*_d"
                    If Left(case_select, Len(case_select) - 2) = "pm" Then
                        xlWt.Cells(2, p_code).Value = (tbl.ListColumns(11).DataBodyRange.Cells(Nplace, 1).Value * 100 / 107) - tbl.ListColumns(12).DataBodyRange.Cells(Nplace, 1).Value
                    Else
                        xlWt.Cells(2, p_code).Value = tbl.ListColumns(11).DataBodyRange.Cells(Nplace, 1).Value
                        xlWt.Cells(2, "S").Value = tbl.ListColumns(12).DataBodyRange.Cells(Nplace, 1).Value
                    End If
                    xlWt.Cells(2, "T").Value = "Direct"
                    xlWt.Cells(2, "L").Value = Left(case_select, Len(case_select) - 2)
                                    
                Case case_select Like "*_a"
                    If Left(case_select, Len(case_select) - 2) = "pm" Then
                        xlWt.Cells(2, p_code).Value = tbl.ListColumns(11).DataBodyRange.Cells(Nplace, 1).Value - tbl.ListColumns(12).DataBodyRange.Cells(Nplace, 1).Value
                    Else
                        xlWt.Cells(2, p_code).Value = (tbl.ListColumns(11).DataBodyRange.Cells(Nplace, 1).Value * 107 / 100)
                        xlWt.Cells(2, "S").Value = tbl.ListColumns(12).DataBodyRange.Cells(Nplace, 1).Value
                    End If
                    xlWt.Cells(2, "T").Value = "Auction"
                    xlWt.Cells(2, "L").Value = Left(case_select, Len(case_select) - 2)
                                    
                Case Else
                    xlWt.Cells(2, p_code).Value = tbl.ListColumns(11).DataBodyRange.Cells(Nplace, 1).Value
                    xlWt.Cells(2, "L").Value = tbl.ListColumns(10).DataBodyRange.Cells(Nplace, 1).Value
            End Select
        End If
    End If
    
    'newFilePath & "_" & xlWs.Cells(7, "DC").Value & ".xlsx"
    results.Add True ' Result 1 as Boolean
    results.Add newFilePath & "_" & xlWs.Cells(7, "DC").Value & ".xlsx" ' Result 2 as String"
    
    ' Save the workbook as the new file path
    xlWb.SaveAs newFilePath & "_" & xlWs.Cells(7, "DC").Value & ".xlsx"
    xlWb.Close SaveChanges:=True
    'xlMain.Close SaveChanges:=False
    
    Set UpdateExcelFile = results
End Function
Private Function FindExcelFile(ByVal folderPath As String, ByVal searchText As String) As String
    Dim filename As String
    Dim filePath As String

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Look for files with the search text in the folder
    filename = Dir(folderPath & "*" & searchText & "*.xlsx")

    ' Loop through each .xlsx file in the folder
    Do While filename <> ""
    
        ' Construct the full file path
        filePath = folderPath & filename
        
        ' Split the filename on the delimiter (e.g., underscore or hyphen)
        Dim parts As Variant
        parts = Split(filename, "_") ' Change "_" to "-" or your specific delimiter as necessary

        Dim part As Variant
        For Each part In parts
            ' Check if the part matches the searchText, also remove file extension before comparison
            If StrComp(Replace(part, ".xlsx", ""), searchText, vbTextCompare) = 0 Then
                FindExcelFile = folderPath & filename
                Exit Function
            End If
        Next part

        ' Get the next file in the directory.
        filename = Dir()
    Loop

    ' If no file with the search text is found, return an empty string
    FindExcelFile = ""
End Function
Private Function RevealPayment(newFilePath As String, MainPath As String, Nplace As Variant, core_id As String, xlApp As Object) As ReturnValue
    Dim return_code As ReturnValue
    Dim lastStatRow As Variant
    ' Value in this worksheet table
    Set targetWb = ThisWorkbook
    Set targetWs = targetWb.Sheets("Hold_Cutting")
    Set tbl = targetWs.ListObjects("Payment_list")
        
    'On Error Resume Next
    'Set xlApp = CreateObject("Excel.Application")
    'On Error GoTo 0
        
    Set xlMain = xlApp.Workbooks.Open(MainPath)
    Set xlMainS = xlMain.Sheets("DATA")
    
    Set xlWb = xlApp.Workbooks.Open(newFilePath)
    Set xlWs = xlWb.Sheets("Payment Term History")
    
    If Not targetWs.UsedRange.Find(What:=searchText, LookAt:=xlWhole, MatchCase:=True) Is Nothing Then
                
        lr = targetWs.Range("AK18").Value
        For l = 6 To lr + 6
            If targetWs.Cells(l, "B").Value = core_id Then
                If targetWs.Cells(l, "O").Value = vbNullString Then
                    
                    lastRow = 6 + xlWs.Cells(3, "I").Value
                    
                    Dim cellValue As Variant
                    cellValue = targetWs.Cells(l, 27).Value

                    If IsError(cellValue) Then
                        ' Additional check for specific errors if needed
                        If cellValue = CVErr(xlErrNA) Then
                            return_code.Result = "Miss"
                            RevealPayment = return_code
                            Exit Function
                        ElseIf cellValue = CVErr(xlErrValue) Then
                            return_code.Result = "Value"
                            RevealPayment = return_code
                            Exit Function
                        ElseIf cellValue = CVErr(xlErrRef) Then
                            return_code.Result = "Path"
                            RevealPayment = return_code
                            Exit Function
                        Else
                            return_code.Result = "Error"
                            RevealPayment = return_code
                            Exit Function
                        End If
                    End If
                    
                    If targetWs.Cells(l, 13).Value < xlWs.Cells(lastRow, "N").Value Then
                        targetWs.Cells(l, 13).Value = xlWs.Cells(lastRow, "N").Value
                    End If
        
                    targetWs.Cells(l, 15).Value = xlWs.Cells(lastRow, "AZ").Value
                    targetWs.Cells(l, 16).Value = xlWs.Cells(lastRow, "BC").Value
                    targetWs.Cells(l, 17).Value = xlWs.Cells(lastRow, "BF").Value
                    targetWs.Cells(l, 18).Value = xlWs.Cells(lastRow, "BI").Value
                    targetWs.Cells(l, 19).Value = xlWs.Cells(lastRow, "BM").Value
                    targetWs.Cells(l, 20).Value = xlWs.Cells(lastRow, "BP").Value
                    targetWs.Cells(l, 21).Value = xlWs.Cells(lastRow, "BS").Value
                    targetWs.Cells(l, 22).Value = xlWs.Cells(lastRow, "BV").Value
                    targetWs.Cells(l, 23).Value = xlWs.Cells(lastRow, "BY").Value
                    targetWs.Cells(l, 24).Value = xlWs.Cells(lastRow, "CB").Value
                    targetWs.Cells(l, 25).Value = xlWs.Cells(lastRow, "CC").Value
                    targetWs.Cells(l, 26).Value = xlWs.Cells(lastRow, "CD").Value
                    targetWs.Cells(l, 27).Value = xlWs.Cells(lastRow, "CE").Value 'check error return code hold
                    targetWs.Cells(l, 28).Value = xlWs.Cells(lastRow, "CH").Value
                    
                    If targetWs.Cells(l, 10).Value Like "sold_*" Then
                        targetWs.Cells(l, 7).Value = targetWs.Cells(l, 7).Value & "_loss_after"
                        Set xStat = xlApp.Workbooks.Open(ThisWorkbook.Sheets("Hold_Cutting").Range("an13").Value)
                        Set xStatS = xStat.Sheets("payment_mode_status")
                        
                        lastStatRow = xStatS.Cells(Rows.Count, 1).End(xlUp).Row + 1
                        
                        xStatS.Cells(lastStatRow, 1).Value = targetWs.Cells(l, 2).Value
                        xStatS.Cells(lastStatRow, 3).Value = targetWs.Cells(l, 7).Value
                        xStatS.Cells(lastStatRow, 4).Value = targetWs.Cells(l, 8).Value
                        xStatS.Cells(lastStatRow, 5).Value = targetWs.Cells(l, 13).Value
                        
                        xStat.Close SaveChanges:=True
                    End If
                    
                    If targetWs.Cells(l, 27).Value < 20 Then
                        targetWs.Cells(l, 7).Value = targetWs.Cells(l, 7).Value & "_closed"
                        Set xStat = xlApp.Workbooks.Open(ThisWorkbook.Sheets("Hold_Cutting").Range("an13").Value)
                        Set xStatS = xStat.Sheets("payment_mode_status")
                        
                        lastStatRow = xStatS.Cells(Rows.Count, 1).End(xlUp).Row + 1
                        
                        xStatS.Cells(lastStatRow, 1).Value = targetWs.Cells(l, 2).Value
                        xStatS.Cells(lastStatRow, 3).Value = targetWs.Cells(l, 7).Value
                        xStatS.Cells(lastStatRow, 4).Value = targetWs.Cells(l, 8).Value
                        xStatS.Cells(lastStatRow, 5).Value = targetWs.Cells(l, 13).Value
                        
                        xStat.Close SaveChanges:=True
                    End If

                    targetWs.Cells(l, 29).Value = xlWs.Cells(3, "F").Value
                    targetWs.Cells(l, 30).Value = xlWs.Cells(3, "P").Value
                    targetWs.Cells(l, 31).Value = xlWs.Cells(3, "Q").Value
                    targetWs.Cells(l, 32).Value = xlWs.Cells(3, "R").Value
                    targetWs.Cells(l, 33).Value = xlWs.Cells(3, "AA").Value
                    
                    If (lastRow > 7) And (xlWs.Cells(lastRow, "CE").Value < 0) Then
                        return_code.Result = "Eq_pm"
                        RevealPayment = return_code
                        Exit Function
                    End If
                    
                    targetWs.Cells(l, 10).Interior.Color = RGB(0, 255, 0)
                    Exit For
                End If
            End If
        Next l
    End If
    
    ThisWorkbook.Save
    xlWb.Close SaveChanges:=False
    'xlMain.Close SaveChanges:=False
    
    return_code.Result = "N_pm"
    RevealPayment = return_code
    
    
End Function

