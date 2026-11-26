Attribute VB_Name = "UserForm1"

Attribute VB_Base = "0{08BB084F-0528-442A-966E-4102990CBFBA}{3A57493D-12C1-4275-AFA6-48214864DE0B}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = False

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = False

Dim fileName As String

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

Dim lastrow As Long

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

    Dim targetWb As Workbook

    Dim targetWs As Worksheet

    Dim tbl As ListObject

    Dim folderPath As String

    Dim selectedFile As String

    Dim latestFile As String

    Dim fileName As String

    Dim latestDate As Date

    Dim fileDate As Date

    Dim fso As Object, folder As Object, file As Object



    ' Define the default folder path to search

    folderPath = "C:\card_data\payment\" ' <-- Adjust this path as needed

    

    ' Value in this worksheet table

    Set targetWb = ThisWorkbook

    Set targetWs = targetWb.Sheets("Hold_Cutting")

    Set tbl = targetWs.ListObjects("Payment_list")

    

    ' Initialize variables

    latestFile = ""

    latestDate = ThisWorkbook.Sheets("Hold_Cutting").Range("AO11").Value ' A very old default date



    ' Try to find the latest Excel file in the folder

    If Dir(folderPath, vbDirectory) <> "" Then

        Set fso = CreateObject("Scripting.FileSystemObject")

        Set folder = fso.GetFolder(folderPath)



        For Each file In folder.Files

            If LCase(Right(file.Name, 5)) = ".xlsx" Or LCase(Right(file.Name, 4)) = ".xls" Then

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

        selectedFile = latestFile

    Else

        With Application.FileDialog(msoFileDialogFilePicker)

            .Title = "Select the Excel file manually"

            .Filters.Clear

            .Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsm"

            If .Show = -1 Then

                selectedFile = .SelectedItems(1)

            Else

                MsgBox "No file selected. Exiting."

                Exit Sub

            End If

        End With

    End If



    ' Put the selected file path into TextBox1 if it exists

    On Error Resume Next

    TextBox1.Value = Left(selectedFile, InStrRev(selectedFile, Application.PathSeparator)) & Mid(selectedFile, InStrRev(selectedFile, Application.PathSeparator) + 1)

    On Error GoTo 0



    ' Find the position of "C:\Users\" in the path

    Dim startPos As Integer, endPos As Integer

    Dim User_name As String

    

    startPos = InStr(selectedFile, "C:\Users\")

    If startPos > 0 Then

        User_name = Mid(selectedFile, startPos + Len("C:\Users\"))

        endPos = InStr(User_name, "\")

        If endPos > 0 Then

            User_name = Left(User_name, endPos - 1)

            ThisWorkbook.Sheets("Hold_Cutting").Range("AO17").Value = User_name

            MsgBox "Username: " & User_name

        End If

    End If

    

    'set excel instance

    Set xlApp = CreateObject("Excel.Application")

    'foundFile = FindExcelFile("C:\Users\" & User_name & "\Shared Drive\Team Data\Data For QMC\Payment Term\Cutoff-beta\Payment report\", "Data_Status")

    On Error Resume Next

    

    If test_mode Then

        MainPath = "C:\tools\cut-off\Testing env\_Data_for_Cut_off_Table.xlsx"

        MacroPath = "C:\tools\cut-off\Cut_off\Cut-Off-Database_marco.xlsm"

    Else

        MainPath = TextBox2.Value

        MacroPath = TextBox3.Value

    End If

    On Error GoTo 0

    

    Set xlMain = xlApp.Workbooks.Open(MainPath)

    Set xlMainS = xlMain.Sheets("DATA")

    

    Set xlMacro = xlApp.Workbooks.Open(ThisWorkbook.Sheets("Hold_Cutting").Range("AR13").Value)

    Set xlMacroS = xlMacro.Sheets("payment_mode_status")

    

    ' Open the selected workbook

    On Error Resume Next

    Set wb = xlApp.Workbooks.Open(selectedFile, password:=ThisWorkbook.Sheets("Hold_Cutting").Range("AO15").Value)

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

    tbl.ListColumns(12).DataBodyRange.ClearFormats

    

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

    Set SourceRange = ws.Range("S2:S" & row_count)

    Set targetRange = targetWs.Cells(startRow, 5)

    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value

    

    ' Copy Payment info

    Set SourceRange = ws.Range("K2:L" & row_count)

    Set targetRange = targetWs.Cells(startRow, 11)

    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value

    Set SourceRange = ws.Range("Q2:Q" & row_count)

    Set targetRange = targetWs.Cells(startRow, 13)

    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value

    Set SourceRange = ws.Range("M2:P" & row_count)

    Set targetRange = targetWs.Cells(startRow, 14)

    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value

    targetWs.Range("P6:Q" & row_count + startRow).NumberFormat = "dd-mm-yy"

    

    ' Copy Note

    Set SourceRange = ws.Range("T2:T" & row_count)

    Set targetRange = targetWs.Cells(startRow, 18)

    targetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value

    

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

        MainPath = "C:\tools\cut-off\Testing env\_Data_for_Cut_off_Table.xlsx"

        MacroPath = "C:\tools\cut-off\Cut_off\Cut-Off-Database_marco.xlsm"

    Else

        MainPath = TextBox2.Value

        MacroPath = ThisWorkbook.Sheets("Hold_Cutting").Range("AQ13").Value 'TextBox3.Value

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

    

    lastrow = tbl.ListColumns(2).DataBodyRange.Cells(tbl.ListRows.Count, 1).End(xlUp).Row - tbl.HeaderRowRange.Row

    For i = 1 To ThisWorkbook.Sheets("Hold_Cutting").Range("AO18").Value

    

        ' Check CountPay before run update

        If tbl.ListColumns(33).DataBodyRange.Cells(i, 1).Value = vbNullString And tbl.ListColumns(16).DataBodyRange.Cells(i, 1).Value <> vbNullString And tbl.ListColumns(17).DataBodyRange.Cells(i, 1).Value <> vbNullString Then

            

            searchText = tbl.ListColumns(2).DataBodyRange.Cells(i, 1).Value

            

            Dim success As Boolean

            Dim newFilePath As String

            

            T_code = LogicTemplate(tbl.ListColumns(12).DataBodyRange.Cells(i, 1).Value)

            foundFile = FindExcelFile(folderPath, searchText)

        

            If foundFile <> vbNullString Then

                'MsgBox "File found: " & foundFile

                

                'response = MsgBox("Do you want to update values?", vbQuestion + vbYesNo, "Update Confirmation")

            

                'If response = vbYes Then

            

                    'Attempt to open the file with the provided password

                    Set xlWb = xlApp.Workbooks.Open(foundFile)

                    Set xlWs = xlWb.Sheets("Transaction")

                    Set tbl_trans = xlWs.ListObjects("Transaction")

                    

                    If xlWb.Sheets("Payment Term History").Cells(1, "A").Value <> ThisWorkbook.Sheets("Hold_Cutting").Range("AO13").Value Then

                        xlWb.Sheets("Payment Term History").Cells(1, "A").Value = ThisWorkbook.Sheets("Hold_Cutting").Range("AO13").Value

                    End If

    

                    ' Find the latest blank row in xlWs

                    lastRowTB = 2 + xlWb.Sheets("Transaction").Range("AC2").Value

                    

                    If lastRowTB > 30 Then

                        MsgBox "Expanding formula range"

                        Dim insert_count As Integer

                        insert_count = 8

                        Do

                            insert_count = insert_count + 1

                            xlWb.Sheets("Payment Term History").Range("A" & insert_count & ":EH" & insert_count).Copy

                        Loop While xlWb.Sheets("Payment Term History").Range("AA" & insert_count + 1).Value <> vbNullString

                        xlWb.Sheets("Payment Term History").Cells(insert_count + 1, 1).EntireRow.Insert Shift:=xlDown

                        

                    End If

                    

                    If xlWs.Cells(lastRowTB - 1, "Q").Value = tbl.ListColumns(16).DataBodyRange.Cells(i, 1).Value And xlWs.Cells(lastRowTB - 1, "R").Value = tbl.ListColumns(17).DataBodyRange.Cells(i, 1).Value And Application.WorksheetFunction.Max(xlWs.Cells(lastRowTB - 1, "T").Value, xlWs.Cells(lastRowTB - 1, "V").Value) = tbl.ListColumns(14).DataBodyRange.Cells(i, 1).Value Then

                        response_next = MsgBox("This may be a dupplicate Transaction, still want to continue?", vbQuestion + vbYesNo, "Update Confirmation")

                        If response_next = vbNo Then

                            xlApp.Quit

                            Set xlApp = Nothing

                            Exit Sub

                        End If

                    End If

                    

                    tbl_trans.ListRows.Add

                    xlWs.Cells(lastRowTB, "A").Formula = "=ROW()-1"

                    

                    If tbl.ListColumns(12).DataBodyRange.Cells(i, 1).Value <> vbNullString And (tbl.ListColumns(18).DataBodyRange.Cells(i, 1).Value <> "edit" Or tbl.ListColumns(18).DataBodyRange.Cells(i, 1).Value <> "delete") Then

                        

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

                        xlWs.Cells(lastRowTB, "H").Value = tbl.ListColumns(9).DataBodyRange.Cells(i, 1).Value

                        xlWs.Cells(lastRowTB, "I").Value = tbl.ListColumns(10).DataBodyRange.Cells(i, 1).Value

                        xlWs.Cells(lastRowTB, "L").Value = tbl.ListColumns(11).DataBodyRange.Cells(i, 1).Value

                        xlWs.Cells(lastRowTB, "O").Value = tbl.ListColumns(13).DataBodyRange.Cells(i, 1).Value

                        

                        xlWs.Cells(lastRowTB, "P").Value = CDate(Format(ThisWorkbook.Sheets("Hold_Cutting").Range("AO11").Value, "dd/mm/yyyy"))

                    Else

                        card_update = tbl.ListColumns(18).DataBodyRange.Cells(i, 1).Value

                        For N = 1 To xlWs.Cells(2, "AC").Value + 1

                            'check date

                            If xlWs.Cells(N, "R").Value = tbl.ListColumns(17).DataBodyRange.Cells(i, 1).Value Then

                                Select Case card_update

                                    Case "edit"

                                        xlWs.Cells(N, "T").Value = tbl.ListColumns(14).DataBodyRange.Cells(i, 1).Value

                                    Case "delete"

                                        response_del = MsgBox("This action will remove the selected Transaction, still want to continue?", vbQuestion + vbYesNo, "Update Confirmation")

                                        If response_del = vbNo Then

                                            xlApp.Quit

                                            Set xlApp = Nothing

                                            Exit Sub

                                        Else

                                            xlWs.ListObjects("Transaction").ListRows(N - 1).Delete



                                        End If

                                End Select

                            End If

                        Next N

                    End If

                    

                    'Logic Payment Function

                    If T_code.Result <> vbNullString Then

                            case_select = tbl.ListColumns(12).DataBodyRange.Cells(i, 1).Value

                            Select Case True

                                Case case_select = "pm"

                                    xlWs.Cells(lastRowTB, "T").Value = tbl.ListColumns(14).DataBodyRange.Cells(i, 1).Value

                                    xlWs.Cells(lastRowTB, "N").Value = "pm"

                                Case case_select Like "*_d"

                                    If Left(case_select, Len(case_select) - 2) = "pm" Then

                                        xlWs.Cells(lastRowTB, T_code.Result).Value = (tbl.ListColumns(14).DataBodyRange.Cells(i, 1).Value * 100 / 107) - tbl.ListColumns(15).DataBodyRange.Cells(i, 1).Value

                                    Else

                                        xlWs.Cells(lastRowTB, T_code.Result).Value = tbl.ListColumns(14).DataBodyRange.Cells(i, 1).Value

                                        xlWs.Cells(lastRowTB, "W").Value = tbl.ListColumns(15).DataBodyRange.Cells(i, 1).Value

                                    End If

                                    xlWs.Cells(lastRowTB, "X").Value = "Direct"

                                    xlWs.Cells(lastRowTB, "N").Value = Left(case_select, Len(case_select) - 2)

                                    

                                Case case_select Like "*_a"

                                    If Left(case_select, Len(case_select) - 2) = "pm" Then

                                        xlWs.Cells(lastRowTB, T_code.Result).Value = tbl.ListColumns(14).DataBodyRange.Cells(i, 1).Value - tbl.ListColumns(15).DataBodyRange.Cells(i, 1).Value

                                    Else

                                        xlWs.Cells(lastRowTB, T_code.Result).Value = (tbl.ListColumns(14).DataBodyRange.Cells(i, 1).Value * 107 / 100)

                                        xlWs.Cells(lastRowTB, "W").Value = tbl.ListColumns(15).DataBodyRange.Cells(i, 1).Value

                                    End If

                                    xlWs.Cells(lastRowTB, "X").Value = "Auction"

                                    xlWs.Cells(lastRowTB, "N").Value = Left(case_select, Len(case_select) - 2)

                                    

                                Case case_select = "dis"

                                    Dim numberAfterSlash As Integer

                                    Dim slashPosition As Integer

                                    Dim total_hc_receive As Long

                                    

                                    slashPosition = InStr(tbl.ListColumns(18).DataBodyRange.Cells(i + 1, 1).Value, "/")

                                    numberAfterSlash = Mid(tbl.ListColumns(18).DataBodyRange.Cells(i + 1, 1).Value, slashPosition + 1)

                                    

                                    total_hc_receive = 0

                                        

                                    'discount logic

                                    If numberAfterSlash = 1 Then

                                        total_hc_receive = tbl.ListColumns(14).DataBodyRange.Cells(i + 1, 1).Value

                                        xlWs.Cells(lastRowTB, "Y").Value = 1 - (total_hc_receive / xlWb.Sheets("Payment Term History").Cells(lastRowTB + 6, "CI").Value)

                                    Else

                                        For m = 1 To numberAfterSlash

                                            total_hc_receive = total_hc_receive + tbl.ListColumns(14).DataBodyRange.Cells(i + m, 1).Value

                                        Next m

                                        

                                        Dim paid_count As Integer

                                        Dim paid_amount As Long

                                        

                                        paid_amount = 0

                                        For paid_count = 1 To numberAfterSlash

                                            paid_amount = paid_amount + xlWb.Sheets("Payment Term History").Cells(lastRowTB + 6 - paid_count, "T").Value

                                        Next paid_count

                                        

                                        paid_amount = paid_amount + tbl.ListColumns(14).DataBodyRange.Cells(i, 1).Value

                                        

                                        xlWs.Cells(lastRowTB, "Y").Value = 1 - (paid_amount / xlWb.Sheets("Payment Term History").Cells(lastRowTB - numberAfterSlash + 7, "CI").Value)

                                        tbl.ListColumns(16).DataBodyRange.Cells(i, 1).Value = (xlWs.Cells(lastRowTB - numberAfterSlash + 1, "R").Value - 2)

                                        tbl.ListColumns(17).DataBodyRange.Cells(i, 1).Value = (xlWs.Cells(lastRowTB - numberAfterSlash + 1, "R").Value - 2)

                                    End If

                                    

                                    tbl.ListColumns(14).DataBodyRange.Cells(i, 1).Value = xlWs.Cells(lastRowTB, T_code.Result).Value

                                    xlWs.Cells(lastRowTB, "N").Value = "dis"

                                    

                                Case Else

                                    xlWs.Cells(lastRowTB, T_code.Result).Value = tbl.ListColumns(14).DataBodyRange.Cells(i, 1).Value

                                    xlWs.Cells(lastRowTB, "N").Value = tbl.ListColumns(12).DataBodyRange.Cells(i, 1).Value

                            End Select

                        

                        xlWs.Cells(lastRowTB, "Q").Value = CDate(Format(tbl.ListColumns(16).DataBodyRange.Cells(i, 1).Value, "mm/dd/yyyy"))

                        xlWs.Cells(lastRowTB, "R").Value = CDate(Format(tbl.ListColumns(17).DataBodyRange.Cells(i, 1).Value, "mm/dd/yyyy"))

                    End If

                        

                    ' Recalculate the worksheet

                    xlApp.Calculate

                    xlWb.Close SaveChanges:=True

                    'xlMain.Close SaveChanges:=True

                    'xlApp.Quit

                    'Set xlApp = Nothing

            

                    err_code = RevealPayment(foundFile, MainPath, T_code.Result, searchText, xlApp)

                    

                    Select Case err_code.Result

                        Case "Miss"

                            tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Database Missing"

                        Case "Value"

                            tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Database Error"

                        Case "Path"

                            tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Path error"

                        Case "Error"

                            tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Checking Request"

                        Case "Eq_pm"

                            tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Payment Error"

                        Case Else

                            tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Update Payment"

                    End Select

            

                    

            

            Else

                    'MsgBox "File not found."

        

                    'response = MsgBox("Do you want to Create new profile account_id: " & searchText & " ?", vbQuestion + vbYesNo, "Update Confirmation")

                

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

                                tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Path error"

                                Exit Sub

                            End If

                        

                            Select Case err_code.Result

                                Case "Miss"

                                    tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Database Missing"

                                Case "Value"

                                    tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Database Error"

                                Case "Path"

                                    tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Path error"

                                Case "Error"

                                    tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Checking Request"

                                Case "Eq_pm"

                                    tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Payment Error"

                                Case Else

                                    tbl.ListColumns(38).DataBodyRange.Cells(i, 1).Value = "Create new card"

                            End Select

                    

            End If

            

            tbl.ListColumns(4).DataBodyRange.Cells(i, 1).Interior.Color = RGB(150, 200, 0)



            TextBox4.Text = "Updating : " & i & " / " & ThisWorkbook.Sheets("Hold_Cutting").Range("AO18").Value

            

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

            ThisWorkbook.Sheets("Hold_Cutting").Range("AQ13").Value = Left(selectedFile, InStrRev(selectedFile, Application.PathSeparator)) & Mid(selectedFile, InStrRev(selectedFile, Application.PathSeparator) + 1)

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

            ThisWorkbook.Sheets("Hold_Cutting").Range("AR13").Value = Left(selectedFile, InStrRev(selectedFile, Application.PathSeparator)) & Mid(selectedFile, InStrRev(selectedFile, Application.PathSeparator) + 1)

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

        folderPath = "C:\tools\cut-off\Testing env\"

        MainPath = "C:\tools\cut-off\Testing env\_Data_for_Cut_off_Table.xlsx"

        MacroPath = "C:\tools\cut-off\Cut_off\Cut-Off-Database_marco.xlsm"

    Else

        MainPath = ThisWorkbook.Sheets("Hold_Cutting").Range("AQ13").Value 'TextBox2.Value

        MacroPath = TextBox3.Value

        folderPath = Left(MainPath, InStrRev(MainPath, "\"))

    End If

    On Error GoTo 0

    

    'Set xlApp = CreateObject("Excel.Application")

    Dim year_f_name As String

    Dim yearValue As String

    Dim yearPosition As Integer

    

    ' Extract the file name from the path

    year_f_name = Mid(MacroPath, InStrRev(MacroPath, "\") + 1)

    

    ' Find the year in the file name

    yearPosition = InStr(year_f_name, "-") ' Look for the first hyphen as a reference

    If yearPosition > 0 Then

        yearValue = Mid(year_f_name, Len(year_f_name) - 8, 4) ' Extract the year

    Else

    End If

    

    ' Value in this worksheet table

    Set targetWb = ThisWorkbook

    Set targetWs = targetWb.Sheets("Hold_Cutting")

    Set tbl = targetWs.ListObjects("Payment_list")

    

    ' Macro path

    Set xlMacro = Workbooks.Open(MacroPath, password:=Environ("MACRO_PASSWORD")  'redacted: real password comes from the environment)

    Set xlMacroS = xlMacro.Sheets("Daily_Report")

    Set tbMac = xlMacroS.ListObjects("Data_daily")

    

    lr = ThisWorkbook.Sheets("Hold_Cutting").Range("AO18").Value 'tbl.ListRows.Count ' Get the number of rows in the table

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

            

            If (targetWs.Cells(copy_i, "AG") <> vbNullString) And (targetWs.Cells(copy_i, "A") = vbNullString) Then



                'status mode

                targetWs.Range("G" & copy_i & ":H" & copy_i).Copy

                xlMacroS.Cells(l, 12).PasteSpecial xlPasteValues

                'count payment

                'targetWs.Cells(copy_i, 28).Copy

                'xlMacroS.Cells(l, 28).PasteSpecial xlPasteValues

                'pm_date, eff_date

                If targetWs.Cells(copy_i, 18).Value <> "delete" Or targetWs.Cells(copy_i, 18).Value <> "edit" Then

                    xlMacroS.Cells(l, 30).Value = targetWs.Cells(copy_i, 16).Value

                    xlMacroS.Cells(l, 31).Value = targetWs.Cells(copy_i, 17).Value

                End If

                

                If (targetWs.Cells(copy_i, "L") = "vd") Then

                    If xlMacroS.Cells(l, 32) < targetWs.Cells(copy_i, 34).Value Then

                        'Balance and vat

                        targetWs.Range("S" & copy_i & ":AG" & copy_i).Copy

                        xlMacroS.Cells(l, 14).PasteSpecial xlPasteValues

                        'total payment vat

                        targetWs.Range("AH" & copy_i & ":AK" & copy_i).Copy

                        xlMacroS.Cells(l, 32).PasteSpecial xlPasteValues

                    Else

                        'Balance and vat

                        targetWs.Range("S" & copy_i & ":AF" & copy_i).Copy

                        xlMacroS.Cells(l, 14).PasteSpecial xlPasteValues

                    End If

 

                ElseIf (targetWs.Cells(copy_i, "L") = "dis") Then

                    If xlMacroS.Cells(l, 32) < targetWs.Cells(copy_i, 34).Value Then

                        'Balance and vat

                        targetWs.Range("S" & copy_i & ":AG" & copy_i).Copy

                        xlMacroS.Cells(l, 14).PasteSpecial xlPasteValues

                        'total payment vat

                        targetWs.Range("AH" & copy_i & ":AK" & copy_i).Copy

                        xlMacroS.Cells(l, 32).PasteSpecial xlPasteValues

                    Else

                        'Balance and vat

                        targetWs.Range("S" & copy_i & ":AF" & copy_i).Copy

                        xlMacroS.Cells(l, 14).PasteSpecial xlPasteValues

                    End If

                        

                Else

                    'Balance and vat

                    targetWs.Range("S" & copy_i & ":AG" & copy_i).Copy

                    xlMacroS.Cells(l, 14).PasteSpecial xlPasteValues

                    

                    If targetWs.Cells(copy_i, 18).Value <> "delete" Or targetWs.Cells(copy_i, 18).Value <> "edit" Then

                        'Amount

                        targetWs.Cells(copy_i, 14).Copy

                        xlMacroS.Cells(l, 29).PasteSpecial xlPasteValues

                    End If



                    'xlMacroS.Cells(l, 28).Value = targetWs.Cells(copy_i, 24).Value

                    'total payment vat

                    targetWs.Range("AH" & copy_i & ":AK" & copy_i).Copy

                    xlMacroS.Cells(l, 32).PasteSpecial xlPasteValues

                End If

                'DocID

                'targetWs.Cells(copy_i, 9).Copy

                'xlMacroS.Cells(l, 39).PasteSpecial xlPasteValues

                

                'payment_discount

                temp_dis = Payment_code(targetWs.Cells(copy_i, 12).Value)

                If xlMacroS.Cells(l, 36) <> temp_dis.Result Then

                    Select Case targetWs.Cells(copy_i, 12)

                        Case "dis"

                            If targetWs.Cells(copy_i, 14).Value <> vbNullString Then

                                xlMacroS.Cells(l, 36) = "Discount"

                                xlMacroS.Cells(l, 37).Value = targetWs.Cells(copy_i, 14).Value

                                xlMacroS.Cells(l, 38).Value = targetWs.Cells(copy_i, 14).Value * (targetWs.Cells(copy_i, 29).Value / (1 - targetWs.Cells(copy_i, 14).Value))

                                xlMacroS.Cells(l, 39).Value = targetWs.Cells(copy_i, 14).Value * (targetWs.Cells(copy_i, 31).Value / (1 - targetWs.Cells(copy_i, 14).Value))

                            End If

                        Case "add"

                            xlMacroS.Cells(l, 36) = "Normal"

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

    'PivotTable1

    xlMacro.Sheets("OA").PivotTables("PivotTable1").RefreshTable



    ' Save the workbook with password protection

    'SaveWorkbookWithPassword "DailyReport_" & ThisWorkbook.Sheets("Hold_Cutting").Range("AK11").Value & ".xlsx", xlMacro, ThisWorkbook.Sheets("Hold_Cutting").Range("AK15").Value

    SaveWorkbookWithDialog "DailyReport_" & ThisWorkbook.Sheets("Hold_Cutting").Range("AO11").Value & ".xlsx", xlMacro, ThisWorkbook.Sheets("Hold_Cutting").Range("AO15").Value

    

    MsgBox ("Update Report Done")

   

    xlMacro.Close SaveChanges:=False

    

    ThisWorkbook.Sheets("Hold_Cutting").Range("AP13").Value = Left(MacroPath, InStrRev(MacroPath, "\")) & "DailyReport_" & ThisWorkbook.Sheets("Hold_Cutting").Range("AO11").Value & ".xlsx"

    'xlApp.Quit

    'Set xlApp = Nothing

    'Application.ScreenUpdating = True

    'Call SendEmail

    

End Sub

Sub SaveWorkbookWithPassword(ByVal fileName As String, ByVal wb As Workbook, ByVal password As String)

    ' Save workbook with password protection

    wb.SaveAs fileName:=fileName, password:=password

End Sub

Private Sub CommandButton5_Click()



    Dim iter_rows As Integer

    Dim port_fix, contract_fix

    

    port_fix = ""

    contract_fix = ""

    

    On Error Resume Next

    If test_mode Then

        MainPath = " "

        MacroPath = " "

    Else

        MainPath = TextBox2.Value

        MacroPath = ThisWorkbook.Sheets("Hold_Cutting").Range("AQ13").Value

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

    fileName = Dir(folderPath & "\StatementCard_*.xlsx")

    Do While fileName <> ""

        totalFiles = totalFiles + 1

        fileName = Dir

    Loop

    

    ' Loop through each file in the folder

    fileName = Dir(folderPath & "\StatementCard_*.xlsx")

    Do While fileName <> ""

        ' Open the Excel file

        Set xlWb = xlApp.Workbooks.Open(folderPath & fileName)

        Set xlWs = xlWb.Sheets("Payment Term History")

        

        ' Update the value of a specific cell / Formula

        xlWs.Cells(1, "A").Value = MainPath

        'xlWs.Cells(7, "C").Formula = "=TEXT(XLOOKUP(B4,INDIRECT(""'""&$A$1&""'!""&$D$1&""[account_code]""),INDIRECT(""'""&$A$1&""'!""&$D$1&""[old_contract_number]"")),""0000"")"

        

        'xlWs.Cells(7, "H").Value = "-"

        'xlWs.Cells(7, "I").Value = "-"

        'xlWs.Cells(7, "O").Value = "-"

        'xlWs.Cells(8, "B").Formula2 = "=SORT(FILTER(Transaction[[account_code]:[Addition]],Transaction[account_code]<>""""),17,1)"

        'xlWs.Cells(3, "AA").Value = ""

        

        Dim lastrow As Long

        

        lastrow = xlWb.Sheets("Transaction").Range("AC2").Value

        

        If 2 + lastrow > 30 Then

            lastrow = 15 + lastrow

        Else

            lastrow = 44

        End If

        

        ' Loop through rows 7 to 44

        With xlWs

            ' Insert the formula into column 41 (column AO)

            .Range("AC7:AC" & lastrow).Formula2 = "=IF(N7=""sold"",7%,$AC$7)"

            .Range("AD7:AD" & lastrow).Formula2 = "=IFERROR(IFS(G6=""jm_vat"",(BT6+BW6+CC6)/CG6,G6=""npl_vat"",(BQ6+BZ6)/CG6,G6=""cm_vat"",(BT6+BW6+BZ6)/CG6,ISNUMBER(FIND(""_all_vat"",G7)),CF6/CG6,NOT(ISNUMBER(FIND(""_vat"",G7))),0,G6=0,0,AND(N7=""sold"",$DG$7=""hp""),0),0)"

            .Range("AE7:AE" & lastrow).Formula2 = "=IF(AND(N7=""sold"",$DG$7=""hp""),V7-(V7*(1/(1+AC7)))," & _

                                            "IF(ISNUMBER(FIND(""_vat"",G7)),AG7-(AG7*(1/(1+AC7)))," & _

                                            "IF(MAX(T7,V7)<>0,MAX(T7,V7)-(MAX(T7,V7)*(1/(1+AC7))),0)))"

            .Range("AF7:AF" & lastrow).Formula2 = "=MAX(T7,V7)-AE7"

            .Range("AG7:AG" & lastrow).Formula2 = "=IF(AND(ISNUMBER(FIND(""_vat"",G7)),AD7>0),(MAX(T7,V7)-AH7),0)"

            .Range("AH7:AH" & lastrow).Formula2 = "=MIN(MAX(T7,V7)*(1-AD7),CK6)"

            

            .Range("AL7:AL" & lastrow).Formula2 = "=IF(AND(N7=""sold"",$DG$7=""hp""),MIN(AF7,AM6+AK7)," & _

                                            "IF(MAX(T7,V7)<>0,MIN(MAX(T7,V7)-AE7,AM6)," & _

                                            "IF(N7=""repo"",0,0)))"

            .Range("AQ7:AQ" & lastrow).Formula2 = "=IF(N7=""pm"",MIN(T7-AE7-AL7,AR6+AP7)," & _

                                            "IF(N7=""repo"",0," & _

                                            "IF(N7=""sold"",MIN(AF7-AL7,AR6+AP7),0)))"

            .Range("AT7:AT" & lastrow).Formula2 = "=IF(AND(N7=""sold"",$DG$7=""hp""),MIN(AF7-AL7-AQ7,AU6)," & _

                                            "IF(MAX(T7,V7)<>0,MIN((MAX(T7,V7)-AE7-AL7-AQ7),AU6)," & _

                                            "IF(N7=""repo"",0,0)))"

            .Range("AW7:AW" & lastrow).Formula2 = "=IF(AND(N7=""sold"",$DG$7=""hp""),AF7-AL7-AQ7-AT7," & _

                                            "IF(MAX(T7,V7)<>0,MAX(T7,V7)-AE7-AL7-AQ7-AT7," & _

                                            "IF(N7=""repo"",0,0)))"

            

            .Range("BC7:BC" & lastrow).Formula2 = "=IF(N7=""dis"",Y7*BD6," & _

                                            "IF(N7=""vd"",BD6," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",G7=""npl_vat""),MIN(AF7-BP7-BY7,BD6+BB7)," & _

                                            "IF(OR(G7=""jm_vat"",G7=""npl_vat"",G7=""cm_vat""),MIN(AH7,BD6)," & _

                                            "IF(OR(G7=""jm"",G7=""npl"",G7=""cm""),MIN(AF7,BD6)," & _

                                            "IF(OR(G7=""jm_all_vat"",G7=""npl_all_vat"",G7=""cm_all_vat""),MIN(AH7,BD6),0))))))"

            .Range("BF7:BF" & lastrow).Formula2 = "=IF(N7=""dis"",Y7*BG6," & _

                                            "IF(N7=""vd"",BG6," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",G7=""npl_vat""),MIN(AF7-BP7-BY7-BC7,BG6)," & _

                                            "IF(OR(G7=""jm_vat"",G7=""npl_vat"",G7=""cm_vat""),MIN(AH7-BC7,BG6)," & _

                                            "IF(OR(G7=""jm"",G7=""npl"",G7=""cm""),MIN(AF7-BC7,BG6),0)))))"

            .Range("BI7:BI" & lastrow).Formula2 = "=IF(N7=""dis"",Y7*BJ6," & _

                                            "IF(N7=""vd"",BJ6," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",G7=""npl_vat""),MIN(AF7-BP7-BY7-BC7-BF7,BJ6)," & _

                                            "IF(OR(G7=""jm_vat"",G7=""npl_vat"",G7=""cm_vat""),MIN(AH7-BC7-BF7,BJ6)," & _

                                            "IF(OR(G7=""jm"",G7=""npl"",G7=""cm""),MIN(AF7-BC7-BF7,BJ6),0)))))"

            .Range("BL7:BL" & lastrow).Formula2 = "=IF(N7=""dis"",Y7*BM6," & _

                                            "IF(N7=""vd"",BM6," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",G7=""npl_vat""),MIN(AF7-BP7-BY7-BC7-BF7-BI7,BM6)," & _

                                            "IF(OR(G7=""jm_vat"",G7=""npl_vat"",G7=""cm_vat""),MIN(AH7-BC7-BF7-BI7,BM6)," & _

                                            "IF(OR(G7=""jm"",G7=""npl"",G7=""cm""),MIN(AF7-BC7-BF7-BI7,BM6),0)))))"

            .Range("BP7:BP" & lastrow).Formula2 = "=IF(N7=""dis"",Y7*BQ6," & _

                                            "IF(N7=""vd"",BQ6," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",G7=""npl_vat""),MIN(AF7*(BQ6/(BQ6+BZ6)),BQ6)," & _

                                            "IF(OR(G7=""jm_vat"",G7=""cm_vat""),MIN(AH7-BC7-BF7-BI7-BL7,BQ6)," & _

                                            "IF(OR(G7=""jm"",G7=""npl"",G7=""cm""),MIN(AF7-BC7-BF7-BI7-BL7,BQ6)," & _

                                            "IF(G7=""npl_vat"",MIN(AG7-AE7,BQ6),0))))))"

            .Range("BS7:BS" & lastrow).Formula2 = "=IF(N7=""dis"",Y7*BT6," & _

                                            "IF(N7=""vd"",BT6," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",G7=""npl_vat""),MIN(AF7-BP7-BY7-BC7-BF7-BI7-BL7,BT6)," & _

                                            "IF(OR(G7=""jm_vat"",G7=""cm_vat""),MIN(AG7-AE7,BT6)," & _

                                            "IF(OR(G7=""jm"",G7=""npl"",G7=""cm""),MIN(AF7-BC7-BF7-BI7-BL7-BP7,BT6)," & _

                                            "IF(G7=""npl_vat"",MIN(AH7-BC7-BF7-BI7-BL7,BT6),0))))))"

            .Range("BV7:BV" & lastrow).Formula2 = "=IF(N7=""dis"",Y7*BW6," & _

                                            "IF(N7=""vd"",BW6," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",G7=""npl_vat""),MIN(AF7-BP7-BY7-BC7-BF7-BI7-BL7-BS7,BW6)," & _

                                            "IF(OR(G7=""jm_vat"",G7=""cm_vat""),MIN(AG7-AE7-BS7,BW6)," & _

                                            "IF(OR(G7=""jm"",G7=""npl"",G7=""cm""),MIN(AF7-BC7-BF7-BI7-BL7-BP7-BS7,BW6)," & _

                                            "IF(G7=""npl_vat"",MIN(AH7-BC7-BF7-BI7-BL7-BS7,BW6),0))))))"

            .Range("BY7:BY" & lastrow).Formula2 = "=IF(N7=""dis"",Y7*BZ6," & _

                                            "IF(N7=""vd"",BZ6," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",G7=""npl_vat""),MIN(AF7-BP7,BZ6)," & _

                                            "IF(G7=""jm_vat"",MIN(AH7-BC7-BF7-BI7-BL7-BP7,BZ6)," & _

                                            "IF(OR(G7=""jm"",G7=""npl"",G7=""cm""),MIN(AF7-BC7-BF7-BI7-BL7-BP7-BS7-BV7,BZ6)," & _

                                            "IF(G7=""npl_vat"",MIN(AG7-AE7-BP7,BZ6)," & _

                                            "IF(G7=""cm_vat"",MIN(AG7-AE7-BS7-BV7,BZ6),0)))))))"

            .Range("CB7:CB" & lastrow).Formula2 = "=IF(N7=""dis"",Y7*CC6," & _

                                            "IF(N7=""vd"",CC6," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",G7=""jm_vat""),MIN(AF7,CC6)," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",G7=""npl_vat""),MIN(AF7-BP7-BY7-BC7-BF7-BI7-BL7-BS7-BV7,CC6)," & _

                                            "IF(G7=""jm_vat"",MIN(AG7-AE7-BS7-BV7,CC6)," & _

                                            "IF(OR(G7=""jm"",G7=""npl"",G7=""cm""),MIN(AF7-BC7-BF7-BI7-BL7-BP7-BS7-BV7-BY7,CC6)," & _

                                            "IF(G7=""npl_vat"",MIN(AH7-BC7-BF7-BI7-BL7-BS7-BV7,CC6)," & _

                                            "IF(G7=""cm_vat"",MIN(AH7-BC7-BF7-BI7-BL7-BP7,CC6),0))))))))"

            .Range("CE7:CE" & lastrow).Formula2 = "=IF(N7=""dis"",Y7*CF6," & _

                                            "IF(N7=""vd"",CF6," & _

                                            "IF(AND(N7=""sold"",$DG$7=""hp"",OR(G7=""jm_all_vat"",G7=""npl_all_vat"",G7=""cm_all_vat"")),MIN(AF7,CF6)," & _

                                            "IF(OR(G7=""jm_all_vat"",G7=""npl_all_vat"",G7=""cm_all_vat""),MIN(AG7-AE7,CF6)," & _

                                            "IF(OR(G7=""jm_all"",G7=""npl_all"",G7=""cm_all""),MIN(AF7,CF6),0)))))"

            

            .Range("CX7:CX" & lastrow).Formula2 = "=IF(AND(N7=""sold"",$DG$7=""hp""),MIN(AF7,CY6+CW7)," & _

                                            "IF(MAX(T7,V7)<>0,MIN(MAX(T7,V7)-AE7,CY6),0))"

            .Range("CZ7:CZ" & lastrow).Formula2 = "=IF(AND(N7=""sold"",$DG$7=""hp""),(AF7-CX7)*CV7," & _

                                            "IF(MAX(T7,V7)<>0,(MAX(T7,V7)-AE7-CX7)*CV7,0))"

            .Range("DB7:DB" & lastrow).Formula2 = "=IF(AND(N7=""sold"",$DG$7=""hp""),MIN(AF7-CX7-CZ7,DC6)," & _

                                            "IF(MAX(T7,V7)<>0,MIN((MAX(T7,V7)-AE7-CX7-CZ7),DC6),0))"

            .Range("DD7:DD" & lastrow).Formula2 = "=IF(AND(N7=""sold"",$DG$7=""hp""),AF7-CX7-CZ7-DB7," & _

                                            "IF(MAX(T7,V7)<>0,MAX(T7,V7)-AE7-CX7-CZ7-DB7,0))"

            

            .Range("DR7:DR" & lastrow).Formula2 = "=MAX(T7,V7)"

            .Range("DW7:DW" & lastrow).Formula2 = "=IF(OR(N7=""aq"",N7=""vd""),DV7/DU7,IF(OR(N7=""pm"",ROW(N7)-6<$K$3-1),DW6,DV7/DU7))"

            .Range("DY7:DY" & lastrow).Formula2 = "=IF(N7=""repo"",DZ6,IF(N7=""sold"",DT7,IF(ISERR(FILTER($N$7:N7,$N$7:N7=""sold"")),MIN(DT7*DW7,DZ6),IF(AND(DT7<>0,CI7=0),DZ6,0))))"

            '.Range("DZ7:DZ" & lastrow).Formula2 = "=UNIQUE(IF(N7=""aq"",DQ7,IF(N7=""sold"",FILTER($DY$7:DY7,$N$7:N7=""repo""),IF(ISERR(FILTER($N$7:N7,$N$7:N7=""sold"")),DZ6-DY7,0))))"

            .Range("EA7:EA" & lastrow).Formula2 = "=IF(OR(N7=""repo"",IFERROR(FILTER($N$7:N7,$N$7:N7=""sold"")=""sold"",FALSE)),0,DP7*DZ6*(DX7/365))"

            .Range("EB7:EB" & lastrow).Formula2 = "=IF(OR(N7=""repo"",N7=""sold""),0,MIN(DT7-DY7,EC6+EA7))"

            .Range("ED7:ED" & lastrow).Formula2 = "=IFS(N7=""repo"",0," & _

                                            "N7=""sold"",DY7-DZ7," & _

                                            "N7<>"""",DT7-DY7-EB7," & _

                                            "N7="""",0)"

            .Range("EE7:EE" & lastrow).Formula2 = "=IF($DG$7=""loan"",IF(ED7<0,0,EB7+ED7),0)"

            .Range("EG7:EG" & lastrow).Formula2 = "=IF(OR(N7=""repo"",N7=""sold""),0,DR7-DS7-DY7-EB7-ED7)"

        End With

        

        xlWb.Sheets("Payment Term History").Cells(7, "AC").Formula2 = "=XLOOKUP(B4,INDIRECT(""'""&$A$1&""'!""&$D$1&""[account_code]""),INDIRECT(""'""&$A$1&""'!""&$D$1&""[vat_rate]""))"

        

        Application.Calculation = xlCalculationAutomatic

        

        ' Get the new filename

        Dim newFileName As String

        newFileName = "StatementCard_" & xlWs.Cells(4, "B").Value & "_" & xlWs.Cells(7, "DG").Value & ".xlsx"

        

        ' Save and close the file

        xlWb.Close SaveChanges:=True

        

        'change file name

        If newFileName <> "" Then

        '    Name folderPath & "\" & filename As folderPath & "\" & newFileName

        End If

        

        processedFiles = processedFiles + 1

        TextBox4.Text = "Updating Name : " & processedFiles & " / " & totalFiles

        

        ' Get the next file in the folder

        fileName = Dir

    Loop

    

    MsgBox "Update Card Done"

    

End Sub



Private Sub CommandButton6_Click()



    Dim temp_code As String

    

    On Error Resume Next

    If test_mode Then

        MainPath = " "

        MacroPath = " "

    Else

        MainPath = TextBox2.Value

        MacroPath = ThisWorkbook.Sheets("Hold_Cutting").Range("AQ13").Value

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

    fileName = Dir(folderPath & "\StatementCard_*.xlsx")

    Do While fileName <> ""

        totalFiles = totalFiles + 1

        fileName = Dir

    Loop

    

    tbl.DataBodyRange.ClearContents

    tbl.ListColumns(1).DataBodyRange.ClearFormats

    tbl.ListColumns(4).DataBodyRange.ClearFormats

    tbl.ListColumns(12).DataBodyRange.ClearFormats

    

    ' Loop through each file in the folder

    processedFiles = 1

    fileName = Dir(folderPath & "\StatementCard_*.xlsx")

    Do While fileName <> ""

        ' Open the Excel file

        Set xlWb = xlApp.Workbooks.Open(folderPath & "\" & fileName)

        Set xlWs = xlWb.Sheets("Payment Term History")

        

        lastRowTB = 6 + xlWs.Cells(3, "K").Value

        ' Update the value of a specific cell to sheet

        'ID name

        If Left(xlWs.Cells(4, "B").Value, 1) = "0" Then

            tbl.ListColumns(2).DataBodyRange.Cells(processedFiles, 1).NumberFormat = "@"

        End If

        temp_code = xlWs.Cells(4, "B").Value

        tbl.ListColumns(2).DataBodyRange.Cells(processedFiles, 1).Value = temp_code

        tbl.ListColumns(3).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "C").Value

        tbl.ListColumns(4).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(7, "A").Value

        'Status mode Port

        tbl.ListColumns(5).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "D").Value

        tbl.ListColumns(6).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "E").Value

        tbl.ListColumns(7).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "F").Value

        tbl.ListColumns(8).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "G").Value

        tbl.ListColumns(9).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "H").Value

        tbl.ListColumns(10).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "I").Value

        'amount eff_date

        tbl.ListColumns(12).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "N").Value

        tbl.ListColumns(14).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "T").Value

        

        If (xlWs.Cells(lastRowTB, "N").Value = "pm") And (xlWs.Cells(lastRowTB, "T").Value > 0) Then

            tbl.ListColumns(14).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "T").Value

        ElseIf xlWs.Cells(lastRowTB, "V").Value > 0 Then

            tbl.ListColumns(14).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "V").Value

        Else

            

        End If

        

        Select Case xlWs.Cells(lastRowTB, "X").Value

            Case "Auction"

                tbl.ListColumns(12).DataBodyRange.Cells(processedFiles, 1).Value = tbl.ListColumns(12).DataBodyRange.Cells(processedFiles, 1).Value & "_a"

            Case "Direct"

                tbl.ListColumns(12).DataBodyRange.Cells(processedFiles, 1).Value = tbl.ListColumns(12).DataBodyRange.Cells(processedFiles, 1).Value & "_d"

            Case Else

        End Select

        

        tbl.ListColumns(16).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "Q").Value

        tbl.ListColumns(17).DataBodyRange.Cells(processedFiles, 1).Value = xlWs.Cells(lastRowTB, "R").Value

        xlWb.Close SaveChanges:=True

        

        foundFile = folderPath & fileName

        T_code = LogicTemplate(tbl.ListColumns(12).DataBodyRange.Cells(processedFiles, 1).Value)

        err_code = RevealPayment(foundFile, MainPath, T_code.Result, temp_code, xlApp)

                    

        Select Case err_code.Result

            Case "Miss"

                tbl.ListColumns(38).DataBodyRange.Cells(processedFiles, 1).Value = "Database Missing"

            Case "Value"

                tbl.ListColumns(38).DataBodyRange.Cells(processedFiles, 1).Value = "Database Error"

            Case "Path"

                tbl.ListColumns(38).DataBodyRange.Cells(processedFiles, 1).Value = "Path error"

            Case "Error"

                tbl.ListColumns(38).DataBodyRange.Cells(processedFiles, 1).Value = "Checking Request"

            Case "Eq_pm"

                tbl.ListColumns(38).DataBodyRange.Cells(processedFiles, 1).Value = "Payment Error"

            Case Else

                tbl.ListColumns(38).DataBodyRange.Cells(processedFiles, 1).Value = "Update Payment"

        End Select

        

        

        processedFiles = processedFiles + 1

        TextBox4.Text = "Recall : " & processedFiles - 1 & " / " & totalFiles

        

        ' Get the next file in the folder

        fileName = Dir

    Loop

    

    MsgBox "Recall Done"

End Sub



Private Sub CommandButton7_Click()

    ' Count the total number of files in the folder

    fileName = Dir("C:\card_data\summary\done-add-date" & "\StatementCard_*.xlsx")

    Do While fileName <> ""

        totalFiles = totalFiles + 1

        fileName = Dir

    Loop

    

    If totalFiles = 0 Then

        MsgBox ("Bring Statement Card to card_data -> summary -> file")

        Exit Sub

    Else

        TextBox4.Text = "Summary total  " & totalFiles & " File"

        Call SummaryCard



    End If

    

End Sub



Private Sub CommandButton8_Click()

    Call AssignOA

End Sub



Private Sub TextBox1_Change()

    ' Update the value in cell when TextBox1 changes

    ThisWorkbook.Sheets("Hold_Cutting").Range("AN13").Value = TextBox1.Value

End Sub

Private Sub TextBox2_Change()

    ThisWorkbook.Sheets("Hold_Cutting").Range("AO13").Value = TextBox2.Value

End Sub

Private Sub TextBox3_Change()

    ThisWorkbook.Sheets("Hold_Cutting").Range("AP13").Value = TextBox3.Value

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

    TextBox1.Value = ThisWorkbook.Sheets("Hold_Cutting").Range("AN13").Value

    TextBox2.Value = ThisWorkbook.Sheets("Hold_Cutting").Range("AO13").Value

    TextBox3.Value = ThisWorkbook.Sheets("Hold_Cutting").Range("AP13").Value

    User_name = ThisWorkbook.Sheets("Hold_Cutting").Range("AO17").Value

    

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

            Result.Result = "T"

        Case C_Code Like "sold*"

            Result.Result = "V"

        Case C_Code = "esue", C_Code = "sue"

            Result.Result = "U"

        Case C_Code = "dis"

            Result.Result = "Y"

        Case C_Code = "vd"

            Result.Result = "R"

        Case C_Code = "repo"

            Result.Result = "V"

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



Function UpdateExcelFile(foundFile As String, account_code As Variant, Nplace As Variant, p_code As String, newFilePath As String, MainPath As String, xlApp As Object) As Collection

    Dim results As New Collection

    Dim xlWb As Object

    Dim xlWs As Object

    Dim xlMain As Object

    Dim xlMainS As Object

    Dim lastrow As Long

    

    Set tbl = ThisWorkbook.Sheets("Hold_Cutting").ListObjects("Payment_list")

    

    'On Error Resume Next

    'Set xlApp = CreateObject("Excel.Application")

    'On Error GoTo 0

    

    'Set xlMain = xlApp.Workbooks.Open(MainPath)

    'Set xlMainS = xlMain.Sheets("DATA")

    

    Set xlWb = xlApp.Workbooks.Open(foundFile)

    Set xlWs = xlWb.Sheets("Payment Term History")

    Set xlWt = xlWb.Sheets("Transaction")



    ' Find the latest blank row in xlWs

    xlWs.Cells(4, 2).Value = account_code

    xlWs.Cells(7, "P").Value = CDate(Format(ThisWorkbook.Sheets("Hold_Cutting").Range("AO11").Value, "dd/mm/yyyy"))

    

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

    xlWt.Cells(2, "H").Value = tbl.ListColumns(9).DataBodyRange.Cells(Nplace, 1).Value

    xlWt.Cells(2, "I").Value = tbl.ListColumns(10).DataBodyRange.Cells(Nplace, 1).Value

    xlWt.Cells(2, "L").Value = tbl.ListColumns(11).DataBodyRange.Cells(Nplace, 1).Value

    xlWs.Cells(2, "O").Value = tbl.ListColumns(13).DataBodyRange.Cells(Nplace, 1).Value

    xlWt.Cells(2, "P").Value = CDate(Format(ThisWorkbook.Sheets("Hold_Cutting").Range("AO11").Value, "dd/mm/yyyy"))

    xlWt.Cells(2, "W").Value = tbl.ListColumns(15).DataBodyRange.Cells(Nplace, 1).Value

    

    'Logic Payment Function

    If p_code <> vbNullString Then

        case_select = tbl.ListColumns(12).DataBodyRange.Cells(Nplace, 1).Value

                        

        If case_select Like "sold_*" And InStr(1, tbl.ListColumns(8).DataBodyRange.Cells(Nplace, 1).Value, "cm") > 0 Then

            xlWt.Cells(2, "N").Value = "pm"

            Select Case case_select

                Case "sold_d"

                    xlWt.Cells(2, "T").Value = (tbl.ListColumns(14).DataBodyRange.Cells(Nplace, 1).Value * 100 / 107) - tbl.ListColumns(15).DataBodyRange.Cells(Nplace, 1).Value

                    xlWt.Cells(2, "X").Value = "Direct"

                Case "sold_a"

                    xlWt.Cells(2, "T").Value = tbl.ListColumns(14).DataBodyRange.Cells(Nplace, 1).Value - tbl.ListColumns(15).DataBodyRange.Cells(Nplace, 1).Value

                    xlWt.Cells(2, "X").Value = "Auction"

            End Select

    

        Else

            Select Case True

                Case case_select = "pm"

                    xlWt.Cells(2, p_code).Value = tbl.ListColumns(14).DataBodyRange.Cells(Nplace, 1).Value

                    xlWt.Cells(2, "N").Value = "pm"

                Case case_select Like "*_d"

                    If Left(case_select, Len(case_select) - 2) = "pm" Then

                        xlWt.Cells(2, p_code).Value = (tbl.ListColumns(14).DataBodyRange.Cells(Nplace, 1).Value * 100 / 107) - tbl.ListColumns(15).DataBodyRange.Cells(Nplace, 1).Value

                    Else

                        xlWt.Cells(2, p_code).Value = tbl.ListColumns(14).DataBodyRange.Cells(Nplace, 1).Value

                        xlWt.Cells(2, "W").Value = tbl.ListColumns(15).DataBodyRange.Cells(Nplace, 1).Value

                    End If

                    xlWt.Cells(2, "X").Value = "Direct"

                    xlWt.Cells(2, "N").Value = Left(case_select, Len(case_select) - 2)

                                    

                Case case_select Like "*_a"

                    If Left(case_select, Len(case_select) - 2) = "pm" Then

                        xlWt.Cells(2, p_code).Value = tbl.ListColumns(14).DataBodyRange.Cells(Nplace, 1).Value - tbl.ListColumns(15).DataBodyRange.Cells(Nplace, 1).Value

                    Else

                        xlWt.Cells(2, p_code).Value = (tbl.ListColumns(14).DataBodyRange.Cells(Nplace, 1).Value * 107 / 100)

                        xlWt.Cells(2, "W").Value = tbl.ListColumns(15).DataBodyRange.Cells(Nplace, 1).Value

                    End If

                    xlWt.Cells(2, "X").Value = "Auction"

                    xlWt.Cells(2, "N").Value = Left(case_select, Len(case_select) - 2)

                                    

                Case case_select = "dis"

                    Dim numberAfterSlash As Integer

                    Dim slashPosition As Integer

                                    

                    slashPosition = InStr(tbl.ListColumns(18).DataBodyRange.Cells(Nplace + 1, 1).Value, "/")

                    numberAfterSlash = Mid(tbl.ListColumns(18).DataBodyRange.Cells(Nplace + 1, 1).Value, slashPosition + 1)

                    

                    total_hc_receive = 0

                    For m = 1 To numberAfterSlash

                        total_hc_receive = total_hc_receive + tbl.ListColumns(14).DataBodyRange.Cells(Nplace + m, 1).Value

                    Next m

                    

                    xlWt.Cells(2, p_code).Value = (xlWb.Sheets("Payment Term History").Cells(7, "CI").Value - total_hc_receive) / xlWs.Cells(7, "CI").Value

                    tbl.ListColumns(14).DataBodyRange.Cells(Nplace, 1).Value = xlWt.Cells(2, p_code).Value

                    xlWt.Cells(2, "N").Value = "dis"

                                    

                Case Else

                    xlWt.Cells(2, p_code).Value = tbl.ListColumns(14).DataBodyRange.Cells(Nplace, 1).Value

                    xlWt.Cells(2, "N").Value = tbl.ListColumns(12).DataBodyRange.Cells(Nplace, 1).Value

            End Select

        End If

    End If

    

    xlWt.Cells(2, "Q").Value = CDate(Format(tbl.ListColumns(16).DataBodyRange.Cells(Nplace, 1).Value, "mm/dd/yyyy"))

    xlWt.Cells(2, "R").Value = CDate(Format(tbl.ListColumns(17).DataBodyRange.Cells(Nplace, 1).Value, "mm/dd/yyyy"))

    

    'newFilePath & "_" & xlWs.Cells(7, "DC").Value & ".xlsx"

    results.Add True ' Result 1 as Boolean

    results.Add newFilePath & "_" & CStr(xlWs.Cells(7, "DG").Value) & ".xlsx" ' Result 2 as String"

    

    ' Save the workbook as the new file path

    xlWb.SaveAs newFilePath & "_" & CStr(xlWs.Cells(7, "DG").Value) & ".xlsx"

    xlWb.Close SaveChanges:=True

    'xlMain.Close SaveChanges:=False

    

    Set UpdateExcelFile = results

End Function

Private Function FindExcelFile(ByVal folderPath As String, ByVal searchText As String) As String

    Dim fileName As String

    Dim filePath As String



    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    

    ' Look for files with the search text in the folder

    fileName = Dir(folderPath & "*" & searchText & "*.xlsx")



    ' Loop through each .xlsx file in the folder

    Do While fileName <> ""

    

        ' Construct the full file path

        filePath = folderPath & fileName

        

        ' Split the filename on the delimiter (e.g., underscore or hyphen)

        Dim parts As Variant

        parts = Split(fileName, "_") ' Change "_" to "-" or your specific delimiter as necessary



        Dim part As Variant

        For Each part In parts

            ' Check if the part matches the searchText, also remove file extension before comparison

            If StrComp(Replace(part, ".xlsx", ""), searchText, vbTextCompare) = 0 Then

                FindExcelFile = folderPath & fileName

                Exit Function

            End If

        Next part



        ' Get the next file in the directory.

        fileName = Dir()

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

    

    If Not targetWs.UsedRange.Find(What:=core_id, LookAt:=xlWhole, MatchCase:=True) Is Nothing Then

                

        lr = targetWs.Range("AO18").Value

        For l = 6 To lr + 6

            If targetWs.Cells(l, "B").Value = core_id Then

                If targetWs.Cells(l, "S").Value = vbNullString Then

                    

                    lastrow = 6 + xlWs.Cells(3, "K").Value

                    

                    Dim cellValue As Variant

                    cellValue = targetWs.Cells(l, 31).Value



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

                    

                    If targetWs.Cells(l, 16).Value < xlWs.Cells(lastrow, "Q").Value Then

                        targetWs.Cells(l, 16).Value = xlWs.Cells(lastrow, "Q").Value

                        targetWs.Cells(l, 17).Value = xlWs.Cells(lastrow, "R").Value

                    End If

        

                    targetWs.Cells(l, 19).Value = xlWs.Cells(lastrow, "BD").Value

                    targetWs.Cells(l, 20).Value = xlWs.Cells(lastrow, "BG").Value

                    targetWs.Cells(l, 21).Value = xlWs.Cells(lastrow, "BJ").Value

                    targetWs.Cells(l, 22).Value = xlWs.Cells(lastrow, "BM").Value

                    targetWs.Cells(l, 23).Value = xlWs.Cells(lastrow, "BQ").Value

                    targetWs.Cells(l, 24).Value = xlWs.Cells(lastrow, "BT").Value

                    targetWs.Cells(l, 25).Value = xlWs.Cells(lastrow, "BW").Value

                    targetWs.Cells(l, 26).Value = xlWs.Cells(lastrow, "BZ").Value

                    targetWs.Cells(l, 27).Value = xlWs.Cells(lastrow, "CC").Value

                    targetWs.Cells(l, 28).Value = xlWs.Cells(lastrow, "CF").Value

                    targetWs.Cells(l, 29).Value = xlWs.Cells(lastrow, "CG").Value

                    targetWs.Cells(l, 30).Value = xlWs.Cells(lastrow, "CH").Value

                    targetWs.Cells(l, 31).Value = xlWs.Cells(lastrow, "CI").Value 'check error return code hold

                    targetWs.Cells(l, 32).Value = xlWs.Cells(lastrow, "CL").Value

                    

                    If targetWs.Cells(l, 12).Value Like "sold_*" And Not targetWs.Cells(l, 7).Value Like "*_closed*" Then

                        If Not targetWs.Cells(l, 7).Value Like "*_loss_after*" Then

                            targetWs.Cells(l, 7).Value = targetWs.Cells(l, 7).Value & "_loss_after"

                            Set xStat = xlApp.Workbooks.Open(ThisWorkbook.Sheets("Hold_Cutting").Range("AR13").Value)

                            Set xStatS = xStat.Sheets("payment_mode_status")

                            

                            lastStatRow = xStatS.Cells(Rows.Count, 1).End(xlUp).Row + 1

                            

                            xStatS.Cells(lastStatRow, 1).Value = targetWs.Cells(l, 2).Value

                            xStatS.Cells(lastStatRow, 3).Value = targetWs.Cells(l, 7).Value

                            xStatS.Cells(lastStatRow, 4).Value = targetWs.Cells(l, 8).Value

                            xStatS.Cells(lastStatRow, 5).Value = CDate(Format(ThisWorkbook.Sheets("Hold_Cutting").Range("AO11").Value, "dd/mm/yyyy"))

                            

                            If xStatS.Cells(lastStatRow, 7).Value = 2 Then

                                xStatS.Rows(lastStatRow).EntireRow.Delete

                            End If

                            

                            xStat.Close SaveChanges:=True

                        End If

                    End If

                    

                    If targetWs.Cells(l, 31).Value < 2 Then

                        If Not targetWs.Cells(l, 7).Value Like "*_closed*" Then

                            targetWs.Cells(l, 7).Value = targetWs.Cells(l, 7).Value & "_closed"

                            Set xStat = xlApp.Workbooks.Open(ThisWorkbook.Sheets("Hold_Cutting").Range("AR13").Value)

                            Set xStatS = xStat.Sheets("payment_mode_status")

                            

                            lastStatRow = xStatS.Cells(Rows.Count, 1).End(xlUp).Row + 1

                            

                            xStatS.Cells(lastStatRow, 1).Value = targetWs.Cells(l, 2).Value

                            xStatS.Cells(lastStatRow, 3).Value = targetWs.Cells(l, 7).Value

                            xStatS.Cells(lastStatRow, 4).Value = targetWs.Cells(l, 8).Value

                            xStatS.Cells(lastStatRow, 5).Value = CDate(Format(ThisWorkbook.Sheets("Hold_Cutting").Range("AO11").Value, "dd/mm/yyyy"))

                                                        

                            If xStatS.Cells(lastStatRow, 7).Value = 2 Then

                                xStatS.Rows(lastStatRow).EntireRow.Delete

                            End If

                            

                            xStat.Close SaveChanges:=True

                        End If

                    End If



                    targetWs.Cells(l, 33).Value = xlWs.Cells(3, "F").Value

                    targetWs.Cells(l, 34).Value = xlWs.Cells(3, "T").Value

                    targetWs.Cells(l, 35).Value = xlWs.Cells(3, "U").Value

                    targetWs.Cells(l, 36).Value = xlWs.Cells(3, "V").Value

                    targetWs.Cells(l, 37).Value = xlWs.Cells(3, "AE").Value

                    

                    If (lastrow > 7) And (xlWs.Cells(lastrow, "CI").Value < 0) Then

                        return_code.Result = "Eq_pm"

                        RevealPayment = return_code

                        Exit Function

                    End If

                    

                    targetWs.Cells(l, 12).Interior.Color = RGB(0, 255, 0)

                    Exit For

                End If

            End If

        Next l

    End If

    

    ThisWorkbook.Save

    xlApp.Calculate

    xlWb.Close SaveChanges:=False

    'xlMain.Close SaveChanges:=False

    

    return_code.Result = "N_pm"

    RevealPayment = return_code

    

    

End Function



