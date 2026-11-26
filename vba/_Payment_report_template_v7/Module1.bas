Attribute VB_Name = "Module1"
Sub OpenLinkedFiles()
    Dim wb As Workbook
    Dim filePathB As String
    Dim filePathC As String
    
    ' Get the file paths from cells V2 and V3
    filePathB = ThisWorkbook.Sheets("Payment_Term_History").Range("Z2").Value
    filePathC = ThisWorkbook.Sheets("Payment_Term_History").Range("Z3").Value
    
    If (Len(filePathB) > 0) And (Len(filePathC) > 0) Then
        ' Open workbook B
        Set wb = Workbooks.Open(filePathB)
        ' Open workbook C
        Set wb = Workbooks.Open(filePathC)
    Else
        MsgBox "No file _Data_CutOff  or _Data_Status in card_data"
    End If

End Sub

