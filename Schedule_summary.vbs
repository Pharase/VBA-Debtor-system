'Input source and destination folder paths
SourceFolder = "Z:\CutOff\3.Statement Card"
DestinationFolder = "C:\Pam_card\summary\file"

'Create a FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

'Check if source folder exists
If fso.FolderExists(SourceFolder) Then
    'Check if destination folder exists, if not, create it
    If Not fso.FolderExists(DestinationFolder) Then
        fso.CreateFolder (DestinationFolder)
    End If
    
    'Copy specific files from source folder to destination folder
    Set SourceFiles = fso.GetFolder(SourceFolder).Files
    For Each file In SourceFiles
        If InStr(file.Name, "StatementCard_") > 0 Then
            fso.CopyFile file.Path, DestinationFolder & "\", True
        End If
    Next
    MsgBox "Specific files copied from " & SourceFolder & " to " & DestinationFolder, vbInformation
Else
    MsgBox "Source folder does not exist.", vbExclamation
End If

'Input Excel File's Full Path
  ExcelFilePath = "C:\Pam_card\processing\program\CardProgram_Tran_v15.xlsm"

'Input Module/Macro name within the Excel File
  MacroPath = "Module5.SummaryCard"

'Create an instance of Excel
  Set ExcelApp = CreateObject("Excel.Application")

'Do you want this Excel instance to be visible?
  ExcelApp.Visible = True  'or "False"

'Prevent any App Launch Alerts (ie Update External Links)
  ExcelApp.DisplayAlerts = False

'Open Excel File
  Set wb = ExcelApp.Workbooks.Open(ExcelFilePath)

'Execute Macro Code
  ExcelApp.Run MacroPath

'Save Excel File (if applicable)
  wb.Save

'Reset Display Alerts Before Closing
  ExcelApp.DisplayAlerts = True

'Close Excel File
  wb.Close

'End instance of Excel
  ExcelApp.Quit

'Leaves an onscreen message!
  MsgBox "Automated Summary Transaction successfully ran at " & TimeValue(Now), vbInformation