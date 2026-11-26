Attribute VB_Name = "Module5"
Option Explicit
Sub SummaryCard()
    Dim objShell As Object
    Dim User_name As String
    Set objShell = VBA.CreateObject("Wscript.shell")
    User_name = Environ("USERNAME")
    
    Dim PythonExePath As String
    PythonExePath = "C:\Users\" & User_name & "\miniconda3\python.exe"
    
    Dim PythonScriptPath As String
    PythonScriptPath = "C:\card_data\processing\program\Summary_transaction_program-v2.py"
    
    objShell.Run "C:\Users\" & User_name & "\miniconda3\python.exe install pywin32"
    objShell.Run "C:\Users\" & User_name & "\miniconda3\python.exe install xlwings"
    objShell.Run "C:\Users\" & User_name & "\miniconda3\python.exe install tqdm"
    objShell.Run PythonExePath & " " & PythonScriptPath, 0
    
End Sub
