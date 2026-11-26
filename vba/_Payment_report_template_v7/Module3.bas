Attribute VB_Name = "Module3"
Sub python_payment()
    Dim objShell As Object
    Dim User_name As String
    Set objShell = VBA.CreateObject("Wscript.shell")
    User_name = Environ("USERNAME")
    
    Dim PythonExePath As String
    PythonExePath = "C:\Users\" & User_name & "\miniconda3\python.exe"
    
    Dim PythonScriptPath As String
    PythonScriptPath = "C:\card_data\processing\program\CF_payment.py"
    
    objShell.Run "C:\Users\" & User_name & "\miniconda3\python.exe install pywin32 selenium selenium pandas openpyxl webdriver-manager"
    objShell.Run PythonExePath & " " & PythonScriptPath, 0
    
End Sub
