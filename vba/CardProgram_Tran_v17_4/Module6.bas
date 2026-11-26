Attribute VB_Name = "Module6"
Sub AssignOA()
    Dim objShell As Object
    Dim User_name As String
    Set objShell = VBA.CreateObject("Wscript.shell")
    User_name = Environ("USERNAME")
    
    Dim PythonExePath As String
    PythonExePath = "C:\Users\" & User_name & "\miniconda3\python.exe"
    
    Dim PythonScriptPath As String
    PythonScriptPath = "C:\card_data\processing\program\Update_assign.py"
    
    objShell.Run "C:\Users\" & User_name & "\miniconda3\python.exe install pywin32"
    objShell.Run "C:\Users\" & User_name & "\miniconda3\python.exe install pandas"
    objShell.Run "C:\Users\" & User_name & "\miniconda3\python.exe install msoffcrypto-tool"
    objShell.Run PythonExePath & " " & PythonScriptPath, 0
    
End Sub
