Attribute VB_Name = "Module1"
Sub Rectangle1_Click()

    Dim RunForm As Object ' Change to the actual UserForm class name
    
    ' Initialize the UserForm
    Set RunForm = UserForm1 ' Replace with your UserForm's actual name
    
    RunForm.Width = 350 ' Set the width as per your requirement
    RunForm.Height = 414 ' Set the height as per your requirement
    
    ' Show the UserForm
    RunForm.Show vbModeless
    
End Sub

