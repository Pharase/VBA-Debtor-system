Attribute VB_Name = "PasswordForm"
Attribute VB_Base = "0{2CB80F2F-A271-424F-9B54-9FF763C8D850}{1B3436E2-6D73-4F25-8C25-08D56655404D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private passwordEntered As Boolean

Private Sub CommandButton1_Click()
    passwordEntered = True
    Me.Hide
End Sub
Private Sub CommandButton2_Click()
    Me.Hide
End Sub

Private Sub PasswordForm_Initialize()
    passwordEntered = False
    Me.TextBox1.PasswordChar = "*"
End Sub

Public Function GetPassword() As String
    Me.Show
    Do While Me.Visible
        DoEvents
    Loop
    If passwordEntered Then
        GetPassword = Me.TextBox1.Text
    Else
        GetPassword = vbNullString
    End If
End Function

