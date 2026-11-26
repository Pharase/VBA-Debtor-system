Attribute VB_Name = "Module4"

Sub SendEmail()

    Dim OutlookApp As Object

    Dim OutlookMail As Object

    Dim prompt As VbMsgBoxResult

    Dim mySyncObjects As Object

    Dim syc As Object

    

    ' Create a new instance of Outlook

    Set OutlookApp = CreateObject("Outlook.Application")

    

    ' Create a new email item

    Set OutlookMail = OutlookApp.CreateItem(0) ' 0 indicates MailItem

    

    With OutlookMail

        .To = "recipient1@example.com;recipient3@example.com;recipient4@example.com;recipient2@example.com;"

        .CC = "" ' Change to CC recipient's email address (optional)

        .BCC = "" ' Change to BCC recipient's email address (optional)

        .Subject = "Data Daily " & ThisWorkbook.Sheets("Hold_Cutting").Range("AO11").Value

        .Body = ThisWorkbook.Sheets("Hold_Cutting").Range("AN25").Value & vbCrLf & vbCrLf & ThisWorkbook.Sheets("Hold_Cutting").Range("AN26").Value & vbCrLf & vbCrLf & ThisWorkbook.Sheets("Hold_Cutting").Range("AN27").Value

        .Attachments.Add ThisWorkbook.Sheets("Hold_Cutting").Range("AP13").Value ' Attach the file

        ' Display the email for review

        .Display

    End With

    

    ' Prompt the user to confirm sending the email after reviewing it

    prompt = MsgBox("Do you want to send the email?", vbYesNo + vbQuestion, "Send Email")

    

    If prompt = vbYes Then

        ' Send the email

        OutlookMail.Send

        

        

        ' Ensure the email is sent immediately by synchronizing

        Application.Wait (Now + TimeValue("0:00:05"))

        On Error Resume Next



        Set OutlookApp = CreateObject(Class:="Outlook.Application")

        

        On Error GoTo 0

        

        Set mySyncObjects = OutlookApp.GetNamespace("MAPI").SyncObjects

        Set syc = mySyncObjects.Item(1)

        syc.Start

        Application.Wait (Now + TimeValue("0:00:15"))

        MsgBox "Email sent successfully!"

        

    Else

        MsgBox "Email sending canceled."

        Set OutlookMail = Nothing

        Set OutlookApp = Nothing

    End If

    

    ' Clean up

    Set OutlookMail = Nothing

    Set OutlookApp = Nothing

    Set mySyncObjects = Nothing

    Set syc = Nothing

End Sub



