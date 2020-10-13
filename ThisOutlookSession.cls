VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public WithEvents objInbox As Outlook.Folder
Attribute objInbox.VB_VarHelpID = -1
Public WithEvents objInboxItems As Outlook.Items
Attribute objInboxItems.VB_VarHelpID = -1

Private Sub Application_Startup()
    Set objInbox = Outlook.Application.Session.GetDefaultFolder(olFolderInbox).Folders("Beachbody Analytics")
    Set objInboxItems = objInbox.Items
End Sub

Private Sub objInboxItems_ItemAdd(ByVal Item As Object)
    Dim objMail As Outlook.MailItem
    Dim objForward As Outlook.MailItem
    'MsgBox "Message subject: " & Item.Subject & vbCrLf & "Message sender: " & Item.SenderName & " (" & Item.SenderEmailAddress & ")"
    
    If TypeOf Item Is MailItem Then
       Set objMail = Item
       Dim strHTMLBody As String
       'objMail.Display
       'strHTMLBody = objMail.HTMLBody
       'MsgBox "Message Sender: " & objMail.Sender & vbCrLf & "Message SenderEmailAddress: " & objMail.SenderEmailAddress & vbCrLf & "Message SentOn: " & objMail.SentOn & vbCrLf & "Message SenderEmailType: " & objMail.SenderEmailType & vbCrLf & "Message SendUsingAccount: " & objMail.SendUsingAccount
       
       'If it is a specific new email
       'check the sender and subject
       'forward the email with original content to different email address
       If (objMail.SendUsingAccount = "analytics@beachbody.com") And (objMail.Subject = "Daily BOD Streaming Detail") Then
            'Find email addresses in the body via Regular Expression
            'Set objRegExp = CreateObject("vbscript.RegExp")
            'With objRegExp
            '     .Pattern = "(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|""(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\[\x01-\x09\x0b\x0c\x0e-\x7f])*"")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])"
            '     .IgnoreCase = True
            '     .Global = True
            'End With
       
            'Remove email addresses
            'If objRegExp.Test(strHTMLBody) Then
            '    Set objFoundResults = objRegExp.Execute(strHTMLBody)
            '    For i = 1 To objFoundResults.Count
            '        objMail.HTMLBody = Replace(objMail.HTMLBody, objFoundResults.Item(i - 1).Value, "")
            '    Next
            ' End If
               
            'Remove useless characters
            'objMail.HTMLBody = Replace(objMail.HTMLBody, "[mailto:]", "")
            'objMail.HTMLBody = Replace(objMail.HTMLBody, "()", "")
            
            Set objForward = objMail.Forward
            'Customize the forward subject, body and recipients
            With objForward
                 .Subject = "Daily BOD Streaming Detail"
                 .HTMLBody = objMail.HTMLBody
                 .Recipients.Add ("abhishek.chakraborthy@cesltd.com")
                 .Recipients.ResolveAll
                 '.Importance = olImportanceHigh
                 .Send
            End With
       End If
    End If
End Sub
