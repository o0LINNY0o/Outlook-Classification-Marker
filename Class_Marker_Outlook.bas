Attribute VB_Name = "Module1"
Sub CLASS_MARKING()
    Dim objMail As Outlook.mailItem
    Dim strSensitiveText As String
    
    If Application.ActiveInspector Is Nothing Then
        MsgBox "Please open an email message first.", vbExclamation
        Exit Sub
    End If
    
    Set objMail = Application.ActiveInspector.CurrentItem
    
    If objMail.Class = olMail Then
        ' Modify subject
        objMail.Subject = objMail.Subject & " [SEC=OFFICIAL:SENSITIVE]"
        
        ' Add sensitive information text in red and centered at the top of the email body
        strSensitiveText = "<div style='text-align: center; color: red;'>" & _
                           "<p>This email contains <strong>OFFICIAL: Sensitive</strong> information. " & _
                           "This information must be stored, shared and destroyed in accordance with the DSPF and/or " & _
                           "the BDA Security Manual as appropriate.</p></div>"
        
        ' Check if the email is in HTML format
        If objMail.BodyFormat <> olFormatHTML Then
            objMail.BodyFormat = olFormatHTML
        End If
        
        ' Add the sensitive text at the beginning of the email body
        objMail.HTMLBody = strSensitiveText & objMail.HTMLBody
        
        objMail.Save
    Else
        MsgBox "This is not an email message.", vbExclamation
    End If
    
    Set objMail = Nothing
End Sub

