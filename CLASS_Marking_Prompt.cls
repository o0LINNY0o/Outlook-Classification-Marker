VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error Resume Next
    
    ' Check if the subject line includes "[SEC="
    If InStr(1, Item.Subject, "[SEC=", vbTextCompare) = 0 Then
        Prompt$ = "The Subject line does not include a SECURITY CLASSIFICATION" & vbNewLine & _
                  "Are you sure you want to send it without a classification?"
        If MsgBox(Prompt$, vbYesNo + vbQuestion + vbMsgBoxSetForeground, "Security Classification Check") = vbNo Then
            Cancel = True
        End If
    End If
End Sub


