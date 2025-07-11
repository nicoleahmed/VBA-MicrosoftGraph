VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SendEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim sGrantType As String

Private Sub cmdBrowse_Click()
    Dim Fd As Variant
    Set Fd = Application.FileDialog(3) 'msoFileDialogFilePicker
    Fd.Show
    Me.sAttachment = Fd.SelectedItems(1)
End Sub

'File attachment size limited to under 3 MB.  To upload larger files you would need to create a session and load in
'groups, see https://learn.microsoft.com/en-us/graph/api/attachment-createuploadsession?view=graph-rest-1.0&tabs=http

Private Sub cmdCreateDraft_Click()
    If Nz(Me.sFrom, "") = "" And sGrantType = "client_credentials" Then
        MsgBox "From address is mandatory and must be a valid mailbox in this client."
        Me.sFrom.SetFocus
    Else
        CreateDraftEmail Graph.CreateDraftMessage(Nz(Me.sFrom, ""), Nz(Me.sSubject, ""), "HTML", Nz(Me.sBody, ""), Nz(Me.sTo, ""), Nz(Me.sCC, ""), Nz(Me.sBcc, ""), Nz(Me.sAttachment, ""))
    End If
End Sub

Private Sub CreateDraftEmail(Response As WebResponse)
    If Response.StatusCode = WebStatusCode.Created Then
        MsgBox "Email created. View, edit if desired, and Send from your Drafts folder!"
'       ? Response.Data("id")  This value could be used to do something else with the draft
    Else
        MsgBox "Error " & Response.StatusCode & ": " & Response.Content
    End If
End Sub

Private Sub cmdSend_Click()
    If Nz(Me.sFrom, "") = "" And sGrantType = "client_credentials" Then
        MsgBox "From address is mandatory and must be a valid mailbox in this client."
        Me.sFrom.SetFocus
    Else
        SendEmail Graph.GraphSendMail(Nz(Me.sFrom, ""), Nz(Me.sSubject, ""), "HTML", Nz(Me.sBody, ""), Nz(Me.sTo, ""), Nz(Me.sCC, ""), Nz(Me.sBcc, ""), Nz(Me.sAttachment, ""))
    End If
End Sub

Private Sub SendEmail(Response As WebResponse)
    If Response.StatusCode = WebStatusCode.Accepted Then
        MsgBox "Email Sent!"
    Else
        MsgBox "Error " & Response.StatusCode & ": " & Response.Content
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    sGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If sGrantType = "client_credentials" Then
        Me.sFrom.Visible = True
    Else
        Me.sFrom.Visible = False
    End If
End Sub
