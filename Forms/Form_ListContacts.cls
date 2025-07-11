VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ListContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim sGrantType As String

Private Sub cmdListContacts_Click()
    CurrentDb.Execute "DELETE FROM TempContactsTable"
    If Nz(Me.fromMailbox, "") = "" And sGrantType = "client_credentials" Then
        MsgBox "From whose mailbox is mandatory and must be a valid mailbox in this client."
        Me.fromMailbox.SetFocus
    Else
        'Second parameter is optional, if filled in use the name of a folder,
        'for example my contacts are in TempContacts so I use that
        ListOutlookContacts Graph.ListContacts(Nz(Me.fromMailbox, ""), Nz(Me.folderName, ""))
        Me.lstContacts.Requery
        Me.lstContacts.Visible = True
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    CurrentDb.Execute "DELETE FROM TempContactsTable"
    sGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If sGrantType = "client_credentials" Then
        Me.fromMailbox.Visible = True
    Else
        Me.fromMailbox.Visible = False
    End If
End Sub

Private Sub ListOutlookContacts(Response As WebResponse)
    If Response.StatusCode = WebStatusCode.OK Then
        Dim ContactInfo As Dictionary
        Dim EmailColl As Collection
        Dim EmailAddress As Dictionary
        Dim rs As Recordset
                
        If IsEmpty(Response.Data("value")) Then
            MsgBox "No contacts for this folder in this tenant."
        Else
            Set rs = CurrentDb.OpenRecordset("TempContactsTable")
            For Each ContactInfo In Response.Data("value")
                Set EmailColl = ContactInfo("emailAddresses")
                For Each EmailAddress In EmailColl
                    If EmailAddress("address") <> "" Then
                        rs.AddNew
                        rs.Fields("id").Value = ContactInfo("id")
                        rs.Fields("createdDateTime").Value = ContactInfo("createdDateTime")
                        rs.Fields("lastModifiedDateTime").Value = ContactInfo("lastModifiedDateTime")
                        rs.Fields("displayName").Value = ContactInfo("displayName")
                        rs.Fields("givenName").Value = ContactInfo("givenName")
                        rs.Fields("surname").Value = ContactInfo("surname")
                        rs.Fields("emailAddress").Value = EmailAddress("address")
                        rs.Update
                    End If
                Next EmailAddress
            Next ContactInfo
            rs.Close
        End If
        Set rs = Nothing
    Else
        MsgBox "Error " & Response.StatusCode & ": " & Response.Content
    End If
End Sub
