Private Sub Application_Startup()
    Dim olApp As Outlook.Application
    Dim objNS As Outlook.NameSpace
    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
End Sub
Private Sub Items_ItemAdd(ByVal cusItem As Object)
    Dim olApp As Outlook.Application
    Dim objNS As Outlook.NameSpace
    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    
    Dim strEntryID As String
    Dim objAddressentry As Outlook.AddressEntry
    Dim objRecipient As Outlook.Recipient
    Dim objDestFolder As Outlook.MAPIFolder
    Dim objRecs As Outlook.Recipients
    Dim i As Integer
    
    If TypeName(cusItem) = "MailItem" Then
        
        On Error GoTo ErrorHandler
        
        Set objRecs = cusItem.Recipients
        
        For i = 1 To objRecs.Count
        
            Set objRecipient = objRecs.Item(i)
        
            strEntryID = objRecipient.EntryID

            Set objAddressentry = objNS.GetAddressEntryFromID(strEntryID)
   
            If objAddressentry = "Customer Service" And TimeValue(Now()) >= TimeValue("08:00:00 AM") And TimeValue(Now()) <= TimeValue("05:00:00 PM") Then
        
            Set objDestFolder = objNS.GetDefaultFolder(olFolderInbox).Folders("Customer Service")
    
            cusItem.Move objDestFolder
            
            Exit For
        
            End If
        
        Next

    End If

ProgramExit:
  Exit Sub
  
ErrorHandler:
  MsgBox Err.Number & " - " & Err.Description & vbCrLf & "Click Ok to continue operations"
  Resume ProgramExit
    
End Sub
