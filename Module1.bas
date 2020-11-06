Public dictSent As New Scripting.Dictionary
Public dictSub As New Scripting.Dictionary
Sub Inbox_Compare()
    Dim duplicateFolder As Outlook.Folder
    Dim sharedInbox As Outlook.Folder
    Dim Z_diffFolder As Outlook.Folder
    Dim dMail As Object
    Dim m As Long
    Dim ns As Outlook.NameSpace

    Set ns = Application.GetNamespace("MAPI")
    
    For Each oaccount In Application.Session.Accounts
        If oaccount = "5555@mail.com" Then
            Set myFolder = ns.GetDefaultFolder(olFolderInbox)
            Set duplicateFolder = myFolder.Folders("Duplicate")
        ElseIf oaccount = "5555@mail.com" Then
            Set myFolder = ns.Folders("5555@mail.com").Folders("Inbox")
            Set sharedInboxFolder = myFolder
            Set Z_diffFolder = ns.Folders("5555@mail.com").Folders("Z_Diff")
        End If
    Next
    
    If duplicateFolder.Items.Count = 0 Then
        response = MsgBox("No mail items in duplicate folder.", vbOKOnly)
        Exit Sub
    End If
    
    ' flag all unique mail items in sharedInbox
    processF sharedInboxFolder

    ' loop through duplicateFolder and compare - move mail items that are not in sharedInbox to Z_Diff folder, otherwise delete
    For m = duplicateFolder.Items.Count To 1 Step -1
        Set dMail = duplicateFolder.Items(m)
        If Not (dictSent.Exists(dMail.SentOn)) And Not (dictSub.Exists(dMail.Subject)) Then
            dMail.Move Z_diffFolder
            dictSent(dMail.SentOn) = True
            dictSub(dMail.Subject) = True
        Else
            dMail.Delete
        End If
    Next m
        
    Set dictSent = Nothing
    Set dictSub = Nothing
End Sub

Private Sub processF(ByVal parentF As Outlook.MAPIFolder)
    Dim currF As Outlook.MAPIFolder
    Dim mailItem As Object
    
    For Each mailItem In parentF.Items
        dictSent(mailItem.SentOn) = True
        dictSub(mailItem.Subject) = True
    Next
    
    If (parentF.Folders.Count > 0) Then
        For Each currF In parentF.Folders
            processF currF
        Next
    End If
End Sub