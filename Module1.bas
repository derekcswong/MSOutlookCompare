Attribute VB_Name = "Module1"
Sub Inbox_Compare()
    Dim currentDLFolder As Outlook.Folder
    Dim diffFolder As Outlook.Folder
    Dim sharedInboxFolder As Outlook.Folder
    Dim sharedProcessedFolder As Outlook.Folder
    Dim sMail As Object
    Dim dMail As Object
    Dim MailC As Object
    Dim dictSent As New Scripting.Dictionary, m As Long
    Dim dictSub As New Scripting.Dictionary
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    
    For Each oaccount In Application.Session.Accounts
'        Debug.Print oaccount
        If oaccount = "first account" Then
            Set myFolder = ns.GetDefaultFolder(olFolderInbox)
            Set currentDLFolder = myFolder.Folders("Duplicate")
            Set diffFolder = myFolder.Folders("Diff")
        End If
        
        If oaccount = "second account" Then
            Set myFolder = ns.Folders("second account").Folders("Inbox")
            Set sharedInboxFolder = myFolder
            Set sharedProcessedFolder = myFolder.Folders("Processed")
        End If
    Next
    
    ' flag all unique mail items in sharedInboxFolder
    For Each dMail In sharedInboxFolder.Items
        dictSent(dMail.SentOn) = True
        dictSub(dMail.Subject) = True
    Next
    ' loop through SCP inbox and compare - copy any mail items not in SCP inbox to Diff
    For m = currentDLFolder.Items.Count To 1 Step -1
        Set sMail = currentDLFolder.Items(m)
        If Not (dictSent.Exists(sMail.SentOn)) And Not (dictSub.Exists(sMail.Subject)) Then
            Set MailC = sMail.Copy
            MailC.Move diffFolder
            dictSent(sMail.SentOn) = True
            dictSub(sMail.Subject) = True
        End If
    Next m
    
    ' move compared mail to processed folder
    For Each dMail In sharedInboxFolder.Items
        If (dictSent.Exists(dMail.SentOn)) And (dictSub.Exists(dMail.Subject)) Then
            dMail.Move sharedProcessedFolder
        End If
    Next
    
    
    Set dictSent = Nothing
    Set dictSub = Nothing
    
'    Debug.Print currentDLFolder.Name
'    Debug.Print diffFolder.Name
'    Debug.Print sharedInboxFolder.Name
'    Debug.Print sharedProcessedFolder.Name
End Sub