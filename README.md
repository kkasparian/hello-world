This is a VBA code that imports information about the emails stored in outlook, specifically in folders called "new" and "AQA new".


Sub OUTLOOK_IMPORT_EMAILBODY()

Dim O As Outlook.Application
Set O = New Outlook.Application
Dim ONS As Outlook.Namespace
Set ONS = O.GetNamespace("MAPI")
Dim MYFOL As Outlook.Folder
Set MYFOL = ONS.GetDefaultFolder(olFolderInbox).Folders("new")
Dim OMAIL As Outlook.MailItem
Set OMAIL = O.CreateItem(olMailItem)
Dim R As Long
R = 570
For Each OMAIL In MYFOL.Items
Cells(R, 2).Value = OMAIL.ReceivedTime
Cells(R, 8).Value = OMAIL.Body
Cells(R, 3).Value = OMAIL.CreationTime
R = R + 1
Next OMAIL
End Sub


Sub AQA()

Dim O As Outlook.Application
Set O = New Outlook.Application
Dim ONS As Outlook.Namespace
Set ONS = O.GetNamespace("MAPI")
Dim MYFOL As Outlook.Folder
Set MYFOL = ONS.GetDefaultFolder(olFolderInbox).Folders("AQA new")
Dim OMAIL As Outlook.MailItem
Set OMAIL = O.CreateItem(olMailItem)
Dim R As Long
R = 570
For Each OMAIL In MYFOL.Items
Cells(R, 2).Value = OMAIL.ReceivedTime
Cells(R, 8).Value = OMAIL.TaskSubject
Cells(R, 3).Value = OMAIL.CreationTime
R = R + 1
Next OMAIL
End Sub
