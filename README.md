This is a VBA code that imports information about the emails stored in outlook, specifically in folders called "new" and "AQA new".

Sub OUTLOOK_IMPORT()

Dim O As Outlook.Application
Set O = New Outlook.Application

Dim ONS As Outlook.Namespace
Set ONS = O.GetNamespace("MAPI")

Dim MYFOL As Outlook.Folder
Set MYFOL = ONS.GetDefaultFolder(olFolderInbox).Folders("new")

Dim OMAIL As Outlook.MailItem
Set OMAIL = O.CreateItem(olMailItem)

Dim NextRow As Long
With ActiveSheet
NextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1

For Each OMAIL In MYFOL.Items
Cells(NextRow, 2).Value = OMAIL.ReceivedTime
Cells(NextRow, 8).Value = OMAIL.Body
Cells(NextRow, 3).Value = OMAIL.CreationTime
NextRow = NextRow + 1
Next OMAIL
End With

Set MYFOL = ONS.GetDefaultFolder(olFolderInbox).Folders("Morning HC")


For Each OMAIL In MYFOL.Items
Cells(NextRow, 2).Value = OMAIL.ReceivedTime
Cells(NextRow, 8).Value = OMAIL.TaskSubject
Cells(NextRow, 3).Value = OMAIL.CreationTime

Worksheets("Morning Health Check Dashboard").Rows(3).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
Worksheets("Morning Health Check Dashboard").Cells(3, 1).Value = OMAIL.ReceivedTime
NextRow = NextRow + 1

Next OMAIL

Set MYFOL = ONS.GetDefaultFolder(olFolderInbox).Folders("AQA new")


For Each OMAIL In MYFOL.Items
Cells(NextRow, 2).Value = OMAIL.ReceivedTime
Cells(NextRow, 8).Value = OMAIL.TaskSubject
Cells(NextRow, 3).Value = OMAIL.CreationTime
NextRow = NextRow + 1
Next OMAIL



End Sub
