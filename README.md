This is a VBA code that imports information about the emails stored in outlook, specifically in folders called "new" and "AQA new".

Sub OUTLOOK_IMPORT()

Dim O As Outlook.Application
Set O = New Outlook.Application

Dim ONS As Outlook.Namespace
Set ONS = O.GetNamespace("MAPI")

Dim MYFOL As Outlook.Folder
Set MYFOL = ONS.GetDefaultFolder(olFolderInbox).Folders("alerts")

Dim SubFolder411 As MAPIFolder
Set SubFolder411 = MYFOL.Folders("411")

Dim SubFolder911 As MAPIFolder
Set SubFolder911 = MYFOL.Folders("911")

Dim SubFolderalert As MAPIFolder
Set SubFolderalert = MYFOL.Folders("CBT alert")

Dim SubFolderAQA As MAPIFolder
Set SubFolderAQA = MYFOL.Folders("AQA Alert")

Dim SubFolderHC As MAPIFolder
Set SubFolderHC = MYFOL.Folders("Morning")


Dim OMAIL As Outlook.MailItem
Set OMAIL = O.CreateItem(olMailItem)

Dim NextRow As Long
With ActiveSheet
NextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1

If SubFolder411.Items.Restrict("[UnRead]=True").Count > 0 Then
    
        For Each OMAIL In SubFolder411.Items.Restrict("[UnRead]=True")

Cells(NextRow, 2).Value = OMAIL.ReceivedTime
Cells(NextRow, 8).Value = OMAIL.Body
Cells(NextRow, 3).Value = OMAIL.CreationTime
NextRow = NextRow + 1
Next OMAIL

End If

If SubFolder911.Items.Restrict("[UnRead]=True").Count > 0 Then
    
        For Each OMAIL In SubFolder911.Items.Restrict("[UnRead]=True")

Cells(NextRow, 2).Value = OMAIL.ReceivedTime
Cells(NextRow, 8).Value = OMAIL.Body
Cells(NextRow, 3).Value = OMAIL.CreationTime
NextRow = NextRow + 1
Next OMAIL

End If



If SubFolderalert.Items.Restrict("[UnRead]=True").Count > 0 Then
    
        For Each OMAIL In SubFolderalert.Items.Restrict("[UnRead]=True")

Cells(NextRow, 2).Value = OMAIL.ReceivedTime
Cells(NextRow, 8).Value = OMAIL.Body
Cells(NextRow, 3).Value = OMAIL.CreationTime
NextRow = NextRow + 1
Next OMAIL

End If


If SubFolderAQA.Items.Restrict("[UnRead]=True").Count > 0 Then
    
        For Each OMAIL In SubFolderAQA.Items.Restrict("[UnRead]=True")

Cells(NextRow, 2).Value = OMAIL.ReceivedTime
Cells(NextRow, 8).Value = OMAIL.TaskSubject
Cells(NextRow, 3).Value = OMAIL.CreationTime
NextRow = NextRow + 1
Next OMAIL

End If

If SubFolderHC.Items.Restrict("[UnRead]=True").Count > 0 Then
    
        For Each OMAIL In SubFolderHC.Items.Restrict("[UnRead]=True")

Cells(NextRow, 2).Value = OMAIL.ReceivedTime
Cells(NextRow, 8).Value = OMAIL.TaskSubject
Cells(NextRow, 3).Value = OMAIL.CreationTime

Worksheets("Morning Health Check Dashboard").Rows(3).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
Worksheets("Morning Health Check Dashboard").Cells(3, 1).Value = OMAIL.ReceivedTime
NextRow = NextRow + 1
Next OMAIL
End If



End Sub
