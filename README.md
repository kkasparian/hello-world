This is a VBA code that imports information about the emails stored in outlook, specifically in folders called "new" and "AQA new".

Sub OUTLOOK_IMPORT()

Dim O As Outlook.Application
Set O = New Outlook.Application

Dim ONS As Outlook.Namespace
Set ONS = O.GetNamespace("MAPI")

Dim MYFOL As Outlook.Folder
Set MYFOL = ONS.GetDefaultFolder(olFolderInbox).Folders("alerts")

' "411", "911", "AQA Alert", and "Morning" are sub-folders of the folder "alerts", which is a sub-folder of "Inbox"
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
' "NextRow" allows you to paste to the next available row in excel
NextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1

' If the email is unread, then it pastes the date , body, and time the email was sent to the corresponding columns

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

'This inserts a row above thr 3rd row in another worksheet called "Morning Health Check Dashboard" each time an email is being imported from the folder
Worksheets("Morning Health Check Dashboard").Rows(3).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
Worksheets("Morning Health Check Dashboard").Cells(3, 1).Value = OMAIL.ReceivedTime
NextRow = NextRow + 1
Next OMAIL
End If



End Sub
