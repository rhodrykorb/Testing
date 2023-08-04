VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public WithEvents mycol As Items
Attribute mycol.VB_VarHelpID = -1
Public WithEvents thisfolder As Outlook.Folder
Attribute thisfolder.VB_VarHelpID = -1



Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

End Sub

Private Sub Application_Quit()
Set mycol = Nothing
Set thisfolder = Nothing
End Sub

Private Sub Application_Startup()
On Error GoTo errhandler
Dim myns As Outlook.NameSpace, myfolder As Outlook.MAPIFolder
Set myns = Application.GetNamespace("MAPI")
Set myfolder = myns.GetDefaultFolder(olFolderCalendar)
Set thisfolder = myfolder
Set mycol = myfolder.Items
Set myfolder = Nothing
Set myns = Nothing
Exit Sub
errhandler:
If Err.Number <> 0 Then
    MsgBox Err.Description, vbOKOnly + vbInformation, "Meeting"
    Exit Sub
End If
End Sub

Private Sub mycol_ItemAdd(ByVal Item As Object)
On Error GoTo errhandler
If Item.Class = 26 Then
    Dim myappt As Outlook.AppointmentItem
    Set myappt = Item
    Dim found As Boolean
    Dim myrcp  As Outlook.Recipient, thisappt As Outlook.AppointmentItem
    If myappt.Mileage <> "" Then Exit Sub
    If myappt.MeetingStatus = olMeeting Then found = True
    If myappt.MeetingStatus = olMeetingReceived Then found = True
    If myappt.Recipients.Count > 1 Then found = True
        If found = True Then
            If myappt.Mileage <> "" Then Exit Sub
            For Each myrcp In myappt.Recipients
                    If myrcp.Address = "rodk3150@gmail.com" Then
                        Set myappt = Nothing
                        Set myrcp = Nothing
                        Set thisappt = Nothing
                        Exit Sub
                    End If
            Next
        Dim mynewAppt As Outlook.AppointmentItem
       found = False
        For Each thisappt In Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
            If thisappt.Mileage = myappt.EntryID Then
                Set mynewAppt = thisappt
                found = True
                Exit For
            End If
            If thisappt.Start = myappt.Start Then
                If thisappt.End = myappt.End Then
                    If thisappt.Duration = myappt.Duration Then
                        If thisappt.Subject <> myappt.Subject Then
                            found = True
                            Set mynewAppt = thisappt
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
        If found = False Then Set mynewAppt = Application.CreateItem(olAppointmentItem)
        
        With mynewAppt
            If found = True Then
                .AllDayEvent = myappt.AllDayEvent
                .Body = ""
                .BusyStatus = myappt.BusyStatus
                .Categories = myappt.Categories
                .Duration = myappt.Duration
                .End = myappt.End
                .Importance = myappt.Importance
                .Location = myappt.Location
                If myappt.ResponseStatus = olResponseNotResponded Or myappt.ResponseStatus = olResponseTentative Then
                    .Subject = "Work Meeting - Tentative"
                End If
                If myappt.ResponseStatus = olResponseDeclined Then
                    .Delete
                    Exit Sub
                End If
                If myappt.ResponseStatus = olResponseAccepted Then
                    .Subject = "WorkMeeting - Accepted"
                End If
               .Start = myappt.Start
                If myappt.ResponseStatus <> olResponseDeclined Then
                    .End = myappt.End
                    .Mileage = myappt.EntryID
                    '.Save
                    .Send
                    Debug.Print "Same appt.ItemAdd"
                End If
            Else
                .MeetingStatus = olMeeting
                .AllDayEvent = myappt.AllDayEvent
                .Body = ""
                .BusyStatus = myappt.BusyStatus
                .Categories = myappt.Categories
                .Duration = myappt.Duration
                .End = myappt.End
                .Importance = myappt.Importance
                .Location = myappt.Location
                .ResponseRequested = False
                If myappt.ResponseStatus = olResponseNotResponded Then
                    .Subject = "Work Meeting - Not Accepted"
                End If
                If .GetOrganizer = myappt.GetOrganizer Then
                    .Subject = "Work Meeting - My Meeting"
                End If
                If myappt.ResponseStatus = olResponseAccepted Then
                    .Subject = "WorkMeeting - Accepted"
                End If
                Dim myrec As Outlook.Recipient
                Set myrec = .Recipients.Add("rodk3150@gmail.com")
                myrec.Type = olOptional
                .Start = myappt.Start
                If myappt.ResponseStatus <> olResponseDeclined Then
                    .End = myappt.End
                    .Mileage = myappt.EntryID
                    '.Save
                    .Send
                    Debug.Print "New Message"
                End If
            End If
        End With
        Set mynewAppt = Nothing
    End If
    Set myappt = Nothing
End If
errhandler:
If Err.Number <> 0 Then
    MsgBox Err.Description, vbOKOnly + vbInformation, "Meeting"
    Exit Sub
End If
End Sub

Private Sub mycol_ItemChange(ByVal Item As Object)
On Error GoTo errhandler
If Item.Class = 26 Then
  Dim myappt As Outlook.AppointmentItem
    'Set mymeeting = Item
    Dim myrec As Outlook.Recipient
    Set myappt = Item
    On Error GoTo errhandler
    If myappt.Mileage <> "" Then Exit Sub
    Dim myrcp As Recipient
    For Each myrcp In myappt.Recipients
                    If myrcp.Address = "rodk3150@gmail.com" Then
                        Set myappt = Nothing
                        Set myrcp = Nothing
                        Set thisappt = Nothing
                        Exit Sub
                    End If
    Next
    'Set myappt = mymeeting.GetAssociatedAppointment(False)
    Dim found As Boolean
    If myappt.MeetingStatus = olMeeting Then found = True
    If myappt.Recipients.Count > 1 Then found = True
    If found = False Then Exit Sub
    Dim mysame As Outlook.AppointmentItem
    For Each mysame In Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
        If mysame.Mileage = myappt.EntryID Then
            If myappt.MeetingStatus = olMeetingCanceled Or myappt.MeetingStatus = olMeetingReceivedAndCanceled Then
                mysame.MeetingStatus = olMeetingCanceled
                'mysame.Save
                mysame.Send
                mysame.Delete
                myappt.Delete
                Exit For
            Else
                If mysame.Start <> myappt.Start Or mysame.End <> myappt.End Or mysame.Location <> myappt.Location Or mysame.Duration <> myappt.Duration Then
                    mysame.Start = myappt.Start
                    mysame.End = myappt.End
                    mysame.Duration = myappt.Duration
                    mysame.Location = myappt.Location
                    mysame.Body = ""
                    'mysame.Save
                    mysame.Subject = "Work Meeting - Not Accepted"
                    mysame.Send
                    Debug.Print "Same mileage-ItemChange"
                    Exit For
                End If
            End If
        End If
        If mysame.Start = myappt.Start And mysame.End = myappt.End And mysame.Duration = myappt.Duration Then
            If mysame.Subject <> myappt.Subject Then
                If myappt.MeetingStatus = olMeetingCanceled Or myappt.MeetingStatus = olMeetingReceivedAndCanceled Then
                    mysame.MeetingStatus = olMeetingCanceled
                    'mysame.Save
                    mysame.Send
                    mysame.Delete
                    myappt.Delete
                    Exit For
                Else
                    If mysame.Start <> myappt.Start Or mysame.End <> myappt.End Or mysame.Location <> myappt.Location Or mysame.Duration <> myappt.Duration Then
                        mysame.Start = myappt.Start
                        mysame.End = myappt.End
                        mysame.Duration = myappt.Duration
                        mysame.Location = myappt.Location
                        mysame.Body = ""
                        'mysame.Save
                        mysame.Subject = "Work Meeting - Not Accepted"
                        mysame.Send
                        Debug.Print "ItemChange"
                        Exit For
                    End If
                End If
            End If
        End If
      Next
      'mysame = Nothing
      'myappt = Nothing
End If
Exit Sub
errhandler:
If Err.Number <> 0 Then MsgBox Err.Description, vbOKOnly + vbInformation, "Meeting"
End Sub




Private Sub thisfolder_BeforeItemMove(ByVal Item As Object, ByVal MoveTo As MAPIFolder, Cancel As Boolean)
On Error GoTo errhandler
If InStr(1, MoveTo.Name, "Delete", vbTextCompare) = 0 Then Exit Sub
If Item.Class = 26 Then
     Dim myappt As Outlook.AppointmentItem
     Dim mymeeting As Outlook.MeetingItem
    'Set mymeeting = Item
    Dim myrec As Outlook.Recipient
    Set myappt = Item
    If myappt.Mileage <> "" Then Exit Sub
    Dim myrcp As Recipient
    For Each myrcp In myappt.Recipients
                    If myrcp.Address = "rodk3150@gmail.com" Then
                        Set myappt = Nothing
                        Set myrcp = Nothing
                        Set thisappt = Nothing
                        Exit Sub
                    End If
    Next
    'Set myappt = mymeeting.GetAssociatedAppointment(False)
    Dim found As Boolean
    If myappt.MeetingStatus = olMeeting Then found = True
    If myappt.MeetingStatus = 5 Then found = True
    If myappt.Recipients.Count > 1 Then found = True
    If found = False Then Exit Sub
    Dim mysame As Outlook.AppointmentItem
    For Each mysame In Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
        If mysame.Mileage = myappt.EntryID Then
               mysame.MeetingStatus = olMeetingCanceled
                            mysame.Send
                            mysame.Delete
                Exit For
        End If
            
        If mysame.Start = myappt.Start Then
            If mysame.End = myappt.End Then
                If mysame.Duration = myappt.Duration Then
                    If mysame.Location = myappt.Location Then
                        If mysame.Subject <> myappt.Subject Then
                            mysame.MeetingStatus = olMeetingCanceled
                            mysame.Send
                            mysame.Delete
                        End If
                    End If
                End If
            End If
        End If
      Next
      'mysame = Nothing
      'myappt = Nothing
End If
errhandler:
If Err.Number <> 0 Then
    MsgBox Err.Description, vbOKOnly + vbInformation, "Meeting"
    Exit Sub
End If
End Sub







