Attribute VB_Name = "Notes"
'TODOS BEFORE GENERATING EXE
'SET KERNEL/DEVELOPMENT TO A VALUE THAT NOT EQUAL TO 1
'
'
'
'
'
'COLOR ARRANGEMENT
'TextBox &HB4BFCC
'ForeColors &H656D76
'BackGround &HA0ACBA
'
'NOTE0 SAVE A IMAGE
'Dim usage As Long
'usage = GetSetting("FreeExam", "Create", "TrackNumUsage", 1000)
'If Dir(App.Path & "\Cache", vbDirectory) = "" Then MkDir App.Path & "\Cache"
'SavePicture Preview.Exports.Picture, App.Path & "\Cache\" & usage + 1 & ".jpg"
'SaveSetting "FreeExam", "Create", "TrackNumUsage", usage + 1
'List1.AddItem "P" & Left(Combo1.Text, 1) & usage + 1

''===============================NOTE1 COMMON MESSAGE SENDER===============================
'Dim showcnt As long, current As long
'Sub NewMessage(Content As String, Color As Long, Optional ClearList As Boolean = False, Optional ClearOnly = False)
'    current = -1
'    If (ClearOnly And Not ClearList) Then
'        RaiseSysErr "Clear message list only and do not clear message list were both turned on.", "Create/PageSettings/NewEvent"
'        Exit Sub
'    End If
'    If ClearList Then
'        MsgContentList.Clear
'        MsgColorList.Clear
'        MsgTypeList.Clear
'        If Message.Caption <> "" Then Message.Caption = Message.Caption & "(Expired)"
'        If ClearOnly Then Exit Sub
'    End If
'    MsgContentList.AddItem Content
'    MsgColorList.AddItem Color
'    Select Case Color
'        Case vbBlack: MsgTypeList.AddItem "[Info]"
'        Case vbBlue: MsgTypeList.AddItem "[Warning]"
'        Case vbRed: MsgTypeList.AddItem "[Error]"
'    End Select
'    showcnt = 49
'    Timer1_Timer
'End Sub
'Private Sub Form_Load()
'    current = -1
'End Sub
'Private Sub Message_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Timer1.Interval = 1000
'End Sub
'Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Timer1.Interval = 1000
'End Sub
'Private Sub Timer1_Timer()
'    Dim first As Long
'    If Timer1.Interval > 100 Then Timer1.Interval = Timer1.Interval - 100
'    showcnt = showcnt + 1
''    If MsgContentList.ListCount <= 1 Then
''        first = showcnt
''        showcnt = ShowCntPerMsg
''        Message.Caption = ""
''        If MsgContentList.ListCount = 1 Then
''            current = 0
''            MsgContentList.ListIndex = current
''            MsgColorList.ListIndex = current
''            MsgTypeList.ListIndex = current
''            Message.Caption = MsgTypeList.Text & MsgContentList.Text
''            Message.ForeColor = ReverseColor(MsgColorList.Text)
''        End If
''        If showcnt <> first Then ProgressBar.Width = showcnt / ShowCntPerMsg * Picture1.Width
''        Exit Sub
''    End If
'    If MsgContentList.ListCount = 0 Then
'        Message.Caption = translate("No new messages.")
'        Message.ForeColor = vbWhite
'        showcnt = ShowCntPerMsg - 1
'        GoTo rrr
'    End If
'    If current >= MsgContentList.ListCount Then
'        Message.Caption = translate("No new messages.")
'        Message.ForeColor = vbWhite
'        showcnt = ShowCntPerMsg - 1
'        GoTo rrr
'    End If
'    If showcnt = ShowCntPerMsg Then
'        If MsgContentList.ListCount = 0 Then
'            ProgressBar.Width = 15
'            Message.Caption = ""
'            Exit Sub
'        End If
'        If current + 1 >= MsgContentList.ListCount Then
'            Message.Caption = translate("No new messages.")
'            Message.ForeColor = vbWhite
'            showcnt = ShowCntPerMsg - 1
'            GoTo rrr
'        End If
'        showcnt = 0
'        current = current + 1
'        MsgContentList.ListIndex = current
'        MsgColorList.ListIndex = current
'        MsgTypeList.ListIndex = current
'        Message.Caption = MsgTypeList.Text & MsgContentList.Text
'        Message.ForeColor = ReverseColor(MsgColorList.Text)
'    End If
'rrr:
'    ProgressBar.Width = showcnt / ShowCntPerMsg * Picture1.Width
''    Message.Caption = Message.Caption & "(" & current + 1 & "/" & MsgTypeList.ListCount & ")"
'End Sub




'=====================NOTE2 UI/BUTTON=====================
'Private Sub ButtonName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    ButtonName.BackStyle = 1
'    ButtonName.BackColor = vbBlack
'    ButtonName.ForeColor = vbWhite
'End Sub
'
'Private Sub ButtonName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    ButtonName.BackStyle = 0
'    ButtonName.ForeColor = vbBlack
'End Sub
