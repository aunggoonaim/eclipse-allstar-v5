Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub CheckKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    ' Pick up Item
    If GetKeyState(vbKeySpace) < 0 And chaton = False Then
        CheckMapGetItem
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If

    'Move Up
    If GetKeyState(vbKeyUp) < 0 Then
        DirUp = True
        DirDown = False
        DirLeft = False
        DirRight = False
        Exit Sub
    Else
        DirUp = False
    End If
    
    'Move Up WASD
    If frmMain.txtMyChat.Visible = False Then
        If GetKeyState(vbKeyW) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirUp = False
        End If
    End If

    'Move Right
    If GetKeyState(vbKeyRight) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = True
        Exit Sub
    Else
        DirRight = False
    End If
    
    'Move Right WASD
    If frmMain.txtMyChat.Visible = False Then
        If GetKeyState(vbKeyD) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
            Exit Sub
        Else
            DirRight = False
        End If
    End If

    'Move down
    If GetKeyState(vbKeyDown) < 0 Then
        DirUp = False
        DirDown = True
        DirLeft = False
        DirRight = False
        Exit Sub
    Else
        DirDown = False
    End If
    
    'Move down WASD
    If frmMain.txtMyChat.Visible = False Then
        If GetKeyState(vbKeyS) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirDown = False
        End If
    End If

    'Move left
    If GetKeyState(vbKeyLeft) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = True
        DirRight = False
        Exit Sub
    Else
        DirLeft = False
    End If
    
    'Move left WASD
    If frmMain.txtMyChat.Visible = False Then
        If GetKeyState(vbKeyA) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
            Exit Sub
        Else
            DirLeft = False
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
Dim i As Long
Dim n As Long, x As Byte
Dim Command() As String
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ChatText = Trim$(MyText)

    If LenB(ChatText) = 0 Then
        If KeyAscii = vbKeyReturn Then
            chaton = Not chaton
            SetFocusOnGame
        End If
        Exit Sub
    End If
    
    If chaton = False Then Exit Sub
    
    MyText = LCase$(ChatText)

    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
    chaton = False
    SetFocusOnGame

        ' Emote message
     '   If Left$(ChatText, 1) = "-" Then
      '      MyText = Mid$(ChatText, 2, Len(ChatText) - 1)

       '     If Len(ChatText) > 0 Then
       '         Call EmoteMsg(ChatText)
       '     End If

       '     MyText = vbNullString
      '      frmMain.txtMyChat.text = vbNullString
     '       Exit Sub
     '   End If

        ' Player message
        If Left$(ChatText, 1) = "!" Then
            Exit Sub
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
            Name = vbNullString

            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)

                If Mid$(ChatText, i, 1) <> Space(1) Then
                    Name = Name & Mid$(ChatText, i, 1)
                Else
                    Exit For
                End If

            Next

            ChatText = Mid$(ChatText, i, Len(ChatText) - 1)

            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                MyText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                ' Send the message to the player
                Call PlayerMsg(ChatText, Name)
            Else
                Call AddText("�Ը��� : !���ͼ����� (��ͤ���)", AlertColor)
            End If

            MyText = vbNullString
            frmMain.txtMyChat.text = vbNullString
            Exit Sub
        End If

        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case Command(0)
                Case "/help"
                    Call AddText("����觵�ҧ � ���:", HelpColor)
                    Call AddText("����觷���Դ�����ҹ : /info ���ͼ����� [�ʴ�������], /who [��������͹�Ź�]", HelpColor)
                    Call AddText("/fps [�ʴ� fps], /fpslock [��Ѻ���������Ҿ] ,/guild [��Ŵ�]", HelpColor)
                Case "/guild"
                    If UBound(Command) < 1 Then
                        Call AddText("�������к���Ŵ� :", HelpColor)
                        Call AddText("���ҧ��Ŵ� : /guild make (���͡�Ŵ�)", HelpColor)
                        Call AddText("To transfer founder status use /guild founder (name)", HelpColor)
                        Call AddText("�ԭ��ҡ�Ŵ� : /guild invite (���ͼ�����)", HelpColor)
                        Call AddText("�͡�ҡ��Ŵ� : /guild leave", HelpColor)
                        Call AddText("�Դ��ǨѴ��á�Ŵ� (੾�����˹�ҡ�Ŵ�) : /guild admin", HelpColor)
                        Call AddText("��褹�͡�ҡ��Ŵ� : /guild kick (���ͼ�����)", HelpColor)
                        Call AddText("�غ��Ŵ�  : /guild disband yes", HelpColor)
                        Call AddText("�٢����ż�����㹡�Ŵ� : /guild view (online/all/offline)", HelpColor)
                        ' Call AddText("^Default is online, example: /guild view would display all online users.", HelpColor)
                        GoTo Continue
                    End If
                    
                    Select Case Command(1)
                        Case "make"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(1, Command(2))
                            Else
                                Call AddText("��سҵ�駪��͡�Ŵ�, �ٻẺ����� /guild make ���͡�Ŵ�", BrightRed)
                            End If
                            
                        Case "invite"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(2, Command(2))
                            Else
                                Call AddText("��س����͡������, �ٻẺ����� : /guild invite ���ͼ�����", BrightRed)
                            End If
                            
                        Case "leave"
                            Call GuildCommand(3, "")
                            
                        Case "admin"
                            Call GuildCommand(4, "")
                            
                        Case "view"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(5, Command(2))
                            Else
                                Call GuildCommand(5, "")
                            End If
                            
                        Case "accept"
                                Call GuildCommand(6, "")
                            
                        Case "decline"
                                Call GuildCommand(7, "")
                                
                        Case "founder"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(8, Command(2))
                            Else
                                Call AddText("��س����͡������, �ٻẺ����� : /guild founder ���ͼ�����", BrightRed)
                            End If
                        Case "kick"
                            If UBound(Command) = 2 Then
                                Call GuildCommand(9, Command(2))
                            Else
                                Call AddText("��س����͡������, �ٻẺ����� : /guild kick ���ͼ�����", BrightRed)
                            End If
                        Case "disband"
                            If UBound(Command) = 2 Then
                                If LCase(Command(2)) = LCase("yes") Then
                                    Call GuildCommand(10, "")
                                Else
                                    Call AddText("�������� : /guild disband yes", BrightRed)
                                End If
                            Else
                                Call AddText("�������� : /guild disband yes", BrightRed)
                            End If
                            
                    End Select
                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "�Ը������� : /info ���ͼ�����", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "�Ը������� : /info ���ͼ�����", AlertColor
                        GoTo Continue
                    End If

                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CPlayerInfoRequest
                    Buffer.WriteString Command(1)
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                    frmOnline.Show
                Case "/testlvskill"
                    For x = 1 To MAX_PLAYER_SPELLS
                        Call AddText("Skill " & x & " : " & skillLV(x), BrightGreen)
                    Next
                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    ' toggle fps lock
                Case "/dmap"
                    DMAP = Not DMAP
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' Request stats
                Case "/stats"
                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CGetStats
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing
                    ' // Monitor Admin Commands //
                    ' Admin Help
                Case "/Online"
                    frmOnline.Show
                Case "/admin"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue
                    frmMain.picAdmin.Visible = Not frmMain.picAdmin.Visible
                    If frmMain.picAdmin.Visible = True Then frmMain.Width = 14220
                    If frmMain.picAdmin.Visible = False Then frmMain.Width = 11340
                    ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "�Ը������� : /kick ���ͼ�����", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "�Ըժ����� : /kick ���ͼ�����", AlertColor
                        GoTo Continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    BLoc = Not BLoc
                    ' Map Editor
                Case "/editmap"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                    
                    SendRequestEditMap
                    ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If

                    ' Setting sprite
                Case "/setsprite"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    SendSetSprite CLng(Command(1))
                    ' Map report
                Case "/mapreport"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapReport
                    ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(ChatText, Len(ChatText) - 5)
                    ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendBanList
                    ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo Continue
                    End If

                    SendBan Command(1)
                    ' // Developer Admin Commands //
                    ' Editing item request
                Case "/edititem"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditItem
                ' Editing animation request
                Case "/editanimation"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditAnimation
                    ' Editing npc request
                Case "/editnpc"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditNpc
                Case "/editresource"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditResource
                    ' Editing shop request
                Case "/editshop"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditShop
                    ' Editing spell request
                Case "/editspell"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditSpell
                Case "/editquest"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                    SendRequestEditQuest
                    ' // Creator Admin Commands //
                    ' Giving another player access
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                    ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    SendBanDestroy
                    ' Packet debug mode
                Case "/debug"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    DEBUG_MODE = (Not DEBUG_MODE)
                Case Else
                    AddText "��辺�����觹����к���� !", BrightRed
            End Select

            'continue label where we go instead of exiting the sub
Continue:
            MyText = vbNullString
            frmMain.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Say message
        If Len(ChatText) > 0 Then
            Select Case frmMain.cbMAP.ListIndex
                Case 0 ' Ἱ���
                    Call SayMsg(ChatText)
                Case 1 ' �����
                    Call BroadcastMsg(ChatText)
                Case 2 ' ������
                    Call SendPartyChatMsg(ChatText)
                Case 3 ' ��Ҥ�
                    Call GuildMsg(ChatText)
                Case Else
                    Call PlayerMsg("�բ�ͼԴ��Ҵ㹡�����͡ ������ʹ��� ��س��ͧ���������ѧ���.", MyIndex)
            End Select
        End If

        frmMain.tmrChat.Enabled = False
        frmMain.tmrChat.Enabled = True
        ' Reset Chat Again
        frmMain.tmrChat.Interval = 5000

        MyText = vbNullString
        frmMain.txtMyChat.text = vbNullString
        Exit Sub
    End If

    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            MyText = MyText & ChrW$(KeyAscii)
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
