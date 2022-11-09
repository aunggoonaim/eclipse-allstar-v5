Attribute VB_Name = "modGameLogic"
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Sub GameLoop()
Dim FrameTime As Long
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim i As Long
Dim WalkTimer As Long
Dim tmr25 As Long
Dim tmr100 As Long
Dim tmr1000 As Long
Dim tmr3000 As Long
Dim tmr10000 As Long
' Dim LayerAnimTimer As Long
Dim SkillDelay As Double
Dim MyEvent As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' *** Start GameLoop ***
    Do While InGame
        Tick = GetTickCount                            ' Set the inital tick
        ' ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        ' FrameTime = Tick                               ' Set the time second loop time to the first.
        MyEvent = False

        ' * Check surface timers *
        ' Sprites
        ' If Fps_Max = False Then
        
        If tmr10000 < Tick Then
        
            If Fps_Max = True And Fps_Max = False Then
            ' remove this !! when not use
        
            End If
            
            tmr10000 = Tick + 15000
        End If
       
        If tmr1000 < Tick Then
            
            ' Change map animation every 1000 milliseconds
            If MapAnimTimer < Tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = Tick + 1000
            End If
            
            tmr1000 = Tick + 1000
        End If

        If tmr3000 < Tick Then
            ' check ping
            Call GetPing
            ' Call DrawPing
            
            tmr3000 = Tick + 3000
        End If

        If tmr25 < Tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hWnd Or GetForegroundWindow() = frmEditor_Events.hWnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if we need to end the CD icon
            If NumSpellIcons > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i) > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (Spell(PlayerSpells(i)).CDTime * 1000) < Tick Then
                                SpellCD(i) = 0
                                BltPlayerSpells
                                BltHotbar
                            End If
                        End If
                    End If
                Next
            End If
            
            SkillDelay = 1
            
            ' ��Ǩ�ͺ���͹�
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Weapon)).DelayDown > 0 And Item(GetPlayerEquipment(MyIndex, Weapon)).DelayDown < 1 Then
                    SkillDelay = Item(GetPlayerEquipment(MyIndex, Weapon)).DelayDown
                    If SkillDelay <= 0 Then SkillDelay = 1
                End If
            End If
            
            ' ��Ǩ�ͺ��Ҷ���������ʡ������ ���������Թ�黡�� V2
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (((Spell(PlayerSpells(SpellBuffer)).CastTime * 100) * SkillDelay)) / (1 + (GetPlayerStat(MyIndex, willpower) / 50)) < Tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If
            
            ' Update inv animation
            If NumItems > 0 Then
                If tmr100 < Tick Then
                    BltAnimatedInvItems
                    tmr100 = Tick + 100
                End If
            End If
            
            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next
            
            ' update time
            If StunDuration > 0 Then
                StunTime = StunTime + 25
            Else
                StunTime = 0
            End If
            
            tmr25 = Tick + 25
        End If
        
        If Tick > EventChatTimer Then
            If frmMain.lblEventChat.Visible = False Then
                If frmMain.picEventChat.Visible Then
                    frmMain.picEventChat.Visible = False
                End If
            End If
        End If

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < Tick Then
            
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                End If
            Next i

            ' Process npc movements (actually move them)
            For i = 1 To Npc_HighIndex
                If Map.NPC(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If
            Next i

            If Map.CurrentEvents > 0 Then
                For i = 1 To Map.CurrentEvents
                    Call ProcessEventMovement(i)
                Next i
            End If
            
            WalkTimer = Tick + 20 ' edit this value to change WalkTimer
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics
        DoEvents
        Sleep 5
        
    If Fps_Max = True Then
        ' Not Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < Tick + 15
                DoEvents
                Sleep 10
            Loop
        End If
    Else
        ' Not Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < Tick + 30
                DoEvents
                Sleep 10
            Loop
        End If
    End If
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
        
    ' loop mapmusic if needed and its a mp3 file
    LoopMp3

    Loop

    frmMain.Visible = False
    
    If isLogging Then
        isLogging = False
        frmMain.picScreen.Visible = False
        frmMenu.Visible = True
        GettingMap = True
        ' ����ŧ����
        StopMidi
        PlayMidi Options.MenuMusic
    Else
        ' Shutdown the game
        frmLoad.Visible = True
        Call SetStatus("���ѧ�Դ��...")
        Call DestroyGame
    End If
    
    MyEvent = True

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long

Dim SpeedMove, MaxSpeedMove As Integer

SpeedMove = Int(GetPlayerStat(Index, Stats.Agility) / 50)
MaxSpeedMove = 2

'If SpeedMove > MaxSpeedMove Then SpeedMove = MaxSpeedMove


    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' ��������㹡���Թ
    Select Case Player(Index).Moving
        Case MOVING_WALKING: MovementSpeed = ((0.015) * (WALK_SPEED * SIZE_X)) ' + SpeedMove)
        ' Call AddText(ElapsedTime / 1000, Yellow)
        Case MOVING_RUNNING: MovementSpeed = ((0.015) * (WALK_SPEED * SIZE_X))
        ' Call AddText(ElapsedTime / 1000, Yellow)
        'Case MOVING_RUNNING: MovementSpeed = ((ElapsedTime / 1000) * (RUN_SPEED * SIZE_X) + SpeedMove)
        Case Else: Exit Sub
    End Select
    
    If Player(Index).Step = 0 Then Player(Index).Step = 1
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).yOffset = Player(Index).yOffset - MovementSpeed
            If Player(Index).yOffset < 0 Then Player(Index).yOffset = 0
        Case DIR_DOWN
            Player(Index).yOffset = Player(Index).yOffset + MovementSpeed
            If Player(Index).yOffset > 0 Then Player(Index).yOffset = 0
        Case DIR_LEFT
            Player(Index).xOffset = Player(Index).xOffset - MovementSpeed
            If Player(Index).xOffset < 0 Then Player(Index).xOffset = 0
        Case DIR_RIGHT
            Player(Index).xOffset = Player(Index).xOffset + MovementSpeed
            If Player(Index).xOffset > 0 Then Player(Index).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        Else
            If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - ((0.015) * (RUN_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0
                
            Case DIR_DOWN
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + ((0.015) * (RUN_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0
                
            Case DIR_LEFT
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - ((0.015) * (RUN_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0
                
            Case DIR_RIGHT
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + ((0.015) * (RUN_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
                
                ' * 0.015 = ElapsedTime / 1000
                
        End Select
    
        ' Check if completed walking over to the next tile
        If MapNpc(MapNpcNum).Moving > 0 Then
            If MapNpc(MapNpcNum).Dir = DIR_RIGHT Or MapNpc(MapNpcNum).Dir = DIR_DOWN Then
                If (MapNpc(MapNpcNum).xOffset >= 0) And (MapNpc(MapNpcNum).yOffset >= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 1 Then
                        MapNpc(MapNpcNum).Step = 3
                    Else
                        MapNpc(MapNpcNum).Step = 1
                    End If
                End If
            Else
                If (MapNpc(MapNpcNum).xOffset <= 0) And (MapNpc(MapNpcNum).yOffset <= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 1 Then
                        MapNpc(MapNpcNum).Step = 3
                    Else
                        MapNpc(MapNpcNum).Step = 1
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckMapGetItem()
Dim Buffer As New clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer

    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            Player(MyIndex).MapGetTimer = GetTickCount
            Buffer.WriteLong CMapGetItem
            SendData Buffer.ToArray()
        End If
    End If

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAttack()
Dim Buffer As clsBuffer
Dim AttackSpeed As Long, X As Long, Y As Long, i As Long

    ' ��Ǩ�ͺ Debug mode ���ӧҹ�����͹䢶�Ҿ��Ѥ
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ControlDown Then
    If InEvent = True Then Exit Sub
    
        If SpellBuffer > 0 Then Exit Sub ' �������բ����ʡ��
        If StunDuration > 0 Then Exit Sub ' ��ҵԴʶҹ��ֹ �з�������������

        frmMain.txtMyChat.Visible = False

        ' ��������㹡������
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            AttackSpeed = ((2000 + Item(GetPlayerEquipment(MyIndex, Weapon)).Speed) - Item(GetPlayerEquipment(MyIndex, Weapon)).SpeedLow) - ((GetPlayerStat(MyIndex, Stats.Agility) * 5))
        Else
            AttackSpeed = 2000 - ((GetPlayerStat(MyIndex, Stats.Agility) * 5))
        End If

        ' ��˹� Attack speed �٧�ش ��� 5���� / �Թҷ�
        If AttackSpeed < 200 Then
            AttackSpeed = 200
        End If

        If Player(MyIndex).AttackTimer + AttackSpeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = GetTickCount
                End With
                
                If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic > 0 Then
                        ' projectile
                        Set Buffer = New clsBuffer
                            Buffer.WriteLong CProjecTileAttack
                            SendData Buffer.ToArray()
                            Set Buffer = Nothing
                            Exit Sub
                    End If
                End If

                ' non projectile
                Set Buffer = New clsBuffer
                Buffer.WriteLong CAttack
                SendData Buffer.ToArray()
                Set Buffer = Nothing
            End If
        End If
        
        Select Case Player(MyIndex).Dir
            Case DIR_UP
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) - 1
            Case DIR_DOWN
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) + 1
            Case DIR_LEFT
                X = GetPlayerX(MyIndex) - 1
                Y = GetPlayerY(MyIndex)
            Case DIR_RIGHT
                X = GetPlayerX(MyIndex) + 1
                Y = GetPlayerY(MyIndex)
        End Select
        
        If GetTickCount > Player(MyIndex).EventTimer Then
            For i = 1 To Map.CurrentEvents
                If Map.MapEvents(i).Visible = 1 Then
                    If Map.MapEvents(i).X = X And Map.MapEvents(i).Y = Y Then
                        Set Buffer = New clsBuffer
                        Buffer.WriteLong CEvent
                        Buffer.WriteLong i
                        SendData Buffer.ToArray()
                        Set Buffer = Nothing
                        Player(MyIndex).EventTimer = GetTickCount + 200
                    End If
                End If
            Next
        End If

    End If
    

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function IsTryingToMove() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CanMove() As Boolean
Dim d As Long
Dim spellslot As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    If InEvent Then
        CanMove = False
        Exit Function
    End If

    ' �ҧʡ������ö�Թ������
    If SpellBuffer > 0 Then
        'For spellslot = 1 To MAX_PLAYER_SPELLS
            'If Spell(PlayerSpells(SpellBuffer)).CanMove <= 0 Then
                CanMove = False
            'End If
        'Next
        Exit Function
    End If
    
    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' in bank
    If InBank Then
        'CanMove = False
        'Exit Function
        InBank = False
        frmMain.picCover.Visible = False
        frmMain.picBank.Visible = False
    End If

    If frmMain.picCurrency.Visible = True Then
        CanMove = False
        Exit Function
    End If

    d = GetPlayerDir(MyIndex)

    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        
        If Last_Dir <> GetPlayerDir(MyIndex) Then
                 Call SendPlayerDir
                 Last_Dir = GetPlayerDir(MyIndex)
        End If

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False
                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        
        If Last_Dir <> GetPlayerDir(MyIndex) Then
                 Call SendPlayerDir
                 Last_Dir = GetPlayerDir(MyIndex)
        End If
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.maxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False
                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        
        If Last_Dir <> GetPlayerDir(MyIndex) Then
                 Call SendPlayerDir
                 Last_Dir = GetPlayerDir(MyIndex)
        End If

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False
                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        If Last_Dir <> GetPlayerDir(MyIndex) Then
                 Call SendPlayerDir
                 Last_Dir = GetPlayerDir(MyIndex)
        End If

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.maxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False
                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean
Dim X As Long
Dim Y As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    CheckDirection = False
    
    ' check directional blocking
    If isDirBlocked(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, Direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case Direction
        Case DIR_UP
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
    End Select

    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the key door is open or not
    If Map.Tile(X, Y).Type = TILE_TYPE_KEY Then

        ' This actually checks if its open or not
        If TempTile(X, Y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If
       ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' ��Ǩ�ͺ����Թ���آͧ Event
    For i = 1 To Map.CurrentEvents
        If Map.MapEvents(i).Visible = 1 Then
            If Map.MapEvents(i).X = X Then
                If Map.MapEvents(i).Y = Y Then
                    If Map.MapEvents(i).WalkThrough = 0 Then ' ��� Event ��� ��꡶١��� Walk Through �зӡ���͡�ѧ��蹹��ѹ���
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    If Map.Moral = MAP_MORAL_SAFE Then Exit Function ' ��Ǩ�ͺ���Ἱ�������ࢵ��ʹ��� ���͡�ѧ��蹹��
    If Map.Moral = MAP_MORAL_PETARENA Then Exit Function ' ��Ǩ�ͺ���Ἱ�������Ἱ����ҹ����ͧ�ѵ������§ ���͡�ѧ��蹹��
    
    ' ��Ǩ�ͺ����Թ���ؼ�����
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If GetPlayerX(i) = X Then
                If GetPlayerY(i) = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next i

    ' Error handler
    Exit Function
errorhandler:
    HandleError "checkDirection", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub CheckMovement()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If IsTryingToMove Then
        If CanMove Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = MOVING_RUNNING
            Else
                Player(MyIndex).Moving = MOVING_WALKING
            End If

            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select

            If Player(MyIndex).xOffset = 0 Then
                If Player(MyIndex).yOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isInBounds()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If (CurX >= 0) Then
        If (CurX <= Map.maxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.maxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub UpdateDrawMapName()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DrawMapNameX = Camera.Left + ((MAX_MAPX + 1) * PIC_X / 2) - getWidth(TexthDC, Trim$(Map.Name))
    DrawMapNameY = Camera.Top + 1

    Select Case Map.Moral
        Case MAP_MORAL_NONE
            DrawMapNameColor = QBColor(BrightRed)
        Case MAP_MORAL_SAFE
            DrawMapNameColor = QBColor(White)
        Case MAP_MORAL_PETARENA
            DrawMapNameColor = QBColor(Blue)
        Case Else
            DrawMapNameColor = QBColor(White)
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateDrawMapName", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UseItem()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ForgetSpell(ByVal spellslot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellslot) > 0 Then
        AddText "�������öźʡ�Ţ�д������� !", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellslot Then
        AddText "�������öźʡ�Ţ������ !", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellslot) > 0 Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CForgetSpell
        Buffer.WriteLong spellslot
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    Else
        AddText "�����ʡ��㹪�ͧ���.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal spellslot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If SpellCD(spellslot) > 0 Then
        If FormatNumber(((SpellCD(spellslot) + (Spell(PlayerSpells(spellslot)).CDTime * 1000)) - GetTickCount) / 1000, 1) >= 1 Then
            'Call CreateActionMsg("ʡ�� " & Trim$(Spell(PlayerSpells(spellslot)).Name) & " ���ѧ������ [" & FormatNumber(((SpellCD(spellslot) + (Spell(PlayerSpells(spellslot)).CDTime * 1000)) - GetTickCount) / 1000, 1) & " s].", BrightRed, 0, Player(MyIndex).X + (Len("ʡ�� " & Trim$(Spell(PlayerSpells(spellslot)).Name) & " ���ѧ������ [" & FormatNumber(((SpellCD(spellslot) + (Spell(PlayerSpells(spellslot)).CDTime * 1000)) - GetTickCount) / 1000, 1) & " s].") * 3.5), frmMain.picScreen.Height - 6)
            AddText "ʡ�� " & Trim$(Spell(PlayerSpells(spellslot)).Name) & " ���ѧ������ [" & FormatNumber(((SpellCD(spellslot) + (Spell(PlayerSpells(spellslot)).CDTime * 1000)) - GetTickCount) / 1000, 1) & " s].", BrightRed
        Else
            'Call CreateActionMsg("ʡ�� " & Trim$(Spell(PlayerSpells(spellslot)).Name) & " ���ѧ������ [0" & FormatNumber(((SpellCD(spellslot) + (Spell(PlayerSpells(spellslot)).CDTime * 1000)) - GetTickCount) / 1000, 1) & " s].", BrightRed, 0, Player(MyIndex).X + (Len("ʡ�� " & Trim$(Spell(PlayerSpells(spellslot)).Name) & " ���ѧ������ [0" & FormatNumber(((SpellCD(spellslot) + (Spell(PlayerSpells(spellslot)).CDTime * 1000)) - GetTickCount) / 1000, 1) & " s].") * 3.5), frmMain.picScreen.Height - 6)
            AddText "ʡ�� " & Trim$(Spell(PlayerSpells(spellslot)).Name) & " ���ѧ������ [0" & FormatNumber(((SpellCD(spellslot) + (Spell(PlayerSpells(spellslot)).CDTime * 1000)) - GetTickCount) / 1000, 1) & " s].", BrightRed
        End If
        Exit Sub
    End If

    
    If PlayerSpells(spellslot) = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellslot)).MPCost Then
        Call AddText("��ͧ��� Mp 㹡������ " & Trim$(Spell(PlayerSpells(spellslot)).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellslot) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong CCast
                Buffer.WriteLong spellslot
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                SpellBuffer = spellslot
                SpellBufferTimer = GetTickCount
            Else
                If Spell(PlayerSpells(spellslot)).CanMove > 0 Then
                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CCast
                    Buffer.WriteLong spellslot
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing
                    SpellBuffer = spellslot
                    SpellBufferTimer = GetTickCount
                Else
                    Call AddText("�������ö����ʡ�Ţ���Թ�� !", BrightRed)
                    ' Make sure they aren't walking
                    'Player(MyIndex).Moving = 0
                    'Player(MyIndex).xOffset = 0
                    'Player(MyIndex).yOffset = 0
                End If
            End If
        End If
    Else
        Call AddText("�����ʡ��㹪�ͧ���.", BrightRed)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CastSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearTempTile()
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ReDim TempTile(0 To Map.maxX, 0 To Map.maxY)

    For X = 0 To Map.maxX
        For Y = 0 To Map.maxY
            TempTile(X, Y).DoorOpen = NO
        Next
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearTempTile", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DevMsg(ByVal text As String, ByVal Color As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(text, Color)
        End If
    End If

    Debug.Print text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DevMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "TwipsToPixels", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "PixelsToTwips", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub DrawPing()
Dim PingToDraw As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PingToDraw = Ping

    Select Case Ping
        Case -1
            PingToDraw = "Network : Internet  � Ping : " & Ping
        Case 0 To 5
            PingToDraw = "Network : Lan � Ping : " & Ping
    End Select

    frmMain.lblPing.Caption = PingToDraw
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPing", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateSpellWindow(ByVal spellnum As Long, ByVal X As Long, ByVal Y As Long, ByVal spellslot As Long)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check for off-screen
    If Y + frmMain.picSpellDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picSpellDesc.Height
    End If
    
    With frmMain
        .picSpellDesc.Top = Y
        .picSpellDesc.Left = X
        .picSpellDesc.Visible = True
        .imgEXPSKILL.Visible = True
        .lblEXPSKILL.Visible = True
        
        'If GetPlayerNextLevelSkill(MyIndex, spellslot) > 1 Then
        '    .lblEXPSKILL.Caption = Player(MyIndex).skillEXP(spellslot) & " / " & GetPlayerNextLevelSkill(MyIndex, spellslot)
       ' Else
       '     .lblEXPSKILL.Caption = "--- MAX ---"
        'End If
            
        .imgEXPSKILL.Width = ((Player(MyIndex).skillEXP(spellslot) / SKILLBar_Width) / (GetPlayerNextLevelSkill(MyIndex, spellslot) / SKILLBar_Width)) * SKILLBar_Width
        
        ' spell level
        If skillLV(spellslot) < MAX_SKILL_LEVEL Then
            .lblSpellName.Caption = Trim$(Spell(spellnum).Name) & "[Lv." & skillLV(spellslot) & "]"
            .lblEXPSKILL.Caption = Player(MyIndex).skillEXP(spellslot) & " / " & GetPlayerNextLevelSkill(MyIndex, spellslot)
        Else
            .lblSpellName.Caption = Trim$(Spell(spellnum).Name)
            .lblEXPSKILL.Caption = "--- MAX ---"
        End If
        
        If skillLV(spellslot) > MAX_SKILL_LEVEL Then
            .lblEXPSKILL.Caption = "--- NO DATA ---"
        End If
        .lblSpellDesc.Caption = Trim$(Spell(spellnum).Desc)
        
        If LastSpellDesc = spellnum Then Exit Sub
        BltSpellDesc spellnum
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdteSpellWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateDescWindow(ByVal itemnum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long
Dim FirstLetter As String * 1
Dim Name As String
Dim Item1 As Long
Dim Item2 As Long
Dim Tool As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    FirstLetter = LCase$(Left$(Trim$(Item(itemnum).Name), 1))
   
    If FirstLetter = "$" Then
        Name = (Mid$(Trim$(Item(itemnum).Name), 2, Len(Trim$(Item(itemnum).Name)) - 1))
    Else
        Name = Trim$(Item(itemnum).Name)
    End If
    
    ' check for off-screen
    If Y + frmMain.picItemDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picItemDesc.Height
    End If
    
    ' set z-order
    frmMain.picItemDesc.ZOrder (0)
    
    Item1 = Item(itemnum).Data1
    Item2 = Item(itemnum).Data2
    
    Select Case Item(itemnum).ToolReq
        Case 0 'None
            Tool = "None"
        Case 1 'Hammer
            Tool = "Hammer"
        Case 2 'Mortar & Pestle
            Tool = "Mortar and Pestle"
    End Select

    With frmMain
        .picItemDesc.Top = Y
        .picItemDesc.Left = X
        .picItemDesc.Visible = True

        If LastItemDesc = itemnum Then Exit Sub ' exit out after setting x + y so we don't reset values

        ' set the name
        Select Case Item(itemnum).Rarity
        ' �س���ѵ� �ժ���/�� ��������
            Case 0 ' white. Item normal.
                .lblItemName.ForeColor = RGB(255, 255, 255)
                .lblItemName.Caption = Name & " [������] "
            Case 1 ' Item Green
                .lblItemName.ForeColor = RGB(0, 255, 0)
                .lblItemName.Caption = Name & " [�Դ�] All +5"
            Case 2 ' Item Yellow
                .lblItemName.ForeColor = RGB(255, 255, 0)
                .lblItemName.Caption = Name & " [�ӹҹ] All +10"
            Case 3 ' Item Red
                .lblItemName.ForeColor = RGB(255, 0, 0)
                .lblItemName.Caption = Name & " [෾] All +15"
        End Select
        
        ' set captions
        '.lblItemName.Caption = Name
        .lblItemDesc.Caption = Trim$(Item(itemnum).Desc)
        
        If Item(itemnum).Type = ITEM_TYPE_RECIPE Then
            If Trim$(Tool) <> "None" Then
                .lblItemDesc.Caption = "㺻�Сͺ��������ͧ������� " & Trim$(Item(Item1).Name) & " ��� " & Trim$(Item(Item2).Name) & ". [�µ�ͧ������ͧ���] : " & Trim$(Tool)
            Else
                .lblItemDesc.Caption = "㺻�Сͺ��������ͧ������� " & Trim$(Item(Item1).Name) & " ��� " & Trim$(Item(Item2).Name) & ". [�µ�ͧ������ͧ���] : ���á���."
            End If
        End If
        
        ' render the item
        BltItemDesc itemnum
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateDescWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CacheResources()
Dim X As Long, Y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource_Count = 0

    For X = 0 To Map.maxX
        For Y = 0 To Map.maxY
            If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If
        Next
    Next

    Resource_Index = Resource_Count
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateActionMsg(ByVal Message As String, ByVal Color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .Message = Message
        .Color = Color
        .Type = MsgType
        .Created = GetTickCount
        .Scroll = 1
        .X = X
        .Y = Y
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    End If
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsg(Index).Message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).Color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).X = 0
    ActionMsg(Index).Y = 0
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
Dim looptime As Long
Dim Layer As Long
Dim FrameCount As Long
Dim lockindex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub
    
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).FrameIndex(Layer) = 0 Then AnimInstance(Index).FrameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).Timer(Layer) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimInstance(Index).FrameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).FrameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).FrameIndex(Layer) = AnimInstance(Index).FrameIndex(Layer) + 1
                End If
                AnimInstance(Index).Timer(Layer) = GetTickCount
            End If
        End If
    Next
    
    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "checkAnimInstance", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenShop(ByVal shopnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InShop = shopnum
    ShopAction = 0
    frmMain.picCover.Visible = True
    frmMain.picShop.Visible = True
    BltShop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    GetBankItemNum = Bank.Item(bankslot).Num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal itemnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).Num = itemnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemValue(ByVal Index As Long, ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    
    GetBankItemValue = Bank.Item(bankslot).Value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).Value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef Dir As Byte, ByVal Block As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Block Then
        blockvar = blockvar Or (2 ^ Dir)
    Else
        blockvar = blockvar And Not (2 ^ Dir)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsHotbarSlot(ByVal X As Single, ByVal Y As Single) As Long
Dim Top As Long, Left As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    IsHotbarSlot = 0

    For i = 1 To MAX_HOTBAR
        Top = HotbarTop
        Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
        If X >= Left And X <= Left + PIC_X Then
            If Y >= Top And Y <= Top + PIC_Y Then
                IsHotbarSlot = i
                Exit Function
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim soundName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If entityNum <= 0 Then Exit Sub
    
    ' find the sound
    Select Case entityType
        ' animations
        Case SoundEntity.seAnimation
            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).Sound)
            
        ' items
        Case SoundEntity.seItem
            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).Sound)
        ' npcs
        Case SoundEntity.seNpc
            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(NPC(entityNum).Sound)
        ' resources
        Case SoundEntity.seResource
            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).Sound)
        ' spells
        Case SoundEntity.seSpell
            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(Spell(entityNum).Sound)
        ' LevelUp
        Case SoundEntity.seLevelUp
            soundName = Trim$(LEVEL_SOUND)
        Case SoundEntity.sePunch
            soundName = Trim$(PUNCH_SOUND)
        Case SoundEntity.seDie
            soundName = Trim$(DIE_SOUND)
        Case Else
        ' other
            Exit Sub
    End Select
    
    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    PlaySound soundName
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal Data1 As Long = 0)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' exit out if we've already got a dialogue open
    If dialogueIndex > 0 Then Exit Sub
    
    ' set global dialogue index
    dialogueIndex = diIndex
    
    ' set the global dialogue data
    dialogueData1 = Data1

    ' set the captions
    frmMain.lblDialogue_Title.Caption = diTitle
    frmMain.lblDialogue_Text.Caption = diText
    
    ' show/hide buttons
    If Not isYesNo Then
        frmMain.lblDialogue_Button(1).Visible = True ' Okay button
        frmMain.lblDialogue_Button(2).Visible = False ' Yes button
        frmMain.lblDialogue_Button(3).Visible = False ' No button
    Else
        frmMain.lblDialogue_Button(1).Visible = False ' Okay button
        frmMain.lblDialogue_Button(2).Visible = True ' Yes button
        frmMain.lblDialogue_Button(3).Visible = True ' No button
    End If
    
    ' show the dialogue box
    frmMain.picDialogue.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Dialogue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub dialogueHandler(ByVal Index As Long)
    ' find out which button
    If Index = 1 Then ' okay button
        ' dialogue index
        Select Case dialogueIndex
        
        End Select
    ElseIf Index = 2 Then ' yes button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendAcceptTradeRequest
            Case DIALOGUE_TYPE_FORGET
                ForgetSpell dialogueData1
            Case DIALOGUE_TYPE_PARTY
                SendAcceptParty
        End Select
    ElseIf Index = 3 Then ' no button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendDeclineTradeRequest
            Case DIALOGUE_TYPE_PARTY
                SendDeclineParty
        End Select
    End If
End Sub


Sub SpawnPet(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CSpawnPet
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing

End Sub

Sub PetFollow(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CPetFollowOwner
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub PetAttack(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CPetAttackTarget
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub PetWander(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CPetWander
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub PetDisband(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CPetDisband
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub ProcessEventMovement(ByVal id As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if NPC is walking, and if so process moving them over
    If Map.MapEvents(id).Moving = 1 Then
        
        Select Case Map.MapEvents(id).Dir
            Case DIR_UP
                Map.MapEvents(id).yOffset = Map.MapEvents(id).yOffset - ((0.015) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).yOffset < 0 Then Map.MapEvents(id).yOffset = 0
                
            Case DIR_DOWN
                Map.MapEvents(id).yOffset = Map.MapEvents(id).yOffset + ((0.015) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).yOffset > 0 Then Map.MapEvents(id).yOffset = 0
                
            Case DIR_LEFT
                Map.MapEvents(id).xOffset = Map.MapEvents(id).xOffset - ((0.015) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).xOffset < 0 Then Map.MapEvents(id).xOffset = 0
                
            Case DIR_RIGHT
                Map.MapEvents(id).xOffset = Map.MapEvents(id).xOffset + ((0.015) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).xOffset > 0 Then Map.MapEvents(id).xOffset = 0
                
                ' * 0.015 = ElapsedTime / 1000
                
        End Select
    
        ' Check if completed walking over to the next tile
        If Map.MapEvents(id).Moving > 0 Then
            If Map.MapEvents(id).Dir = DIR_RIGHT Or Map.MapEvents(id).Dir = DIR_DOWN Then
                If (Map.MapEvents(id).xOffset >= 0) And (Map.MapEvents(id).yOffset >= 0) Then
                    Map.MapEvents(id).Moving = 0
                    If Map.MapEvents(id).Step = 1 Then
                        Map.MapEvents(id).Step = 3
                    Else
                        Map.MapEvents(id).Step = 1
                    End If
                End If
            Else
                If (Map.MapEvents(id).xOffset <= 0) And (Map.MapEvents(id).yOffset <= 0) Then
                    Map.MapEvents(id).Moving = 0
                    If Map.MapEvents(id).Step = 1 Then
                        Map.MapEvents(id).Step = 3
                    Else
                        Map.MapEvents(id).Step = 1
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessEventMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetColorString(Color As Long)
    Select Case Color
        Case 0
            GetColorString = "Black"
        Case 1
            GetColorString = "Blue"
        Case 2
            GetColorString = "Green"
        Case 3
            GetColorString = "Cyan"
        Case 4
            GetColorString = "Red"
        Case 5
            GetColorString = "Magenta"
        Case 6
            GetColorString = "Brown"
        Case 7
            GetColorString = "Grey"
        Case 8
            GetColorString = "Dark Grey"
        Case 9
            GetColorString = "Bright Blue"
        Case 10
            GetColorString = "Bright Green"
        Case 11
            GetColorString = "Bright Cyan"
        Case 12
            GetColorString = "Bright Red"
        Case 13
            GetColorString = "Pink"
        Case 14
            GetColorString = "Yellow"
        Case 15
            GetColorString = "White"

    End Select
End Function

Sub ClearEventChat()
    Dim i As Long
    If AnotherChat = 1 Then
        For i = 1 To 4
            frmMain.lblChoices(i).Visible = False
        Next
        
        frmMain.lblEventChat.Caption = ""
        frmMain.lblEventChatContinue.Visible = False
    ElseIf AnotherChat = 2 Then
        For i = 1 To 4
            frmMain.lblChoices(i).Visible = False
        Next
        
        frmMain.lblEventChat.Visible = False
        frmMain.lblEventChatContinue.Visible = False
        EventChatTimer = GetTickCount + 100
    Else
        frmMain.picEventChat.Visible = False
    End If

End Sub

Sub MultiClient()
    Dim GameClient As Long
    Dim OldAppName As String
    OldAppName = App.Title
    App.Title = ""
    GameClient = FindWindow(vbNullString, OldAppName)
    App.Title = OldAppName
    If App.PrevInstance = True Or GameClient < 0 Then
        Call MsgBox("MWO ���ѧ�ӧҹ���� ! �������ö�Դ�����.", vbExclamation, "Error")
        End
    End If
End Sub

Function GetPlayerNextLevelSkill(ByVal Index As Long, ByVal spellslot As Long) As Long

If Player(Index).skillLV(spellslot) > 0 And Player(Index).skillLV(spellslot) <= MAX_SKILL_LEVEL Then
    Select Case Player(Index).skillLV(spellslot)
    Case 1: GetPlayerNextLevelSkill = 50
    Case 2: GetPlayerNextLevelSkill = 300
    Case 3: GetPlayerNextLevelSkill = 2000
    Case 4: GetPlayerNextLevelSkill = 5000
    Case 5: GetPlayerNextLevelSkill = 15000
    Case 6: GetPlayerNextLevelSkill = 50000
    Case 7: GetPlayerNextLevelSkill = 100000
    Case 8: GetPlayerNextLevelSkill = 500000
    Case 9: GetPlayerNextLevelSkill = 1000000
    Case MAX_LEVELS: GetPlayerNextLevelSkill = 1
    End Select
Else
    GetPlayerNextLevelSkill = 15
End If

End Function
