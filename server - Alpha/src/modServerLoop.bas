Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, x As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long
    Dim mapnum, x1, y1, Anim, playerNumber As Long
    Dim BuffTimer As Long
    Dim LastUpdateMapLogic, LastUpdateLevelMaxAnim As Long
    Dim SkillDelay As Double
    Dim wPower As Double, PT100 As Long

    SkillDelay = 1
    ServerOnline = True

    Do While ServerOnline
        Tick = GetTickCount
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        If Tick > tmr25 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                    
                    ' ตรวจสอบเงื่อนไข
                        If GetPlayerEquipment(i, Weapon) > 0 Then
                            If Item(GetPlayerEquipment(i, Weapon)).DelayDown > 0 And Item(GetPlayerEquipment(i, Weapon)).DelayDown < 1 Then
                                SkillDelay = Item(GetPlayerEquipment(i, Weapon)).DelayDown
                                If SkillDelay <= 0 Then SkillDelay = 1
                            End If
                        End If
                    
                    wPower = (1 + (GetPlayerStat(i, willpower) / 50))
                    
                    ' สูตรคำนวนเวลาร่ายสกิล V2
                        If GetTickCount > TempPlayer(i).spellBuffer.Timer + ((((Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).CastTime * 100) * SkillDelay) / wPower)) Then ' / (1 + (GetPlayerStat(i, willpower) / 50))) Then
                            If Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).Passive <= 0 Then
                                CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.Target, TempPlayer(i).spellBuffer.tType
                                ' Call PlayerMsg(i, Trim$(Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).Name) & " ใช้งาน..", BrightGreen)
                                TempPlayer(i).spellBuffer.Spell = 0
                                TempPlayer(i).spellBuffer.Timer = 0
                                TempPlayer(i).spellBuffer.Target = 0
                                TempPlayer(i).spellBuffer.tType = 0
                                ' Update
                                ' SendPlayerData i
                            Else
                                ' ถ้าเป็นสกิลติดตัว จะสั่ง Passive ให้ทำงาน
                                CastSpellPassive i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.Target, TempPlayer(i).spellBuffer.tType
                                'Call PlayerMsg(i, "[สกิลติดตัว] " & Trim$(Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).Name) & ".", BrightGreen)
                                TempPlayer(i).spellBuffer.Spell = 0
                                TempPlayer(i).spellBuffer.Timer = 0
                                TempPlayer(i).spellBuffer.Target = 0
                                TempPlayer(i).spellBuffer.tType = 0
                                ' Update
                                ' SendPlayerData i
                            End If
                        End If
                    End If
                    
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    
                    ' ตรวจสอบเวลาในการฟื้นฟู
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer < GetTickCount Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For x = 1 To MAX_DOTS
                        HandleDoT_Player i, x
                        HandleHoT_Player i, x
                    Next
                                                                    
                    ' Guild invite
                    If TempPlayer(i).tmpGuildInviteSlot > 0 Then
                        If Tick > TempPlayer(i).tmpGuildInviteTimer Then
                            If GuildData(TempPlayer(i).tmpGuildInviteSlot).In_Use = True Then
                                PlayerMsg i, "หมดเวลาคำขอร้องเข้าร่วมกิล : " & GuildData(TempPlayer(i).tmpGuildInviteSlot).Guild_Name & " แล้ว.", BrightRed
                                TempPlayer(i).tmpGuildInviteSlot = 0
                                TempPlayer(i).tmpGuildInviteTimer = 0
                            Else
                                'Just remove this guild has been unloaded
                                TempPlayer(i).tmpGuildInviteSlot = 0
                                TempPlayer(i).tmpGuildInviteTimer = 0
                            End If
                        End If
                    End If
        
                ' fixed bug for buff
                SendPlayerBuff i
        
                End If
            Next
            
            UpdateEventLogic
            tmr25 = GetTickCount + 25
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).state > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            UpdateMapLogic
            
            ' Checks to update player vitals every 5 seconds - Can be tweaked
            If Tick > LastUpdatePlayerVitals Then
                UpdatePlayerVitals
                LastUpdatePlayerVitals = GetTickCount + 5000
            End If
                          
            ' Checks to spawn map items every 5 minutes - Can be tweaked
            If Tick > LastUpdateMapSpawnItems Then
                UpdateMapSpawnItems
                LastUpdateMapSpawnItems = GetTickCount + 300000
            End If

            ' Checks to save players every 5 minutes - Can be tweaked
            If Tick > LastUpdateSavePlayers Then
                UpdateSavePlayers
                LastUpdateSavePlayers = GetTickCount + 300000
            End If
            
            ' Update Anim Level max every 3 seconds เวลตันอนิเมชั่น
            If Tick > LastUpdateLevelMaxAnim Then
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If Player(i).Level = MAX_LEVELS Then
                            Call SendAnimation(GetPlayerMap(i), LEVELMAX_ANIM, 0, 0, TARGET_TYPE_PLAYER, i)
                            LastUpdateLevelMaxAnim = GetTickCount + 3000
                        End If
                    End If
                Next
            End If
                    
            tmr500 = GetTickCount + 500
        End If

        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            
            ' ckh buff every 1 sec
            For i = 1 To MAX_BUFF
                For x = 1 To Player_HighIndex
                    If Player(x).BuffTime(i) > 0 And Player(x).BuffStatus(i) > 0 Then
                        Player(x).BuffTime(i) = Player(x).BuffTime(i) - 1
                    Else
                        Player(x).BuffTime(i) = 0
                        Player(x).BuffStatus(i) = 0
                    End If
                Next
            Next
            
            tmr1000 = GetTickCount + 1000
        End If
                
        ' projectiles
        If Tick > PT100 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    For x = 1 To MAX_PLAYER_PROJECTILES
                        If TempPlayer(i).Projectile(x).Pic > 0 Then
                            ' handle the projec tile
                            HandleProjecTile i, x
                        End If
                    Next
                End If
            PT100 = GetTickCount + 12
            Next
        End If
      
        ' Anim every 1 sec
           If Tick > Anim Then
            Anim = GetTickCount + 1000
            For mapnum = 1 To MAX_MAPS
                For x1 = 0 To Map(mapnum).MaxX
                    For y1 = 0 To Map(mapnum).MaxY
                        If Map(mapnum).Tile(x1, y1).Type = TILE_TYPE_ANIMATION Then
                            For playerNumber = 1 To Player_HighIndex
                                If IsPlaying(playerNumber) Then
                                    If GetPlayerMap(playerNumber) = mapnum Then
                                        SendAnimation mapnum, Map(mapnum).Tile(x1, y1).Data1, x1, y1
                                    End If
                                End If
                            Next
                        End If
                    Next
                Next
            Next
        End If
        
        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
        
        ' Set server CPS on label
        frmServer.lblCPS.Caption = "CPS : " & Format$(GameCPS, "#,###,###,###")
        
    Loop
    
End Sub

Private Sub UpdateMapSpawnItems()
    Dim x As Long
    Dim y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For y = 1 To MAX_MAPS

        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(y) Then

            ' Clear out unnecessary junk
            For x = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(x, y)
            Next

            ' Spawn the items
            Call SpawnMapItems(y)
            Call SendMapItemsToAll(y)
        End If

        DoEvents
    Next

End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, x As Long, mapnum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, npcNum As Long
    Dim Target As Long, targetType As Byte, didwalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim TargetX As Long, TargetY As Long, target_verify As Boolean, sp As Long

    For mapnum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(mapnum, i).num > 0 Then
                If MapItem(mapnum, i).playerName <> vbNullString Then
                    ' make item public?
                    If MapItem(mapnum, i).playerTimer < GetTickCount Then
                        ' make it public
                        MapItem(mapnum, i).playerName = vbNullString
                        MapItem(mapnum, i).playerTimer = 0
                        ' send updates to everyone
                        SendMapItemsToAll mapnum
                    End If
                    ' despawn item?
                    If MapItem(mapnum, i).canDespawn Then
                        If MapItem(mapnum, i).despawnTimer < GetTickCount Then
                            ' despawn it
                            ClearMapItem i, mapnum
                            ' send updates to everyone
                            SendMapItemsToAll mapnum
                        End If
                    End If
                End If
            End If
        Next
        
        '  Close the doors
        If TickCount > TempTile(mapnum).DoorTimer + 5000 Then
            For x1 = 0 To Map(mapnum).MaxX
                For y1 = 0 To Map(mapnum).MaxY
                    If Map(mapnum).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(mapnum).DoorOpen(x1, y1) = YES Then
                        TempTile(mapnum).DoorOpen(x1, y1) = NO
                        SendMapKeyToMap mapnum, x1, y1, 0
                    End If
                Next
            Next
        End If
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(mapnum).NPC(i).num > 0 Then
                For x = 1 To MAX_DOTS
                    HandleDoT_Npc mapnum, i, x
                    HandleHoT_Npc mapnum, i, x
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(mapnum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(mapnum).Resource_Count
                Resource_index = Map(mapnum).Tile(ResourceCache(mapnum).ResourceData(i).x, ResourceCache(mapnum).ResourceData(i).y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(mapnum).ResourceData(i).ResourceState = 1 Or ResourceCache(mapnum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(mapnum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(mapnum).ResourceData(i).ResourceTimer = GetTickCount
                            ResourceCache(mapnum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(mapnum).ResourceData(i).cur_health = Resource(Resource_index).health
                            SendResourceCacheToMap mapnum, i
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(mapnum) = YES Then
            TickCount = GetTickCount
            
            For x = 1 To MAX_MAP_NPCS
                npcNum = MapNpc(mapnum).NPC(x).num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).NPC(x) > 0 And MapNpc(mapnum).NPC(x).num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(mapnum).NPC(x).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = mapnum And MapNpc(mapnum).NPC(x).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        n = NPC(npcNum).Range
                                        DistanceX = MapNpc(mapnum).NPC(x).x - GetPlayerX(i)
                                        DistanceY = MapNpc(mapnum).NPC(x).y - GetPlayerY(i)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                    
                                                ' fixed
                                                If Player(i).BuffStatus(BUFF_INVISIBLE) <> BUFF_INVISIBLE Then
                                                
                                                If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                                    If NPC(npcNum).AttackSay <> vbNullString Then
                                                        Call PlayerMsg(i, Trim$(NPC(npcNum).Name) & " พูด : " & Trim$(NPC(npcNum).AttackSay), SayColor)
                                                    End If
                                                End If
                                                
                                                MapNpc(mapnum).NPC(x).targetType = 1 ' player
                                                MapNpc(mapnum).NPC(x).Target = i
                                            
                                                End If
                                                
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).NPC(x) > 0 And MapNpc(mapnum).NPC(x).num > 0 Then
                    If MapNpc(mapnum).NPC(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(mapnum).NPC(x).StunTimer + (MapNpc(mapnum).NPC(x).StunDuration * 1000) Then
                            MapNpc(mapnum).NPC(x).StunDuration = 0
                            MapNpc(mapnum).NPC(x).StunTimer = 0
                        End If
                    Else
                            
                        Target = MapNpc(mapnum).NPC(x).Target
                        targetType = MapNpc(mapnum).NPC(x).targetType
    
                        ' Check to see if its time for the npc to walk
                        'If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        
                            If targetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If Target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(Target) And GetPlayerMap(Target) = mapnum Then
                                        didwalk = False
                                        target_verify = True
                                        TargetY = GetPlayerY(Target)
                                        TargetX = GetPlayerX(Target)
                                    Else
                                        MapNpc(mapnum).NPC(x).targetType = 0 ' clear
                                        MapNpc(mapnum).NPC(x).Target = 0
                                    End If
                                End If
                            
                            ElseIf targetType = 2 Then 'npc
                                
                                If Target > 0 Then
                                    
                                    If MapNpc(mapnum).NPC(Target).num > 0 Then
                                        didwalk = False
                                        target_verify = True
                                        TargetY = MapNpc(mapnum).NPC(Target).y
                                        TargetX = MapNpc(mapnum).NPC(Target).x
                                    Else
                                        MapNpc(mapnum).NPC(x).targetType = 0 ' clear
                                        MapNpc(mapnum).NPC(x).Target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                
                                i = Int(Rnd * 5)
    
                                ' Lets move the npc
                                Select Case i
                                    Case 0
    
                                        ' Up
                                        If MapNpc(mapnum).NPC(x).y > TargetY And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_UP) Then
                                                Call NpcMove(mapnum, x, DIR_UP, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(mapnum).NPC(x).y < TargetY And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_DOWN) Then
                                                Call NpcMove(mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(mapnum).NPC(x).x > TargetX And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_LEFT) Then
                                                Call NpcMove(mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(mapnum).NPC(x).x < TargetX And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_RIGHT) Then
                                                Call NpcMove(mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                    Case 1
    
                                        ' Right
                                        If MapNpc(mapnum).NPC(x).x < TargetX And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_RIGHT) Then
                                                Call NpcMove(mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(mapnum).NPC(x).x > TargetX And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_LEFT) Then
                                                Call NpcMove(mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(mapnum).NPC(x).y < TargetY And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_DOWN) Then
                                                Call NpcMove(mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(mapnum).NPC(x).y > TargetY And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_UP) Then
                                                Call NpcMove(mapnum, x, DIR_UP, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                    Case 2
    
                                        ' Down
                                        If MapNpc(mapnum).NPC(x).y < TargetY And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_DOWN) Then
                                                Call NpcMove(mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(mapnum).NPC(x).y > TargetY And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_UP) Then
                                                Call NpcMove(mapnum, x, DIR_UP, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(mapnum).NPC(x).x < TargetX And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_RIGHT) Then
                                                Call NpcMove(mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(mapnum).NPC(x).x > TargetX And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_LEFT) Then
                                                Call NpcMove(mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                    Case 3
    
                                        ' Left
                                        If MapNpc(mapnum).NPC(x).x > TargetX And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_LEFT) Then
                                                Call NpcMove(mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(mapnum).NPC(x).x < TargetX And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_RIGHT) Then
                                                Call NpcMove(mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(mapnum).NPC(x).y > TargetY And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_UP) Then
                                                Call NpcMove(mapnum, x, DIR_UP, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(mapnum).NPC(x).y < TargetY And Not didwalk Then
                                            If CanNpcMove(mapnum, x, DIR_DOWN) Then
                                                Call NpcMove(mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
    
                                End Select
    
                                ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                If Not didwalk Then
                                    If MapNpc(mapnum).NPC(x).x - 1 = TargetX And MapNpc(mapnum).NPC(x).y = TargetY Then
                                        If MapNpc(mapnum).NPC(x).Dir <> DIR_LEFT Then
                                            Call NpcDir(mapnum, x, DIR_LEFT)
                                        End If
    
                                        didwalk = True
                                    End If
    
                                    If MapNpc(mapnum).NPC(x).x + 1 = TargetX And MapNpc(mapnum).NPC(x).y = TargetY Then
                                        If MapNpc(mapnum).NPC(x).Dir <> DIR_RIGHT Then
                                            Call NpcDir(mapnum, x, DIR_RIGHT)
                                        End If
    
                                        didwalk = True
                                    End If
    
                                    If MapNpc(mapnum).NPC(x).x = TargetX And MapNpc(mapnum).NPC(x).y - 1 = TargetY Then
                                        If MapNpc(mapnum).NPC(x).Dir <> DIR_UP Then
                                            Call NpcDir(mapnum, x, DIR_UP)
                                        End If
    
                                        didwalk = True
                                    End If
    
                                    If MapNpc(mapnum).NPC(x).x = TargetX And MapNpc(mapnum).NPC(x).y + 1 = TargetY Then
                                        If MapNpc(mapnum).NPC(x).Dir <> DIR_DOWN Then
                                            Call NpcDir(mapnum, x, DIR_DOWN)
                                        End If
    
                                        didwalk = True
                                    End If
    
                                    ' We could not move so Target must be behind something, walk randomly.
                                    If Not didwalk Then
                                        i = Int(Rnd * 2)
    
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
    
                                            If CanNpcMove(mapnum, x, i) Then
                                                Call NpcMove(mapnum, x, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
    
                            Else
                                i = Int(Rnd * 4)
    
                                If i = 1 Then
                                    i = Int(Rnd * 4)
    
                                    If CanNpcMove(mapnum, x, i) Then
                                        Call NpcMove(mapnum, x, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        'End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).NPC(x) > 0 And MapNpc(mapnum).NPC(x).num > 0 Then
                    
                    Target = MapNpc(mapnum).NPC(x).Target
                    targetType = MapNpc(mapnum).NPC(x).targetType

                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                    
                        If targetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = mapnum Then
                                If NPC(MapNpc(mapnum).NPC(x).num).Behaviour = NPC_BEHAVIOUR_BOSS Then
                                    Call BossLogic(Target, NPC(MapNpc(mapnum).NPC(x).num).BossNum, MapNpc(mapnum).NPC(x).num)
                                Else
                                    TryNpcAttackPlayer x, Target
                                End If
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(mapnum).NPC(x).Target = 0
                                MapNpc(mapnum).NPC(x).targetType = 0 ' clear
                            End If
                        
                        ElseIf targetType = 2 Then
                            
                            ' lol no npc combat :( DATS WAT YOU THINK
                            If CanNpcAttackNpc(mapnum, x, MapNpc(mapnum).NPC(x).Target) = True Then
                                Call NpcAttackNpc(mapnum, x, MapNpc(mapnum).NPC(x).Target, NPC(Map(mapnum).NPC(x)).Damage)
                            End If
                        Else
                            ' Out of rang target?
                        End If
                    
                    End If
                                        
                        ' Spell Casting
                    ' For i = 1 To MAX_NPC_SPELLS
                        
                        sp = rand(1, MAX_NPC_SPELLS)
                        
                        If NPC(npcNum).Spell(sp) > 0 Then
                            If MapNpc(mapnum).NPC(x).SpellTimer(sp) + (Spell(NPC(npcNum).Spell(sp)).CastTime * 100) < GetTickCount Then
                                NpcSpellPlayer x, Target, sp
                            End If
                        End If
                     'Next
                    
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(mapnum).NPC(x).stopRegen Then
                    If MapNpc(mapnum).NPC(x).num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(mapnum).NPC(x).Vital(Vitals.HP) > 0 Then
                            
                            If MapNpc(mapnum).NPC(x).Vital(Vitals.HP) < GetNpcMaxVital(npcNum, Vitals.HP) Then
                                MapNpc(mapnum).NPC(x).Vital(Vitals.HP) = MapNpc(mapnum).NPC(x).Vital(Vitals.HP) + GetNpcVitalRegen(npcNum, Vitals.HP)
                                Call SendAnimation(mapnum, REGENHP_ANIM, 0, 0, TARGET_TYPE_NPC, x)
                                SendActionMsg mapnum, "+" & GetNpcVitalRegen(npcNum, Vitals.HP), BrightGreen, 1, (MapNpc(mapnum).NPC(x).x * 32), (MapNpc(mapnum).NPC(x).y * 32)
                                ' แก้บัค ไม่อัพเดทเลือด npc ตอนรีเจน
                                SendMapNpcVitals mapnum, x
                            End If
                            
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(mapnum).NPC(x).Vital(Vitals.HP) > GetNpcMaxVital(npcNum, Vitals.HP) Then
                                MapNpc(mapnum).NPC(x).Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
                                ' แก้บัค ไม่อัพเดทเลือด npc ตอนรีเจน
                                SendMapNpcVitals mapnum, x
                            End If
                            
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' ตรวจสอบว่า Npc ตายหรือยัง?
                If MapNpc(mapnum).NPC(x).num > 0 Then
                    If MapNpc(mapnum).NPC(x).Vital(HP) <= 0 Then ' Hp < 0
                        MapNpc(mapnum).NPC(x).num = 0
                        MapNpc(mapnum).NPC(x).SpawnWait = TickCount
                   End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(mapnum).NPC(x).num = 0 And Map(mapnum).NPC(x) > 0 Then
                    If TickCount > MapNpc(mapnum).NPC(x).SpawnWait + (NPC(Map(mapnum).NPC(x)).SpawnSecs * 1000) Then
                        Call SpawnNpc(x, mapnum)
                    End If
                End If

            Next

        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

End Sub

Private Sub UpdateNpcAttack()

Dim i As Long, x As Long, mapnum As Long, n As Long, x1 As Long, y1 As Long


End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not TempPlayer(i).stopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, Vitals.HP)
                    ' Say regen hp
                    SendActionMsg GetPlayerMap(i), "+" & GetPlayerVitalRegen(i, Vitals.HP), BrightGreen, 1, (Player(i).x * 32), (Player(i).y * 32)
                    ' อินเมชั่นตอนรีเจน
                    Call SendAnimation(GetPlayerMap(i), REGENHP_ANIM, 0, 0, TARGET_TYPE_PLAYER, i)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
    
                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, Vitals.MP)
                    ' Say regen mp
                    SendActionMsg GetPlayerMap(i), "+" & GetPlayerVitalRegen(i, Vitals.MP), Blue, 1, (Player(i).x * 32), (Player(i).y * 32)
                    ' อินเมชั่นตอนรีเจน
                    Call SendAnimation(GetPlayerMap(i), REGENMP_ANIM, 0, 0, TARGET_TYPE_PLAYER, i)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdatePlayerVital(ByVal index As Long)

        If IsPlaying(index) Then
            Call SendVital(index, Vitals.HP)
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        End If

End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("บันทึกผู้เล่นที่กำลังออนไลน์ทั้งหมด...")

        For i = 1 To Player_HighIndex

            If IsPlaying(i) Then
                Call SavePlayer(i)
                Call SaveBank(i)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("เซิฟเวอร์จะปิดในอีก " & Secs & " วินาที.", BrightBlue)
        Call TextAdd("เซิฟเวอร์จะปิดแบบอัตโนมัตในอีก " & Secs & " วินาที.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub

Function CanEventMoveTowardsPlayer(playerID As Long, mapnum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.
    CanEventMoveTowardsPlayer = 5
    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    WalkThrough = Map(mapnum).Events(TempPlayer(playerID).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(playerID).EventMap.EventPages(eventID).pageID).WalkThrough
    
    i = Int(Rnd * 5)
    didwalk = False
    
    ' Lets move the event
    Select Case i
        Case 0
    
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveTowardsPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveTowardsPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveTowardsPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveTowardsPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 1
        
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveTowardsPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveTowardsPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveTowardsPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveTowardsPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 2
        
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveTowardsPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveTowardsPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveTowardsPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveTowardsPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 3
        
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveTowardsPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveTowardsPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveTowardsPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveTowardsPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
    
        End Select
        
        CanEventMoveTowardsPlayer = Random(0, 3)
End Function

Function CanEventMoveAwayFromPlayer(playerID As Long, mapnum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.
    CanEventMoveAwayFromPlayer = 5
    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    WalkThrough = Map(mapnum).Events(TempPlayer(playerID).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(playerID).EventMap.EventPages(eventID).pageID).WalkThrough
    
    i = Int(Rnd * 5)
    didwalk = False
    
    ' Lets move the event
    Select Case i
        Case 0
    
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 1
        
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 2
        
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 3
        
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
    
        End Select
        
        CanEventMoveAwayFromPlayer = Random(0, 3)
End Function

Function GetDirToPlayer(playerID As Long, mapnum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long, distance As Long
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.
    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    
    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_RIGHT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_LEFT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If y - y1 > 0 Then
        If y - y1 > distance Then
            i = DIR_DOWN
            distance = y - y1
        End If
    ElseIf y - y1 < 0 Then
        If ((y - y1) * -1) > distance Then
            i = DIR_UP
            distance = ((y - y1) * -1)
        End If
    End If
    
    GetDirToPlayer = i
    
End Function

Function GetDirAwayFromPlayer(playerID As Long, mapnum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long, distance As Long
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.
    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    
    
    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_LEFT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_RIGHT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If y - y1 > 0 Then
        If y - y1 > distance Then
            i = DIR_UP
            distance = y - y1
        End If
    ElseIf y - y1 < 0 Then
        If ((y - y1) * -1) > distance Then
            i = DIR_DOWN
            distance = ((y - y1) * -1)
        End If
    End If
    
    GetDirAwayFromPlayer = i
End Function
