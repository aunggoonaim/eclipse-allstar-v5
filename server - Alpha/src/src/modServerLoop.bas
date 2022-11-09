Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim BuffTimer As Long
    Dim I As Long, x As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long
    Dim mapNum, x1, y1, Anim, playerNumber As Long
    ServerOnline = True

    Do While ServerOnline
        Tick = GetTickCount
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        If Tick > tmr25 Then
            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(I).spellBuffer.Spell > 0 Then
                        If GetTickCount > TempPlayer(I).spellBuffer.Timer + (Spell(Player(I).Spell(TempPlayer(I).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell I, TempPlayer(I).spellBuffer.Spell, TempPlayer(I).spellBuffer.target, TempPlayer(I).spellBuffer.tType
                            TempPlayer(I).spellBuffer.Spell = 0
                            TempPlayer(I).spellBuffer.Timer = 0
                            TempPlayer(I).spellBuffer.target = 0
                            TempPlayer(I).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(I).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(I).StunTimer + (TempPlayer(I).StunDuration * 1000) Then
                            TempPlayer(I).StunDuration = 0
                            TempPlayer(I).StunTimer = 0
                            SendStunned I
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(I).stopRegen Then
                        If TempPlayer(I).stopRegenTimer + 5000 < GetTickCount Then
                            TempPlayer(I).stopRegen = False
                            TempPlayer(I).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For x = 1 To MAX_DOTS
                        HandleDoT_Player I, x
                        HandleHoT_Player I, x
                    Next
                End If
            Next
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            tmr25 = GetTickCount + 25
        End If
If Tick > BuffTimer Then
            For I = 1 To Player_HighIndex
                For x = 1 To 10
                    If TempPlayer(I).BuffTimer(x) > 0 Then
                        TempPlayer(I).BuffTimer(x) = TempPlayer(I).BuffTimer(x) - 1
                        If TempPlayer(I).BuffTimer(x) = 0 Then
                            TempPlayer(I).Buffs(x) = 0
                        End If
                    End If
                Next
            Next
            BuffTimer = GetTickCount + 1000
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For I = 1 To MAX_PLAYERS
                If frmServer.Socket(I).State > sckConnected Then
                    Call CloseSocket(I)
                End If
            Next
            UpdateMapLogic
            tmr500 = GetTickCount + 500
        End If

        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            tmr1000 = GetTickCount + 1000
        End If
        
        ' projectiles
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                For x = 1 To MAX_PLAYER_PROJECTILES
                    If TempPlayer(I).ProjecTile(x).Pic > 0 Then
                        ' handle the projec tile
                        HandleProjecTile I, x
                    End If
                Next
            End If
        Next

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If
           ' Anim every 1 sec
           If Tick > Anim Then
            Anim = GetTickCount + 1000
            For mapNum = 1 To MAX_MAPS
                For x1 = 0 To Map(mapNum).MaxX
                    For y1 = 0 To Map(mapNum).MaxY
                        If Map(mapNum).Tile(x1, y1).Type = TILE_TYPE_ANIMATION Then
                            For playerNumber = 1 To Player_HighIndex
                                If IsPlaying(playerNumber) Then
                                    If GetPlayerMap(playerNumber) = mapNum Then
                                        SendAnimation mapNum, Map(mapNum).Tile(x1, y1).Data1, x1, y1
                                    End If
                                End If
                            Next
                        End If
                    Next
                Next
            Next
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
        
                'Handles Guild Invites
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                If TempPlayer(I).tmpGuildInviteSlot > 0 Then
                    If Tick > TempPlayer(I).tmpGuildInviteTimer Then
                        If GuildData(TempPlayer(I).tmpGuildInviteSlot).In_Use = True Then
                            PlayerMsg I, "Time ran out to join " & GuildData(TempPlayer(I).tmpGuildInviteSlot).Guild_Name & ".", BrightRed
                            TempPlayer(I).tmpGuildInviteSlot = 0
                            TempPlayer(I).tmpGuildInviteTimer = 0
                        Else
                            'Just remove this guild has been unloaded
                            TempPlayer(I).tmpGuildInviteSlot = 0
                            TempPlayer(I).tmpGuildInviteTimer = 0
                        End If
                    End If
                End If
            End If
        Next I

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
    Dim I As Long, x As Long, mapNum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, npcNum As Long
    Dim target As Long, targetType As Byte, DidWalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim TargetX As Long, TargetY As Long, target_verify As Boolean

    For mapNum = 1 To MAX_MAPS
        ' items appearing to everyone
        For I = 1 To MAX_MAP_ITEMS
            If MapItem(mapNum, I).Num > 0 Then
                If MapItem(mapNum, I).playerName <> vbNullString Then
                    ' make item public?
                    If MapItem(mapNum, I).playerTimer < GetTickCount Then
                        ' make it public
                        MapItem(mapNum, I).playerName = vbNullString
                        MapItem(mapNum, I).playerTimer = 0
                        ' send updates to everyone
                        SendMapItemsToAll mapNum
                    End If
                    ' despawn item?
                    If MapItem(mapNum, I).canDespawn Then
                        If MapItem(mapNum, I).despawnTimer < GetTickCount Then
                            ' despawn it
                            ClearMapItem I, mapNum
                            ' send updates to everyone
                            SendMapItemsToAll mapNum
                        End If
                    End If
                End If
            End If
        Next
        
        '  Close the doors
        If TickCount > TempTile(mapNum).DoorTimer + 5000 Then
            For x1 = 0 To Map(mapNum).MaxX
                For y1 = 0 To Map(mapNum).MaxY
                    If Map(mapNum).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(mapNum).DoorOpen(x1, y1) = YES Then
                        TempTile(mapNum).DoorOpen(x1, y1) = NO
                        SendMapKeyToMap mapNum, x1, y1, 0
                    End If
                Next
            Next
        End If
        
        ' check for DoTs + hots
        For I = 1 To MAX_MAP_NPCS
            If MapNpc(mapNum).NPC(I).Num > 0 Then
                For x = 1 To MAX_DOTS
                    HandleDoT_Npc mapNum, I, x
                    HandleHoT_Npc mapNum, I, x
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(mapNum).Resource_Count > 0 Then
            For I = 0 To ResourceCache(mapNum).Resource_Count
                Resource_index = Map(mapNum).Tile(ResourceCache(mapNum).ResourceData(I).x, ResourceCache(mapNum).ResourceData(I).y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(mapNum).ResourceData(I).ResourceState = 1 Or ResourceCache(mapNum).ResourceData(I).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(mapNum).ResourceData(I).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(mapNum).ResourceData(I).ResourceTimer = GetTickCount
                            ResourceCache(mapNum).ResourceData(I).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(mapNum).ResourceData(I).cur_health = Resource(Resource_index).health
                            SendResourceCacheToMap mapNum, I
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(mapNum) = YES Then
            TickCount = GetTickCount
            
            For x = 1 To MAX_MAP_NPCS
                npcNum = MapNpc(mapNum).NPC(x).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapNum).NPC(x) > 0 And MapNpc(mapNum).NPC(x).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(mapNum).NPC(x).StunDuration > 0 Then
    
                            For I = 1 To Player_HighIndex
                                If IsPlaying(I) Then
                                    If GetPlayerMap(I) = mapNum And MapNpc(mapNum).NPC(x).target = 0 And GetPlayerAccess(I) <= ADMIN_MONITOR Then
                                        n = NPC(npcNum).Range
                                        DistanceX = MapNpc(mapNum).NPC(x).x - GetPlayerX(I)
                                        DistanceY = MapNpc(mapNum).NPC(x).y - GetPlayerY(I)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(I) = YES Then
                                                If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                                    Call PlayerMsg(I, Trim$(NPC(npcNum).Name) & " says: " & Trim$(NPC(npcNum).AttackSay), SayColor)
                                                End If
                                                MapNpc(mapNum).NPC(x).targetType = 1 ' player
                                                MapNpc(mapNum).NPC(x).target = I
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
                If Map(mapNum).NPC(x) > 0 And MapNpc(mapNum).NPC(x).Num > 0 Then
                    If MapNpc(mapNum).NPC(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(mapNum).NPC(x).StunTimer + (MapNpc(mapNum).NPC(x).StunDuration * 1000) Then
                            MapNpc(mapNum).NPC(x).StunDuration = 0
                            MapNpc(mapNum).NPC(x).StunTimer = 0
                        End If
                    Else
                            
                        target = MapNpc(mapNum).NPC(x).target
                        targetType = MapNpc(mapNum).NPC(x).targetType
    
                        ' Check to see if its time for the npc to walk
                        If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        
                            If targetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(target) And GetPlayerMap(target) = mapNum Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = GetPlayerY(target)
                                        TargetX = GetPlayerX(target)
                                    Else
                                        MapNpc(mapNum).NPC(x).targetType = 0 ' clear
                                        MapNpc(mapNum).NPC(x).target = 0
                                    End If
                                End If
                            
                            ElseIf targetType = 2 Then 'npc
                                
                                If target > 0 Then
                                    
                                    If MapNpc(mapNum).NPC(target).Num > 0 Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = MapNpc(mapNum).NPC(target).y
                                        TargetX = MapNpc(mapNum).NPC(target).x
                                    Else
                                        MapNpc(mapNum).NPC(x).targetType = 0 ' clear
                                        MapNpc(mapNum).NPC(x).target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                
                                I = Int(Rnd * 5)
    
                                ' Lets move the npc
                                Select Case I
                                    Case 0
    
                                        ' Up
                                        If MapNpc(mapNum).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_UP) Then
                                                Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(mapNum).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_DOWN) Then
                                                Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(mapNum).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_LEFT) Then
                                                Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(mapNum).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                                                Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 1
    
                                        ' Right
                                        If MapNpc(mapNum).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                                                Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(mapNum).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_LEFT) Then
                                                Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(mapNum).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_DOWN) Then
                                                Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(mapNum).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_UP) Then
                                                Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 2
    
                                        ' Down
                                        If MapNpc(mapNum).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_DOWN) Then
                                                Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(mapNum).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_UP) Then
                                                Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(mapNum).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                                                Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(mapNum).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_LEFT) Then
                                                Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 3
    
                                        ' Left
                                        If MapNpc(mapNum).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_LEFT) Then
                                                Call NpcMove(mapNum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(mapNum).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_RIGHT) Then
                                                Call NpcMove(mapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(mapNum).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_UP) Then
                                                Call NpcMove(mapNum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(mapNum).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(mapNum, x, DIR_DOWN) Then
                                                Call NpcMove(mapNum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                End Select
    
                                ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(mapNum).NPC(x).x - 1 = TargetX And MapNpc(mapNum).NPC(x).y = TargetY Then
                                        If MapNpc(mapNum).NPC(x).Dir <> DIR_LEFT Then
                                            Call NpcDir(mapNum, x, DIR_LEFT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(mapNum).NPC(x).x + 1 = TargetX And MapNpc(mapNum).NPC(x).y = TargetY Then
                                        If MapNpc(mapNum).NPC(x).Dir <> DIR_RIGHT Then
                                            Call NpcDir(mapNum, x, DIR_RIGHT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(mapNum).NPC(x).x = TargetX And MapNpc(mapNum).NPC(x).y - 1 = TargetY Then
                                        If MapNpc(mapNum).NPC(x).Dir <> DIR_UP Then
                                            Call NpcDir(mapNum, x, DIR_UP)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(mapNum).NPC(x).x = TargetX And MapNpc(mapNum).NPC(x).y + 1 = TargetY Then
                                        If MapNpc(mapNum).NPC(x).Dir <> DIR_DOWN Then
                                            Call NpcDir(mapNum, x, DIR_DOWN)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    ' We could not move so Target must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        I = Int(Rnd * 2)
    
                                        If I = 1 Then
                                            I = Int(Rnd * 4)
    
                                            If CanNpcMove(mapNum, x, I) Then
                                                Call NpcMove(mapNum, x, I, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
    
                            Else
                                I = Int(Rnd * 4)
    
                                If I = 1 Then
                                    I = Int(Rnd * 4)
    
                                    If CanNpcMove(mapNum, x, I) Then
                                        Call NpcMove(mapNum, x, I, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapNum).NPC(x) > 0 And MapNpc(mapNum).NPC(x).Num > 0 Then
                    target = MapNpc(mapNum).NPC(x).target
                    targetType = MapNpc(mapNum).NPC(x).targetType

                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                    
                        If targetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = mapNum Then
                                TryNpcAttackPlayer x, target
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(mapNum).NPC(x).target = 0
                                MapNpc(mapNum).NPC(x).targetType = 0 ' clear
                            End If
                        ElseIf targetType = 2 Then
                            ' lol no npc combat :( DATS WAT YOU THINK
                            If CanNpcAttackNpc(mapNum, x, MapNpc(mapNum).NPC(x).target) = True Then
                                Call NpcAttackNpc(mapNum, x, MapNpc(mapNum).NPC(x).target, NPC(Map(mapNum).NPC(x)).Damage)
                            End If
                        Else
                        
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(mapNum).NPC(x).stopRegen Then
                    If MapNpc(mapNum).NPC(x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(mapNum).NPC(x).Vital(Vitals.HP) > 0 Then
                            MapNpc(mapNum).NPC(x).Vital(Vitals.HP) = MapNpc(mapNum).NPC(x).Vital(Vitals.HP) + GetNpcVitalRegen(npcNum, Vitals.HP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(mapNum).NPC(x).Vital(Vitals.HP) > GetNpcMaxVital(npcNum, Vitals.HP) Then
                                MapNpc(mapNum).NPC(x).Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(mapNum).NPC(x).Num = 0 And Map(mapNum).NPC(x) > 0 Then
                    If TickCount > MapNpc(mapNum).NPC(x).SpawnWait + (NPC(Map(mapNum).NPC(x)).SpawnSecs * 1000) Then
                        Call SpawnNpc(x, mapNum)
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

Private Sub UpdatePlayerVitals()
Dim I As Long
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If Not TempPlayer(I).stopRegen Then
                If GetPlayerVital(I, Vitals.HP) <> GetPlayerMaxVital(I, Vitals.HP) Then
                    Call SetPlayerVital(I, Vitals.HP, GetPlayerVital(I, Vitals.HP) + GetPlayerVitalRegen(I, Vitals.HP))
                    Call SendVital(I, Vitals.HP)
                    ' send vitals to party if in one
                    If TempPlayer(I).inParty > 0 Then SendPartyVitals TempPlayer(I).inParty, I
                End If
    
                If GetPlayerVital(I, Vitals.MP) <> GetPlayerMaxVital(I, Vitals.MP) Then
                    Call SetPlayerVital(I, Vitals.MP, GetPlayerVital(I, Vitals.MP) + GetPlayerVitalRegen(I, Vitals.MP))
                    Call SendVital(I, Vitals.MP)
                    ' send vitals to party if in one
                    If TempPlayer(I).inParty > 0 Then SendPartyVitals TempPlayer(I).inParty, I
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim I As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...")

        For I = 1 To Player_HighIndex

            If IsPlaying(I) Then
                Call SavePlayer(I)
                Call SaveBank(I)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub
