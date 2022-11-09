Attribute VB_Name = "modPlayer"
Option Explicit
Dim Craft(0)

Sub HandleUseChar(ByVal index As Long)
    If Not IsPlaying(index) Then
        Call JoinGame(index)
        Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " ได้เริ่มเล่นเกม " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " ได้เริ่มเข้าสู่เกม " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long)
    Dim i As Long, j As Long
    Dim EXPRATE As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)
    
    ' send the login ok
    SendLoginOk index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call SendPlayerSpells(index)
    Call SendHotbar(index)
    Call SendQuests(index)
    Call SendDoors(index)
    Call SendSwitchesAndVariables(index)
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    
    SendEXP index
    Call SendStats(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    'View Current Pets on Map
    If PetMapCache(Player(index).Map).UpperBound > 0 Then
        For j = 1 To PetMapCache(Player(index).Map).UpperBound
            Call NPCCache_Create(index, Player(index).Map, PetMapCache(Player(index).Map).Pet(j))
        Next
    End If
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(index) & " ได้เข้าสู่" & Options.Game_Name & " แล้ว !", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(index) & "[GM] ได้เข้าสู่ " & Options.Game_Name & " แล้ว !", Yellow)
    End If
    
    If Player(index).WieldDagger > 0 Then
        SetPlayerEquipment index, Player(index).WieldDagger, Shield
        Call SendWornEquipment(index)
    End If
    
    ' Send welcome messages
    Call SendWelcome(index)
    
        ' Do all the guild start up checks
    Call GuildLoginCheck(index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next
    
    ' Send the flag so they know they can start doing stuff
    SendInGame index
    Player(index).message = ""
    Player(index).Killer = 0
    
    ' fixed buff
    For i = 1 To MAX_BUFF
        Player(index).BuffStatus(i) = 0
        Player(index).BuffTime(i) = 0
    Next
    
    ' fixed level spell
    For i = 1 To MAX_PLAYER_SPELLS
        If Player(index).skillLV(i) < 0 Then
            Player(index).skillLV(i) = 0
        End If
        
        If Player(index).skillLV(i) > MAX_SKILL_LEVEL Then
            Player(index).skillLV(i) = MAX_SKILL_LEVEL - 1
            Call SendPlayerData(index)
        End If
    Next
    
    If GetPlayerLevel(index) = MAX_LEVELS Then
        Call SetPlayerExp(index, 1)
        SendEXP index
    End If
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, START_MAP, START_X, START_Y)
    ' Fixed Dir player !
    ' 1 = up , 2 = down , 3 = left , 4 = right
    Call SetPlayerDir(index, 1)
    
    EXPRATE = frmServer.scrlExpRate.Value
    Call PlayerMsg(index, "เกมถูกตั้งให้ผู้เล่นได้รับ Exp " & EXPRATE * 100 & "% จากปกติ.", Yellow)
    Call PlayerMsg(index, "เกมถูกตั้งให้อัตรา Drop ไอเทมเป็น " & frmServer.scrlDropRate.Value * 100 & "% จากปกติ.", Yellow)
    
    ' Regen Full hp & mp
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    Call SendPlayerData(index)
    
End Sub

Sub LeftGame(ByVal index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False
        
        ' Loop through entire map and purge NPC from targets
        ' ทำให้ npc ที่โจมตีเรา ลืมเรา
        
        For i = 1 To MAX_MAP_NPCS
        
            If MapNpc(TempPlayer(index).OldMap).NPC(i).targetType = TARGET_TYPE_PLAYER Then
                If MapNpc(TempPlayer(index).OldMap).NPC(i).Target = index Then
                    ' Set NPC target to 0
                    MapNpc(TempPlayer(index).OldMap).NPC(i).Target = 0
                    MapNpc(TempPlayer(index).OldMap).NPC(i).targetType = 0
                    MapNpc(TempPlayer(index).OldMap).NPC(i).GetDamage = 0
                    SendTarget i
                End If
            End If
            
        Next

        ' Fixed pet bug
    If TempPlayer(index).havePet Then
        PetDisband index, GetPlayerMap(TempPlayer(index).OldMap)
        SendMap index, GetPlayerMap(TempPlayer(index).OldMap)
        PetDisband index, GetPlayerMap(index)
        SendMap index, GetPlayerMap(index)
    End If

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(index).InTrade > 0 Then
            tradeTarget = TempPlayer(index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " ได้ยกเลิกการแลกเปลี่ยน.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        If Player(index).WieldDagger > 0 Then
            If GetPlayerEquipment(index, Shield) = 0 Then
                Call SetPlayerdagger(index, 0)
            End If
        End If

For i = 1 To Player_HighIndex
     If IsPlaying(i) Then
          If GetPlayerMap(i) = GetPlayerMap(index) Then
               Call PlayerWarp(i, GetPlayerMap(index), GetPlayerX(i), GetPlayerY(i))
          End If
     End If
Next
        
        ' leave party.
        Party_PlayerLeave index
        
        ' clear target
For i = 1 To Player_HighIndex
    ' Prevent subscript out range
    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
        ' clear players target
        If TempPlayer(i).targetType = TARGET_TYPE_PLAYER And TempPlayer(i).Target = index Then
            TempPlayer(i).Target = 0
            TempPlayer(i).targetType = TARGET_TYPE_NONE
            SendTarget i
        End If
    End If
Next

        If Player(index).GuildFileId > 0 Then
            'Set player online flag off
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).Online = False
            Call CheckUnloadGuild(TempPlayer(index).tmpGuildSlot)
        End If

        ' save and clear data.
        Call SavePlayer(index)
        Call SaveBank(index)
        Call ClearBank(index)

        ' Send a global message that he/she left
        If GetPlayerAccess(index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(index) & " ได้ออกจากเกม " & Options.Game_Name & " แล้ว !", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(index) & " ได้ตัดการเชื่อมต่อจากเกม " & Options.Game_Name & " แล้ว !", Yellow)
        End If

        Call TextAdd(GetPlayerName(index) & " ได้ตัดการเชื่อมต่อจากเกม " & Options.Game_Name & ".")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If
    
    If frmServer.optReset.Value > 0 Then
        ' Delete names from master name file
        Player(index).Name = ""

        If LenB(Trim$(Player(index).Name)) > 0 Then
            Call DeleteName(Player(index).Name)
        End If
        
        Call SavePlayer(index)
        
        ' Everything went ok
        Call AddLog("ไอดี " & Trim$(Player(index).Login) & " ได้ถูกลบแล้ว.", ADMIN_LOG)
        Call TextAdd("ไอดี " & Trim$(Player(index).Login) & " ได้ถูกลบแล้ว.")
        Call Kill(App.Path & "\data\accounts\" & Trim$(GetPlayerLogin(index)) & ".bin")
    End If
    
    Call ClearPlayer(index)
    
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(index, Armor)
    Helm = GetPlayerEquipment(index, Helmet)
    GetPlayerProtection = (GetPlayerStat(index, Stats.Endurance))

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim rate As Long

    'If GetPlayerEquipment(index, Weapon) > 0 Then
        
      ' rate = (Rnd) * 2

      '  If n = 1 Then
            i = (GetPlayerStat(index, Stats.willpower) \ 2) + (GetPlayerLevel(index) \ 3)
            rate = Int(Rnd * 100) + 1

            If i > rate Then
                CanPlayerCriticalHit = True
            End If
            
       ' End If
        
    'End If

End Function

Sub PlayerWarp(ByVal index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim shopNum As Long
    ' Dim OldMap As Long
    Dim i As Long, n As Long, PartyNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(mapnum).MaxX Then x = Map(mapnum).MaxX
    If y > Map(mapnum).MaxY Then y = Map(mapnum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    If mapnum = GetPlayerMap(index) Then
        Call SendPlayerXYToMap(index)
    End If
    
    TempPlayer(index).EventProcessingCount = 0
    TempPlayer(index).EventMap.CurrentEvents = 0
    
    ' clear target
For i = 1 To Player_HighIndex
    ' Prevent subscript out range
    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
        If TempPlayer(i).targetType = TARGET_TYPE_PLAYER And TempPlayer(i).Target = index Then
            TempPlayer(i).Target = 0
            TempPlayer(i).targetType = TARGET_TYPE_NONE
            SendTarget i
        End If
    End If
Next
    
    ' clear target
    TempPlayer(index).Target = 0
    TempPlayer(index).targetType = TARGET_TYPE_NONE
    SendTarget index

    ' Save old map to send erase player data to
    TempPlayer(index).OldMap = GetPlayerMap(index)
    
' Check to see if its a Party Dungeon
If Map(mapnum).Moral = MAP_MORAL_PARTY_MAP Then

' Check to make sure the player is in a party. If not exit the sub so they dont change maps
If TempPlayer(index).inParty < 1 And GetPlayerAccess(index) = 0 Then
    Call PlayerMsg(index, "นี่คือแผนที่ดันเจี้ยน คุณต้องมีปาร์ตี้เพื่อเข้าดันเจี้ยนนี้.", BrightRed)
    Call PlayerWarp(index, START_MAP, GetPlayerX(index), GetPlayerY(index))
    Exit Sub
End If

End If

    If TempPlayer(index).OldMap <> mapnum Then
        Call SendLeaveMap(index, TempPlayer(index).OldMap)
    End If

    Call SetPlayerMap(index, mapnum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    
    'If 'refreshing' map
    If (TempPlayer(index).OldMap <> mapnum) And TempPlayer(index).TempPetSlot > 0 Then
        'switch maps
        PetDisband index, TempPlayer(index).OldMap
       
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = TempPlayer(index).OldMap Then
                Call PlayerWarp(i, TempPlayer(index).OldMap, GetPlayerX(i), GetPlayerY(i))
            End If
        End If
    Next
       
        'SpawnPet index, mapnum, Trim$(Player(index).Pet.SpriteNum)
        'PetFollowOwner index
    End If

    'View Current Pets on Map
    'For i = 1 To 10
    '    If PetMapCache(Player(index).Map).Pet(i) > 0 Then
    '        Call NPCCache_Create(index, Player(index).Map, PetMapCache(Player(index).Map).Pet(i))
    '    End If
    'Next
    
    ' send player's equipment to new map
    SendMapEquipment index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(mapnum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = mapnum Then
                    SendMapEquipmentTo i, index
                End If
            End If
        Next
    End If
    
        ' Fix bug party hp bar
    
        PartyNum = TempPlayer(index).inParty
        
        If TempPlayer(index).inParty > 0 Then
            For i = 1 To MAX_PARTY_MEMBERS

                ' Recount party
                Party_CountMembers PartyNum
                
                ' Send update to all - including new player
                SendPartyUpdate PartyNum
                
                For n = 1 To MAX_PARTY_MEMBERS
                    SendPartyVitals PartyNum, n
                Next
            Next
        End If
        
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(TempPlayer(index).OldMap) = 0 Then
        PlayersOnMap(TempPlayer(index).OldMap) = NO

        ' ทำให้ npc ที่โจมตีเรา ลืมเรา
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(TempPlayer(index).OldMap).NPC(i).targetType = TARGET_TYPE_PLAYER Then
                If MapNpc(TempPlayer(index).OldMap).NPC(i).Target = index Then
                    ' Set NPC target to 0
                    MapNpc(TempPlayer(index).OldMap).NPC(i).Target = 0
                    MapNpc(TempPlayer(index).OldMap).NPC(i).targetType = 0
                    MapNpc(TempPlayer(index).OldMap).NPC(i).GetDamage = 0
                End If
            End If
        Next

    End If

    SendAnimation mapnum, WARP_ANIM, (Player(index).x), (Player(index).y)
    Call SendPlayerData(index)

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES
    TempPlayer(index).GettingMap = YES
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong mapnum
    Buffer.WriteLong Map(mapnum).Revision
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, mapnum As Long
    Dim x As Long, y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, Amount As Long, begineventprocessing As Boolean
    Dim DoorNum As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Moved = NO
    mapnum = GetPlayerMap(index)
    
    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "คุณจำเป็นต้องมีสิทธิชนิดของกุญแจสำคัญในการเปิดประตูนี้ (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "คุณจำเป็นต้องเปิดใช้งานสวิทช์เพื่อเปิดประตูนี้. ", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).WarpMap
                                                    x = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).WarpX
                                                    y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, x, y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                                    Call SetPlayerY(index, GetPlayerY(index) - 1)
                                    SendPlayerMove index, movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Up).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).Target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(mapnum).MaxY Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "คุณจำเป็นต้องมีสิทธิชนิดของกุญแจสำคัญในการเปิดประตูนี้ (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "คุณจำเป็นต้องเปิดใช้งานสวิทช์เพื่อเปิดประตูนี้. ", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).WarpMap
                                                    x = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).WarpX
                                                    y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, x, y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                                    Call SetPlayerY(index, GetPlayerY(index) + 1)
                                    SendPlayerMove index, movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Down).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).Target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "คุณจำเป็นต้องมีสิทธิชนิดของกุญแจสำคัญในการเปิดประตูนี้ (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "คุณจำเป็นต้องเปิดใช้งานสวิทช์เพื่อเปิดประตูนี้. ", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).WarpMap
                                                    x = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).WarpX
                                                    y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, x, y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                                    Call SetPlayerX(index, GetPlayerX(index) - 1)
                                    SendPlayerMove index, movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Left).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, NewMapX, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).Target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
            
        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < Map(mapnum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "คุณจำเป็นต้องมีสิทธิชนิดของกุญแจสำคัญในการเปิดประตูนี้ (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "คุณจำเป็นต้องเปิดใช้งานสวิทช์เพื่อเปิดประตูนี้. ", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).WarpMap
                                                    x = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).WarpX
                                                    y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, x, y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                                    Call SetPlayerX(index, GetPlayerX(index) + 1)
                                    SendPlayerMove index, movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Right).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).Target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
            
    End Select
    
    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(index, mapnum, x, y)
            Moved = YES
        End If
    
         ' Check to see if the tile is a door tile
        If .Type = TILE_TYPE_DOOR Then
            DoorNum = .Data1
            
            If Player(index).PlayerDoors(DoorNum).state = 1 Then
                mapnum = Doors(DoorNum).WarpMap
                x = Doors(DoorNum).WarpX
                y = Doors(DoorNum).WarpY
                Call PlayerWarp(index, mapnum, x, y)
                Moved = YES
            End If
            
        End If
    
        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            x = .Data1
            y = .Data2
    
            If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                SendMapKey index, x, y, 1
                Call MapMsg(GetPlayerMap(index), "ประตูถูกปลดล็อคแล้ว.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop index, x
                    TempPlayer(index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            TempPlayer(index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            Amount = .Data2
            If Not GetPlayerVital(index, VitalType) = GetPlayerMaxVital(index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = BrightGreen
                Else
                    Colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(index), "+" & Amount, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) + Amount
                PlayerMsg index, "คุณได้รับการรักษาจากพื้นที่พิเศษ.", BrightGreen
                Call SendVital(index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            Amount = .Data1
            SendActionMsg GetPlayerMap(index), "-" & Amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
            If GetPlayerVital(index, HP) - Amount <= 0 Then
                KillPlayer index
                PlayerMsg index, "คุณได้ถูกฆ่าโดย กับดัก.", BrightRed
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - Amount
                PlayerMsg index, "คุณกำลังได้รับบาดเจ็บจากกับดัก.", BrightRed
                Call SendVital(index, HP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
                        'Checks for sprite tile
        If .Type = TILE_TYPE_SPRITE Then
            Amount = .Data1
    If CheckCash(index, 1, 500) = True Then 'x=itemnum of your cash y=amount you want sprites to cost
         Call TakeInvItem(index, 1, 500) ' 32 is your currency(X), and 1 is the amount taken.(Y)
                   Call SetPlayerSprite(index, Amount)
                 Call SendPlayerData(index)
    Else
       Call PlayerMsg(index, "ต้องการเงิน 500 เพื่อเปลี่ยนตัวละคร !", Yellow)
    End If
         Else
            Moved = YES
        End If
        
        ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            ForcePlayerMove index, MOVING_WALKING, GetPlayerDir(index)
            Moved = YES
        End If
        
        ' Checkpoint
        If .Type = TILE_TYPE_CHECKPOINT Then
            SetCheckpoint index, .Data1, .Data2, .Data3
        End If
        
        ' craft
        If .Type = TILE_TYPE_CRAFT Then
        Call PlayerMsg(index, "คุณกำลังยืนอยู่ในเขตของเตาหลอม สามารถผลิตไอเทม/ตีบวกได้ !", White)
        Craft(0) = 1
        End If
        If Not .Type = TILE_TYPE_CRAFT Then
        Craft(0) = 0
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    End If

    x = GetPlayerX(index)
    y = GetPlayerY(index)
    
    If Moved = YES Then
        If TempPlayer(index).EventMap.CurrentEvents > 0 Then
            For i = 1 To TempPlayer(index).EventMap.CurrentEvents
                If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Global = 1 Then
                    If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).x = x And Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).y = y And Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).Trigger = 1 Then begineventprocessing = True
                Else
                    If TempPlayer(index).EventMap.EventPages(i).x = x And TempPlayer(index).EventMap.EventPages(i).y = y And Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).Trigger = 1 Then begineventprocessing = True
                End If
                If begineventprocessing = True Then
                    'Process this event, it is on-touch and everything checks out.
                    If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).CommandListCount > 0 Then
                        TempPlayer(index).EventProcessingCount = TempPlayer(index).EventProcessingCount + 1
                        ReDim Preserve TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).ActionTimer = GetTickCount
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).CurList = 1
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).CurSlot = 1
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).eventID = TempPlayer(index).EventMap.EventPages(i).eventID
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).pageID = TempPlayer(index).EventMap.EventPages(i).pageID
                        TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).WaitingForResponse = 0
                        ReDim TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount).ListLeftOff(0 To Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).CommandListCount)
                    End If
                    begineventprocessing = False
                End If
            Next
        End If
    End If

End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove index, Direction, movement, True
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim itemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(index, i)

        If itemNum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(itemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor

                    If Item(itemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, i
                Case Equipment.Helmet

                    If Item(itemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, i
                Case Equipment.Shield

                    If Item(itemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, i
            End Select

        Else
            SetPlayerEquipment index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal itemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = itemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal itemNum As Long) As Long
    Dim i As Long

    If Not IsPlaying(index) Then Exit Function
    If itemNum <= 0 Or itemNum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = itemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal index As Long, ByVal itemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemNum Then
            If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal index As Long, ByVal itemNum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemNum Then
            If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal index As Long, ByVal InvSlot As Byte, ByVal ItemVal As Long, Optional ByVal Update As Boolean = False) As Boolean
    Dim i As Long
    Dim n As Long
    Dim itemNum As Integer
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvSlot <= 0 Or InvSlot > MAX_ITEMS Then Exit Function
    
    itemNum = GetPlayerInvItemNum(index, InvSlot)

    ' Prevent subscript out of range
    If itemNum < 1 Then Exit Function
    
    If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, InvSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, InvSlot, GetPlayerInvItemValue(index, InvSlot) - ItemVal)
            
            ' Send the inventory update
            If Update Then
                Call SendInventoryUpdate(index, InvSlot)
            End If
            Exit Function
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, InvSlot, 0)
        Call SetPlayerInvItemValue(index, InvSlot, 0)
        
        ' Send the inventory update
        If Update Then
            Call SendInventoryUpdate(index, InvSlot)
        End If
    End If
End Function
Function GiveInvItem(ByVal index As Long, ByVal itemNum As Long, ByVal ItemVal As Long, Optional ByVal sendupdate As Boolean = True) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(index, itemNum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, itemNum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        If sendupdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "ช่องเก็บของเต็ม.", BrightRed)
        GiveInvItem = False
    End If

End Function

Function HasSpell(ByVal index As Long, ByVal spellnum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = spellnum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim mapnum As Long
    Dim Msg As String

    If Not IsPlaying(index) Then Exit Sub
    mapnum = GetPlayerMap(index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapnum, i).num > 0) And (MapItem(mapnum, i).num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, i).x = GetPlayerX(index)) Then
                    If (MapItem(mapnum, i).y = GetPlayerY(index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(index, MapItem(mapnum, i).num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(index, n, MapItem(mapnum, i).num)
    
                            If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(mapnum, i).Value)
                                Msg = MapItem(mapnum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            End If
    
                            ' Erase item from the map
                            ClearMapItem i, mapnum
                            
                            Call SendInventoryUpdate(index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                            SendActionMsg GetPlayerMap(index), Msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            Call CheckTasks(index, QUEST_TYPE_GOGATHER, GetItemNum(Trim$(Item(GetPlayerInvItemNum(index, n)).Name)))
                            Exit For
                        Else
                            Call PlayerMsg(index, "ช่องเก็บของเต็ม.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
Dim mapnum As Long

    mapnum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If MapItem(mapnum, mapItemNum).playerName = vbNullString Or MapItem(mapnum, mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal Amount As Long)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, invNum) > 0) Then
        If (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
            
            ' Work the Bind Type
            If Item(GetPlayerInvItemNum(index, invNum)).BindType = 1 Then Exit Sub
        
            i = FindOpenMapItemSlot(GetPlayerMap(index))

            If i <> 0 Then
                MapItem(GetPlayerMap(index), i).num = GetPlayerInvItemNum(index, invNum)
                MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
                MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
                MapItem(GetPlayerMap(index), i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(index), i).canDespawn = True
                MapItem(GetPlayerMap(index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME

                If Item(GetPlayerInvItemNum(index, invNum)).Type = ITEM_TYPE_CURRENCY Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(index, invNum) Then
                        MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, invNum)
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " ทิ้ง " & GetPlayerInvItemValue(index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(index, invNum, 0)
                        Call SetPlayerInvItemValue(index, invNum, 0)
                    Else
                        MapItem(GetPlayerMap(index), i).Value = Amount
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " ทิ้ง จำนวน " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(index, invNum, GetPlayerInvItemValue(index, invNum) - Amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), i).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " ทิ้ง " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, invNum, 0)
                    Call SetPlayerInvItemValue(index, invNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).num, Amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), MapItem(GetPlayerMap(index), i).canDespawn)
            Else
                Call PlayerMsg(index, "มีไอเทมบนพื้นมากเกินกำหนด.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Exit Sub
        End If
    
        ' Fixed point by Allstar
        Select Case GetPlayerLevel(index)
            Case 2 To 9: Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 2)
            Case 10 To 19: Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 3)
            Case 20 To 29: Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 4)
            Case 30 To 39: Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 6)
            Case 40 To 48: Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 8)
            Case 49 To 55: Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 10)
            Case 56 To 60: Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 15)
        Case Else
            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 5)
        End Select
        
        Call SetPlayerExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(index) & " ได้เลื่อน " & level_count & " เลเวล !", Yellow
            SendActionMsg GetPlayerMap(index), "Level Up !", BrightGreen, 1, (Player(index).x * 32), (Player(index).y * 32) - 16
            SendAnimation GetPlayerMap(index), LEVELUP_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
        Else
            'plural
            GlobalMsg GetPlayerName(index) & " ได้อัพเลเวลขึ้น " & level_count & " เลเวล !", Yellow
            SendActionMsg GetPlayerMap(index), "Level Ups !", BrightGreen, 1, (Player(index).x * 32), (Player(index).y * 32) - 16
            SendAnimation GetPlayerMap(index), LEVELUP_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
        End If
        SendEXP index
        SendPlayerData index
        Call PlayerMsg(index, "คุณมีแต้ม Status คงเหลือ " & GetPlayerPOINTS(index) & " พ้อย ค่ะ.", Yellow)
        
        ' Regen Full hp & mp
        Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
        Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
        Call SendVital(index, Vitals.HP)
        Call SendVital(index, Vitals.MP)
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seLevelUp, 1
    End If
End Sub

Sub CheckPlayerSkillUp(ByVal index As Long, ByVal spellslot As Long)
Dim i As Integer
    If Not Player(index).Spell(spellslot) > 0 Then Exit Sub

For i = 1 To 10
    If Player(index).skillEXP(spellslot) >= GetPlayerNextLevelSkill(index, spellslot) Then
        
        ' can level up?
        If Player(index).skillLV(spellslot) + 1 > MAX_SKILL_LEVEL Then
            Player(index).skillLV(spellslot) = MAX_SKILL_LEVEL - 1
            Exit Sub
        End If
    
        ' Fixed point by Allstar
        Player(index).skillEXP(spellslot) = Player(index).skillEXP(spellslot) - GetPlayerNextLevelSkill(index, spellslot)
        Player(index).skillLV(spellslot) = Player(index).skillLV(spellslot) + 1
        
        PlayerMsg index, "สกิล " & Trim(Spell(Player(index).Spell(spellslot)).Name) & " ได้อัพเลเวล !", BrightGreen
        SendActionMsg GetPlayerMap(index), "Skill Up !", BrightGreen, 1, (Player(index).x * 32), (Player(index).y * 32) - 32
        SendAnimation GetPlayerMap(index), LEVELUP_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seLevelUp, 1
    End If
Next

End Sub


' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim$(Player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(index).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
End Function

Function SetPlayerLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then Exit Function
    Player(index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long

If GetPlayerLevel(index) > 0 And GetPlayerLevel(index) <= MAX_LEVELS Then
    Select Case GetPlayerLevel(index)
    Case 1: GetPlayerNextLevel = 3
    Case 2: GetPlayerNextLevel = 8
    Case 3: GetPlayerNextLevel = 14
    Case 4: GetPlayerNextLevel = 25
    Case 5: GetPlayerNextLevel = 40
    Case 6: GetPlayerNextLevel = 60
    Case 7: GetPlayerNextLevel = 85
    Case 8: GetPlayerNextLevel = 120
    Case 9: GetPlayerNextLevel = 150
    Case 10: GetPlayerNextLevel = 200
    Case 11: GetPlayerNextLevel = 350
    Case 12: GetPlayerNextLevel = 470
    Case 13: GetPlayerNextLevel = 600
    Case 14: GetPlayerNextLevel = 810
    Case 15: GetPlayerNextLevel = 980
    Case 16: GetPlayerNextLevel = 1250
    Case 17: GetPlayerNextLevel = 1600
    Case 18: GetPlayerNextLevel = 2250
    Case 19: GetPlayerNextLevel = 3900
    Case 20: GetPlayerNextLevel = 5150
    Case 21: GetPlayerNextLevel = 7180
    Case 22: GetPlayerNextLevel = 8500
    Case 23: GetPlayerNextLevel = 10250
    Case 24: GetPlayerNextLevel = 13600
    Case 25: GetPlayerNextLevel = 17800
    Case 26: GetPlayerNextLevel = 26500
    Case 27: GetPlayerNextLevel = 38480
    Case 28: GetPlayerNextLevel = 53300
    Case 29: GetPlayerNextLevel = 74870
    Case 30: GetPlayerNextLevel = 111500
    Case 31: GetPlayerNextLevel = 154500
    Case 32: GetPlayerNextLevel = 231500
    Case 33: GetPlayerNextLevel = 354500
    Case 34: GetPlayerNextLevel = 531500
    Case 35: GetPlayerNextLevel = 724500
    Case 36: GetPlayerNextLevel = 911500
    Case 37: GetPlayerNextLevel = 1345000
    Case 38: GetPlayerNextLevel = 1715000
    Case 39: GetPlayerNextLevel = 2645000
    Case 40: GetPlayerNextLevel = 3915000
    Case 41: GetPlayerNextLevel = 5915000
    Case 42: GetPlayerNextLevel = 7915000
    Case 43: GetPlayerNextLevel = 9915000
    Case 44: GetPlayerNextLevel = 13150000
    Case 45: GetPlayerNextLevel = 171150000
    Case 46: GetPlayerNextLevel = 210000000
    Case 47: GetPlayerNextLevel = 279150000
    Case 48: GetPlayerNextLevel = 412100000
    Case 49: GetPlayerNextLevel = 590000000
    Case 50: GetPlayerNextLevel = 730000000
    Case 51: GetPlayerNextLevel = 1000000000
    Case 52: GetPlayerNextLevel = 1000000000
    Case 53: GetPlayerNextLevel = 1000000000
    Case 54: GetPlayerNextLevel = 1000000000
    Case 55: GetPlayerNextLevel = 1000000000
    Case 56: GetPlayerNextLevel = 2000000000
    Case 57: GetPlayerNextLevel = 2000000000
    Case 58: GetPlayerNextLevel = 2000000000
    Case 59: GetPlayerNextLevel = 2000000000
    Case 60: GetPlayerNextLevel = 1
    Case MAX_LEVELS: GetPlayerNextLevel = 1
    End Select
Else
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(index) + 1) ^ 3 - (6 * (GetPlayerLevel(index) + 1) ^ 2) + 17 * (GetPlayerLevel(index) + 1) - 12)
End If

End Function

Function GetPlayerNextLevelSkill(ByVal index As Long, ByVal spellslot As Long) As Long

If Player(index).skillLV(spellslot) > 0 And Player(index).skillLV(spellslot) <= MAX_SKILL_LEVEL Then
    Select Case Player(index).skillLV(spellslot)
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

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal exp As Long)
    Player(index).exp = exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(index).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Player(index).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    If GetPlayerVital(index, Vital) < 0 Then
        Player(index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal index As Long, ByVal stat As Stats) As Long
    Dim x As Long, i As Long
    If index > MAX_PLAYERS Then Exit Function
    
x = Player(index).stat(stat)
    
For i = 1 To Equipment.Equipment_Count - 1
    If Player(index).Equipment(i) > 0 Then
        
        ' ระบบแรร์ไอเทม
        ' If Item(Player(index).Equipment(i)).Add_Stat(stat) > 0 Then
            Select Case Item(Player(index).Equipment(i)).Rarity
             'Add Cases as levels of rarity you want.
            Case 0 'Item Normal
            x = x + Item(Player(index).Equipment(i)).Add_Stat(stat)
            Case 1 'Item Excellent. I will add 5 extra points to all
            x = x + Item(Player(index).Equipment(i)).Add_Stat(stat) + 5
            Case 2 'Item Master. I will add 15 extra points to all
            x = x + Item(Player(index).Equipment(i)).Add_Stat(stat) + 10
            Case 3 'Item Master. I will add 15 extra points to all
            x = x + Item(Player(index).Equipment(i)).Add_Stat(stat) + 15
            End Select
        ' End If
    End If
Next
    
    GetPlayerStat = x
End Function

Public Function GetPlayerRawStat(ByVal index As Long, ByVal stat As Stats) As Long
    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(index).stat(stat)
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal stat As Stats, ByVal Value As Long)
    Player(index).stat(stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    If POINTS <= 0 Then POINTS = 0
    Player(index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal mapnum As Long)

    If mapnum > 0 And mapnum <= MAX_MAPS Then
        Player(index).Map = mapnum
    End If

End Sub

Function GetPlayerX(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    Player(index).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    Player(index).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If InvSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(index).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal itemNum As Long)
    Player(index).Inv(InvSlot).num = itemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(index).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal spellslot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal spellslot As Long, ByVal spellnum As Long)
    Player(index).Spell(spellslot) = spellnum
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long

    If index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Player(index).Equipment(EquipmentSlot) = invNum
End Sub

' ToDo
Sub OnDeath(ByVal index As Long)
    Dim i As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)

    ' ไอเทมตก เมื่อตาย
        If GetPlayerEquipment(index, Armor) > 0 Then
            If Item(GetPlayerEquipment(index, Armor)).DropOnDeath > 0 Then
                Call PlayerMsg(index, "คุณได้สูญเสียไอเทม " & Trim(Item(GetPlayerEquipment(index, Armor)).Name) & " จากการตาย.", BrightRed)
                Call SpawnItem(GetPlayerEquipment(index, Armor), 1, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                Call PlayerMapDropItem(index, GetPlayerEquipment(index, Armor), 1)
                Call SetPlayerEquipment(index, 0, Armor)
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
            End If
        End If
        
        If GetPlayerEquipment(index, Helmet) > 0 Then
            If Item(GetPlayerEquipment(index, Helmet)).DropOnDeath > 0 Then
                Call PlayerMsg(index, "คุณได้สูญเสียไอเทม " & Trim(Item(GetPlayerEquipment(index, Helmet)).Name) & " จากการตาย.", BrightRed)
                Call SpawnItem(GetPlayerEquipment(index, Helmet), 1, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                Call PlayerMapDropItem(index, GetPlayerEquipment(index, Helmet), 1)
                Call SetPlayerEquipment(index, 0, Helmet)
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
            End If
        End If
        
        If GetPlayerEquipment(index, Shield) > 0 Then
            If Item(GetPlayerEquipment(index, Shield)).DropOnDeath > 0 Then
                Call PlayerMsg(index, "คุณได้สูญเสียไอเทม " & Trim(Item(GetPlayerEquipment(index, Shield)).Name) & " จากการตาย.", BrightRed)
                Call SpawnItem(GetPlayerEquipment(index, Shield), 1, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                Call PlayerMapDropItem(index, GetPlayerEquipment(index, Shield), 1)
                Call SetPlayerEquipment(index, 0, Shield)
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
            End If
        End If
        
        If GetPlayerEquipment(index, Weapon) > 0 Then
            If Item(GetPlayerEquipment(index, Weapon)).DropOnDeath > 0 Then
                Call PlayerMsg(index, "คุณได้สูญเสียไอเทม " & Trim(Item(GetPlayerEquipment(index, Weapon)).Name) & " จากการตาย.", BrightRed)
                Call SpawnItem(GetPlayerEquipment(index, Weapon), 1, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                Call PlayerMapDropItem(index, GetPlayerEquipment(index, Weapon), 1)
                Call SetPlayerEquipment(index, 0, Weapon)
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
            End If
        End If
    
    ' Loop through entire map and purge NPC from targets
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And IsConnected(i) Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then
                If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                    If TempPlayer(i).Target = index Then
                        TempPlayer(i).Target = 0
                        TempPlayer(i).targetType = TARGET_TYPE_NONE
                        SendTarget i
                    End If
                End If
            End If
        End If
    Next
    
    ' ทำให้ npc ที่โจมตีเรา ลืมเรา
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(index)).NPC(i).targetType = TARGET_TYPE_PLAYER Then
                If MapNpc(GetPlayerMap(index)).NPC(i).Target = index Then
                    ' Set NPC target to 0
                    MapNpc(GetPlayerMap(index)).NPC(i).Target = 0
                    MapNpc(GetPlayerMap(index)).NPC(i).targetType = 0
                    MapNpc(GetPlayerMap(index)).NPC(i).GetDamage = 0
                End If
            End If
        Next
    
    ' clear all buff
    For i = 1 To MAX_BUFF
        Player(index).BuffStatus(i) = 0
        Player(index).BuffTime(i) = 0
    Next
    
    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    Call PlayerWarp(index, START_MAP, START_X, START_Y)
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.Spell = 0
    TempPlayer(index).spellBuffer.Timer = 0
    TempPlayer(index).spellBuffer.Target = 0
    TempPlayer(index).spellBuffer.tType = 0
    
    Call SendClearSpellBuffer(index)
    
    TempPlayer(index).InBank = False
    TempPlayer(index).InShop = 0
    
    If TempPlayer(index).InTrade > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).num = 0
            TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Value = 0
        Next

        TempPlayer(index).InTrade = 0
        TempPlayer(TempPlayer(index).InTrade).InTrade = 0

        SendCloseTrade index
        SendCloseTrade TempPlayer(index).InTrade
    End If
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    Player(index).Killer = 0
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(index) = YES Then
        Call SetPlayerPK(index, NO)
    End If
    
    Call SendPlayerXY(index)
    Call SendPlayerData(index)

End Sub

Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    Dim ToolpowerReq As Long
    Dim Toolpower As Long
    
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(index)).Tile(x, y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count

            If ResourceCache(GetPlayerMap(index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(index)).ResourceData(i).y = y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Then

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg index, "คุณมีช่องว่างของกระเป๋าไม่เพียงพอที่จะเก็บของ.", BrightRed
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
                        rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y
                        
                        Damage = Item(GetPlayerEquipment(index, Weapon)).Toolpower
                        
                        ' check if tool power is strong enough
                        If (Resource(Resource_index).ToolpowerReq <= Item(GetPlayerEquipment(index, Weapon)).Toolpower) Then
                    
                        ' check if damage is more than health
                        If Damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                SendResourceCacheToMap GetPlayerMap(index), Resource_num
                                
                                If Random(1, 100) <= Resource(Resource_index).SuccessRate Then
                                    SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                    GiveInvItem index, Resource(Resource_index).ItemReward, 1
                                Else
                                    SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                End If
                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage
                                SendActionMsg GetPlayerMap(index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                            End If
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(index), "อ่อนหัด !", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                        Else
                        'Tool power too low
                        SendActionMsg GetPlayerMap(index), "เครื่องมืออ่อนแอเกินไป !", BrightRed, 1, (rX * 32), (rY * 32)
                        Call PlayerMsg(index, "Your " & Trim$(Item(GetPlayerEquipment(index, Weapon)).Name) & " isn't up to the task.", BrightRed)
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                        End If
                    End If

                Else
                    PlayerMsg index, "คุณใช้เครื่องมือผิดประเภท.", BrightRed
                End If

            Else
                PlayerMsg index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(index).Item(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal itemNum As Long)
    Bank(index).Item(BankSlot).num = itemNum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(index).Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank(index).Item(BankSlot).Value = ItemValue
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal InvSlot As Long, ByVal Amount As Long)
Dim BankSlot

    If InvSlot < 0 Or InvSlot > MAX_INV Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerInvItemValue(index, InvSlot) Then
        Exit Sub
    End If
    
    If Item(GetPlayerInvItemNum(index, InvSlot)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Then Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, InvSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(index, InvSlot)).Type = ITEM_TYPE_CURRENCY Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, InvSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), Amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, InvSlot))
                Call SetPlayerBankItemValue(index, BankSlot, Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, InvSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, InvSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
Dim InvSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerBankItemValue(index, BankSlot) Then
        Exit Sub
    End If
    
    If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Then Exit Sub
    End If
    
    InvSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If InvSlot > 0 Then
        If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), Amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - Amount)
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Public Sub KillPlayer(ByVal index As Long)
Dim exp As Long

    ' Calculate exp to give attacker
    exp = GetPlayerNextLevel(index) * 0.05

    ' Make sure we dont get less then 0
    If exp < 0 Then exp = 0
    
    If GetPlayerExp(index) < exp Then
        Call PlayerMsg(index, "คุณไม่เหลือค่าประสบการณ์จากการตาย.", BrightRed)
        Call SetPlayerExp(index, 0)
        SendEXP index
    Else
        Call SetPlayerExp(index, GetPlayerExp(index) - exp)
        Call PlayerMsg(index, "คุณได้สูญเสีย exp " & exp & " จากการตาย.", BrightRed)
        SendEXP index
    End If
    
    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, x As Long, y As Long, itemNum As Long
Dim Item1 As Long
Dim Item2 As Long
Dim Result As Long
Dim b As Long, j As Long
Dim randD As Long
Dim ClassReq As Long, Classchk As Boolean
Dim Classtxt As String

For j = 1 To MAX_INV
Next

b = FindOpenInvSlot(index, j)

Classchk = False

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(index, invNum) > 0) And (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(index, invNum)).Data2
        itemNum = GetPlayerInvItemNum(index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(itemNum).Type
        
            Case ITEM_TYPE_ARMOR
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "คุณมี Status ไม่เพียงพอในการใช้ไอเทมนี้.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "ต้องการเลเวล" & Item(itemNum).LevelReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement by Allstar Perfect
                
                ClassReq = Item(itemNum).ClassReq
    
                ' make sure the classreq > 0
                'If ClassReq > 0 Then ' 0 = no req
                '    If ClassReq <> GetPlayerClass(index) Then
                '        Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้ไอเทมนี้.", BrightRed)
                '        Exit Sub
                '    End If
                'End If
                
                Classtxt = "ต้องการอาชีพ : " ' Real
                    
                ' อาชีพ 1 ใช้ได้
                If Item(itemNum).ClassR1 > 0 Then
                    Classtxt = Classtxt & " มนุษย์,"
                    If GetPlayerClass(index) = 1 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 2 ใช้ได้
                If Item(itemNum).ClassR2 > 0 Then
                    Classtxt = Classtxt & " เอลฟ์,"
                    If GetPlayerClass(index) = 2 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 3 ใช้ได้
                If Item(itemNum).ClassR3 > 0 Then
                    Classtxt = Classtxt & " การ์เดี้ยน,"
                    If GetPlayerClass(index) = 3 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 4 ใช้ได้
                If Item(itemNum).ClassR4 > 0 Then
                    Classtxt = Classtxt & " เบอเซิร์ก,"
                    If GetPlayerClass(index) = 4 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 5 ใช้ได้
                If Item(itemNum).ClassR5 > 0 Then
                    Classtxt = Classtxt & " พาลาดิน,"
                    If GetPlayerClass(index) = 5 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 6 ใช้ได้
                If Item(itemNum).ClassR6 > 0 Then
                    Classtxt = Classtxt & " วิซาร์ด,"
                    If GetPlayerClass(index) = 6 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 7 ใช้ได้
                If Item(itemNum).ClassR7 > 0 Then
                    Classtxt = Classtxt & " ซามูไร,"
                    If GetPlayerClass(index) = 7 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 8 ใช้ได้
                If Item(itemNum).ClassR8 > 0 Then
                    Classtxt = Classtxt & " ฮันเตอร์,"
                    If GetPlayerClass(index) = 8 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 9 ใช้ได้
                If Item(itemNum).ClassR9 > 0 Then
                    Classtxt = Classtxt & " สไนเปอร์,"
                    If GetPlayerClass(index) = 9 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 10 ใช้ได้
                If Item(itemNum).ClassR10 > 0 Then
                    Classtxt = Classtxt & " แอสแซสซิน,"
                    If GetPlayerClass(index) = 10 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 11 ใช้ได้
                If Item(itemNum).ClassR11 > 0 Then
                    Classtxt = Classtxt & " ดาร์คลอร์ด,"
                    If GetPlayerClass(index) = 11 Then
                        Classchk = True
                    End If
                End If
                
                ' แจ้งข้อผิดพลาด
                If Item(itemNum).ClassR1 <= 0 And Item(itemNum).ClassR2 <= 0 And Item(itemNum).ClassR3 <= 0 And Item(itemNum).ClassR4 <= 0 And Item(itemNum).ClassR5 <= 0 And Item(itemNum).ClassR6 <= 0 And Item(itemNum).ClassR7 <= 0 And Item(itemNum).ClassR8 <= 0 And Item(itemNum).ClassR9 <= 0 And Item(itemNum).ClassR10 <= 0 And Item(itemNum).ClassR11 <= 0 Then
                    Call PlayerMsg(index, "ไม่มีอาชีพใดสามารถใช้ไอเทมนี้ได้.", BrightRed)
                    Exit Sub
                End If
                
                Classtxt = Classtxt & " ในการสวมใส่ไอเทมนี้."
                
                ' แจ้งข้อความว่าอาชีพใดใส่ได้บ้าง
                If Not Classchk = True Then
                    Call PlayerMsg(index, "อาชีพของคุณไม่สามารถสวมใส่ไอเทมนี้ได้.", BrightRed)
                    Call PlayerMsg(index, Classtxt, BrightRed)
                    Exit Sub
                End If
                
                ' แก้ไขบัคเพิ่มเติม Hp
                If GetPlayerMaxVital(index, HP) < Item(itemNum).HP Then
                    Call PlayerMsg(index, "ต้องการ MaxHp " & Item(itemNum).HP & " ในการสวมใส่ไอเทมนี้.", BrightRed)
                    Exit Sub
                End If
                
                ' แก้ไขบัคเพิ่มเติม Mp
                If GetPlayerMaxVital(index, MP) < Item(itemNum).MP Then
                    Call PlayerMsg(index, "ต้องการ MaxMp " & Item(itemNum).MP & " ในการสวมใส่ไอเทมนี้.", BrightRed)
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "ต้องการระดับ " & Item(itemNum).AccessReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(index, Armor)
                End If

                SetPlayerEquipment index, itemNum, Armor
                PlayerMsg index, "คุณได้สวมใส่ " & CheckGrammar(Item(itemNum).Name), BrightGreen
                TakeInvItem index, itemNum, 0

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            
            Case ITEM_TYPE_WEAPON
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "คุณมี Status ไม่เพียงพอในการใช้ไอเทมนี้.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "ต้องการเลเวล " & Item(itemNum).LevelReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                
                ClassReq = Item(itemNum).ClassReq
    
                ' make sure the classreq > 0
                'If ClassReq > 0 Then ' 0 = no req
                '    If ClassReq <> GetPlayerClass(index) Then
                '        Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้ไอเทมนี้.", BrightRed)
                '        Exit Sub
                '    End If
                'End If
                
                Classtxt = "ต้องการอาชีพ : " ' Real
                    
                ' อาชีพ 1 ใช้ได้
                If Item(itemNum).ClassR1 > 0 Then
                    Classtxt = Classtxt & " มนุษย์,"
                    If GetPlayerClass(index) = 1 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 2 ใช้ได้
                If Item(itemNum).ClassR2 > 0 Then
                    Classtxt = Classtxt & " เอลฟ์,"
                    If GetPlayerClass(index) = 2 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 3 ใช้ได้
                If Item(itemNum).ClassR3 > 0 Then
                    Classtxt = Classtxt & " การ์เดี้ยน,"
                    If GetPlayerClass(index) = 3 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 4 ใช้ได้
                If Item(itemNum).ClassR4 > 0 Then
                    Classtxt = Classtxt & " เบอเซิร์ก,"
                    If GetPlayerClass(index) = 4 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 5 ใช้ได้
                If Item(itemNum).ClassR5 > 0 Then
                    Classtxt = Classtxt & " พาลาดิน,"
                    If GetPlayerClass(index) = 5 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 6 ใช้ได้
                If Item(itemNum).ClassR6 > 0 Then
                    Classtxt = Classtxt & " วิซาร์ด,"
                    If GetPlayerClass(index) = 6 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 7 ใช้ได้
                If Item(itemNum).ClassR7 > 0 Then
                    Classtxt = Classtxt & " ซามูไร,"
                    If GetPlayerClass(index) = 7 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 8 ใช้ได้
                If Item(itemNum).ClassR8 > 0 Then
                    Classtxt = Classtxt & " ฮันเตอร์,"
                    If GetPlayerClass(index) = 8 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 9 ใช้ได้
                If Item(itemNum).ClassR9 > 0 Then
                    Classtxt = Classtxt & " สไนเปอร์,"
                    If GetPlayerClass(index) = 9 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 10 ใช้ได้
                If Item(itemNum).ClassR10 > 0 Then
                    Classtxt = Classtxt & " แอสแซสซิน,"
                    If GetPlayerClass(index) = 10 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 11 ใช้ได้
                If Item(itemNum).ClassR11 > 0 Then
                    Classtxt = Classtxt & " ดาร์คลอร์ด,"
                    If GetPlayerClass(index) = 11 Then
                        Classchk = True
                    End If
                End If
                
                ' แจ้งข้อผิดพลาด
                If Item(itemNum).ClassR1 <= 0 And Item(itemNum).ClassR2 <= 0 And Item(itemNum).ClassR3 <= 0 And Item(itemNum).ClassR4 <= 0 And Item(itemNum).ClassR5 <= 0 And Item(itemNum).ClassR6 <= 0 And Item(itemNum).ClassR7 <= 0 And Item(itemNum).ClassR8 <= 0 And Item(itemNum).ClassR9 <= 0 And Item(itemNum).ClassR10 <= 0 And Item(itemNum).ClassR11 <= 0 Then
                    Call PlayerMsg(index, "ไม่มีอาชีพใดสามารถใช้ไอเทมนี้ได้.", BrightRed)
                    Exit Sub
                End If
                
                Classtxt = Classtxt & " ในการสวมใส่ไอเทมนี้."
                
                ' แจ้งข้อความว่าอาชีพใดใส่ได้บ้าง
                If Not Classchk = True Then
                    Call PlayerMsg(index, "อาชีพของคุณไม่สามารถสวมใส่ไอเทมนี้ได้.", BrightRed)
                    Call PlayerMsg(index, Classtxt, BrightRed)
                    Exit Sub
                End If
                
                ' แก้ไขบัคเพิ่มเติม Hp
                If GetPlayerMaxVital(index, HP) < Item(itemNum).HP Then
                    Call PlayerMsg(index, "ต้องการ MaxHp " & Item(itemNum).HP & " ในการสวมใส่ไอเทมนี้.", BrightRed)
                    Exit Sub
                End If
                
                ' แก้ไขบัคเพิ่มเติม Mp
                If GetPlayerMaxVital(index, MP) < Item(itemNum).MP Then
                    Call PlayerMsg(index, "ต้องการ MaxMp " & Item(itemNum).MP & " ในการสวมใส่ไอเทมนี้.", BrightRed)
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "ต้องการระดับ " & Item(itemNum).AccessReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If

                'If GetPlayerEquipment(index, Weapon) > 0 Then
                    'tempItem = GetPlayerEquipment(index, Weapon)
                'End If

                'SetPlayerEquipment index, itemnum, Weapon
                
                If Item(itemNum).isTwohanded = True Then
                    If GetPlayerEquipment(index, Shield) > 0 Then
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                            If b < 1 Then
                                Call PlayerMsg(index, "ช่องเก็บของในกระเป๋าเต็ม !!", BrightRed)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                PlayerMsg index, "คุณได้สวมใส่ " & CheckGrammar(Item(itemNum).Name), BrightGreen
                TakeInvItem index, itemNum, 1

                If Item(itemNum).isDagger Then
                    If GetPlayerEquipment(index, Shield) > 0 Then
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                            If Item(GetPlayerEquipment(index, Weapon)).isDagger Then
                                GiveInvItem index, GetPlayerEquipment(index, Shield), 0
                                SetPlayerEquipment index, itemNum, Shield
                                Player(index).WieldDagger = itemNum
                            Else
                                GiveInvItem index, GetPlayerEquipment(index, Weapon), 0
                                SetPlayerEquipment index, itemNum, Weapon
                            End If
                        Else
                            SetPlayerEquipment index, itemNum, Weapon
                        End If
                    Else
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                            If Item(GetPlayerEquipment(index, Weapon)).isDagger Then
                                GiveInvItem index, GetPlayerEquipment(index, Shield), 0
                                SetPlayerEquipment index, itemNum, Shield
                                Player(index).WieldDagger = itemNum
                            Else
                                GiveInvItem index, GetPlayerEquipment(index, Weapon), 0
                                SetPlayerEquipment index, itemNum, Weapon
                            End If
                        Else
                            SetPlayerEquipment index, itemNum, Weapon
                        End If
                    End If
                    ElseIf Item(itemNum).isTwohanded Then
                        If GetPlayerEquipment(index, Shield) > 0 Then
                            If GetPlayerEquipment(index, Weapon) > 0 Then
                                GiveInvItem index, GetPlayerEquipment(index, Weapon), 0
                                GiveInvItem index, GetPlayerEquipment(index, Shield), 0
                                SetPlayerEquipment index, 0, Shield
                                SetPlayerEquipment index, itemNum, Weapon
                            Else
                                GiveInvItem index, GetPlayerEquipment(index, Shield), 0
                                SetPlayerEquipment index, 0, Shield
                                SetPlayerEquipment index, itemNum, Weapon
                            End If
                        Else
                            If GetPlayerEquipment(index, Weapon) > 0 Then
                                GiveInvItem index, GetPlayerEquipment(index, Weapon), 0
                                SetPlayerEquipment index, itemNum, Weapon
                        Else
                            SetPlayerEquipment index, itemNum, Weapon
                        End If
                    End If
                Else
                    If GetPlayerEquipment(index, Shield) > 0 Then
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                            GiveInvItem index, GetPlayerEquipment(index, Weapon), 0
                            SetPlayerEquipment index, itemNum, Weapon
                        Else
                            SetPlayerEquipment index, itemNum, Weapon
                        End If
                    Else
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                            GiveInvItem index, GetPlayerEquipment(index, Weapon), 0
                            SetPlayerEquipment index, itemNum, Weapon
                        Else
                            SetPlayerEquipment index, itemNum, Weapon
                        End If
                    End If
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            
            Case ITEM_TYPE_HELMET
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "คุณมี Status ไม่เพียงพอในการใช้ไอเทมนี้.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "ต้องการเลเวล " & Item(itemNum).LevelReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement

                ClassReq = Item(itemNum).ClassReq
    
                ' make sure the classreq > 0
                'If ClassReq > 0 Then ' 0 = no req
                '    If ClassReq <> GetPlayerClass(index) Then
                '        Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้ไอเทมนี้.", BrightRed)
                '        Exit Sub
                '    End If
                'End If
                
                Classtxt = "ต้องการอาชีพ : " ' Real
                    
                ' อาชีพ 1 ใช้ได้
                If Item(itemNum).ClassR1 > 0 Then
                    Classtxt = Classtxt & " มนุษย์,"
                    If GetPlayerClass(index) = 1 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 2 ใช้ได้
                If Item(itemNum).ClassR2 > 0 Then
                    Classtxt = Classtxt & " เอลฟ์,"
                    If GetPlayerClass(index) = 2 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 3 ใช้ได้
                If Item(itemNum).ClassR3 > 0 Then
                    Classtxt = Classtxt & " การ์เดี้ยน,"
                    If GetPlayerClass(index) = 3 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 4 ใช้ได้
                If Item(itemNum).ClassR4 > 0 Then
                    Classtxt = Classtxt & " เบอเซิร์ก,"
                    If GetPlayerClass(index) = 4 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 5 ใช้ได้
                If Item(itemNum).ClassR5 > 0 Then
                    Classtxt = Classtxt & " พาลาดิน,"
                    If GetPlayerClass(index) = 5 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 6 ใช้ได้
                If Item(itemNum).ClassR6 > 0 Then
                    Classtxt = Classtxt & " วิซาร์ด,"
                    If GetPlayerClass(index) = 6 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 7 ใช้ได้
                If Item(itemNum).ClassR7 > 0 Then
                    Classtxt = Classtxt & " ซามูไร,"
                    If GetPlayerClass(index) = 7 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 8 ใช้ได้
                If Item(itemNum).ClassR8 > 0 Then
                    Classtxt = Classtxt & " ฮันเตอร์,"
                    If GetPlayerClass(index) = 8 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 9 ใช้ได้
                If Item(itemNum).ClassR9 > 0 Then
                    Classtxt = Classtxt & " สไนเปอร์,"
                    If GetPlayerClass(index) = 9 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 10 ใช้ได้
                If Item(itemNum).ClassR10 > 0 Then
                    Classtxt = Classtxt & " แอสแซสซิน,"
                    If GetPlayerClass(index) = 10 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 11 ใช้ได้
                If Item(itemNum).ClassR11 > 0 Then
                    Classtxt = Classtxt & " ดาร์คลอร์ด,"
                    If GetPlayerClass(index) = 11 Then
                        Classchk = True
                    End If
                End If
                
                ' แจ้งข้อผิดพลาด
                If Item(itemNum).ClassR1 <= 0 And Item(itemNum).ClassR2 <= 0 And Item(itemNum).ClassR3 <= 0 And Item(itemNum).ClassR4 <= 0 And Item(itemNum).ClassR5 <= 0 And Item(itemNum).ClassR6 <= 0 And Item(itemNum).ClassR7 <= 0 And Item(itemNum).ClassR8 <= 0 And Item(itemNum).ClassR9 <= 0 And Item(itemNum).ClassR10 <= 0 And Item(itemNum).ClassR11 <= 0 Then
                    Call PlayerMsg(index, "ไม่มีอาชีพใดสามารถใช้ไอเทมนี้ได้.", BrightRed)
                    Exit Sub
                End If
                
                Classtxt = Classtxt & " ในการสวมใส่ไอเทมนี้."
                
                ' แจ้งข้อความว่าอาชีพใดใส่ได้บ้าง
                If Not Classchk = True Then
                    Call PlayerMsg(index, "อาชีพของคุณไม่สามารถสวมใส่ไอเทมนี้ได้.", BrightRed)
                    Call PlayerMsg(index, Classtxt, BrightRed)
                    Exit Sub
                End If
                
                ' แก้ไขบัคเพิ่มเติม Hp
                If GetPlayerMaxVital(index, HP) < Item(itemNum).HP Then
                    Call PlayerMsg(index, "ต้องการ MaxHp " & Item(itemNum).HP & " ในการสวมใส่ไอเทมนี้.", BrightRed)
                    Exit Sub
                End If
                
                ' แก้ไขบัคเพิ่มเติม Mp
                If GetPlayerMaxVital(index, MP) < Item(itemNum).MP Then
                    Call PlayerMsg(index, "ต้องการ MaxMp " & Item(itemNum).MP & " ในการสวมใส่ไอเทมนี้.", BrightRed)
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "ต้องการระดับ " & Item(itemNum).AccessReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(index, Helmet)
                End If

                SetPlayerEquipment index, itemNum, Helmet
                PlayerMsg index, "คุณได้สวมใส่ " & CheckGrammar(Item(itemNum).Name), BrightGreen
                TakeInvItem index, itemNum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            
            Case ITEM_TYPE_SHIELD
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "คุณมีค่า Status ไม่พอในการใช้ไอเทมนี้.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "ต้องการเลเวล " & Item(itemNum).LevelReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                
                ClassReq = Item(itemNum).ClassReq
    
                ' make sure the classreq > 0
                'If ClassReq > 0 Then ' 0 = no req
                '    If ClassReq <> GetPlayerClass(index) Then
                '        Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้ไอเทมนี้.", BrightRed)
                '        Exit Sub
                '    End If
                'End If
                
                Classtxt = "ต้องการอาชีพ : " ' Real
                    
                ' อาชีพ 1 ใช้ได้
                If Item(itemNum).ClassR1 > 0 Then
                    Classtxt = Classtxt & " มนุษย์,"
                    If GetPlayerClass(index) = 1 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 2 ใช้ได้
                If Item(itemNum).ClassR2 > 0 Then
                    Classtxt = Classtxt & " เอลฟ์,"
                    If GetPlayerClass(index) = 2 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 3 ใช้ได้
                If Item(itemNum).ClassR3 > 0 Then
                    Classtxt = Classtxt & " การ์เดี้ยน,"
                    If GetPlayerClass(index) = 3 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 4 ใช้ได้
                If Item(itemNum).ClassR4 > 0 Then
                    Classtxt = Classtxt & " เบอเซิร์ก,"
                    If GetPlayerClass(index) = 4 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 5 ใช้ได้
                If Item(itemNum).ClassR5 > 0 Then
                    Classtxt = Classtxt & " พาลาดิน,"
                    If GetPlayerClass(index) = 5 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 6 ใช้ได้
                If Item(itemNum).ClassR6 > 0 Then
                    Classtxt = Classtxt & " วิซาร์ด,"
                    If GetPlayerClass(index) = 6 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 7 ใช้ได้
                If Item(itemNum).ClassR7 > 0 Then
                    Classtxt = Classtxt & " ซามูไร,"
                    If GetPlayerClass(index) = 7 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 8 ใช้ได้
                If Item(itemNum).ClassR8 > 0 Then
                    Classtxt = Classtxt & " ฮันเตอร์,"
                    If GetPlayerClass(index) = 8 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 9 ใช้ได้
                If Item(itemNum).ClassR9 > 0 Then
                    Classtxt = Classtxt & " สไนเปอร์,"
                    If GetPlayerClass(index) = 9 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 10 ใช้ได้
                If Item(itemNum).ClassR10 > 0 Then
                    Classtxt = Classtxt & " แอสแซสซิน,"
                    If GetPlayerClass(index) = 10 Then
                        Classchk = True
                    End If
                End If
                ' อาชีพ 11 ใช้ได้
                If Item(itemNum).ClassR11 > 0 Then
                    Classtxt = Classtxt & " ดาร์คลอร์ด,"
                    If GetPlayerClass(index) = 11 Then
                        Classchk = True
                    End If
                End If
                
                ' แจ้งข้อผิดพลาด
                If Item(itemNum).ClassR1 <= 0 And Item(itemNum).ClassR2 <= 0 And Item(itemNum).ClassR3 <= 0 And Item(itemNum).ClassR4 <= 0 And Item(itemNum).ClassR5 <= 0 And Item(itemNum).ClassR6 <= 0 And Item(itemNum).ClassR7 <= 0 And Item(itemNum).ClassR8 <= 0 And Item(itemNum).ClassR9 <= 0 And Item(itemNum).ClassR10 <= 0 And Item(itemNum).ClassR11 <= 0 Then
                    Call PlayerMsg(index, "ไม่มีอาชีพใดสามารถใช้ไอเทมนี้ได้.", BrightRed)
                    Exit Sub
                End If
                
                Classtxt = Classtxt & " ในการสวมใส่ไอเทมนี้."
                
                ' แจ้งข้อความว่าอาชีพใดใส่ได้บ้าง
                If Not Classchk = True Then
                    Call PlayerMsg(index, "อาชีพของคุณไม่สามารถสวมใส่ไอเทมนี้ได้.", BrightRed)
                    Call PlayerMsg(index, Classtxt, BrightRed)
                    Exit Sub
                End If
                
                ' แก้ไขบัคเพิ่มเติม Hp
                If GetPlayerMaxVital(index, HP) < Item(itemNum).HP Then
                    Call PlayerMsg(index, "ต้องการ MaxHp " & Item(itemNum).HP & " ในการสวมใส่ไอเทมนี้.", BrightRed)
                    Exit Sub
                End If
                
                ' แก้ไขบัคเพิ่มเติม Mp
                If GetPlayerMaxVital(index, MP) < Item(itemNum).MP Then
                    Call PlayerMsg(index, "ต้องการ MaxMp " & Item(itemNum).MP & " ในการสวมใส่ไอเทมนี้.", BrightRed)
                    Exit Sub
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "ต้องการระดับ " & Item(itemNum).AccessReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(index, Shield)
                End If

                SetPlayerEquipment index, itemNum, Shield
                PlayerMsg index, "คุณได้สวมใส่ " & CheckGrammar(Item(itemNum).Name), BrightGreen
                TakeInvItem index, itemNum, 1

                If GetPlayerEquipment(index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).isTwohanded Then
                        GiveInvItem index, GetPlayerEquipment(index, Weapon), 0
                        SetPlayerEquipment index, 0, Weapon
                    Else
                        If tempItem > 0 Then
                            GiveInvItem index, tempItem, 0 ' give back the stored item
                            tempItem = 0
                        End If
                    End If
                End If
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendStats(index)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "คุณมีค่า Status ไม่พอในการใช้ไอเทมนี้.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "ต้องการเลเวล " & Item(itemNum).LevelReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                
                ClassReq = Item(itemNum).ClassReq
    
                ' make sure the classreq > 0
                If ClassReq > 0 Then ' 0 = no req
                    If ClassReq <> GetPlayerClass(index) Then
                        Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้ไอเทมนี้.", BrightRed)
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "ต้องการระดับ " & Item(itemNum).AccessReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(itemNum).AddHP > 0 Then
                    ' แก้บัค ไอเทม Overflow
                    If GetPlayerMaxVital(index, HP) > Player(index).Vital(Vitals.HP) + Item(itemNum).AddHP Then
                        Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Item(itemNum).AddHP
                    Else
                        Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, HP)
                    End If
                    
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, HP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add mp
                If Item(itemNum).AddMP > 0 Then
                    ' แก้บัค ไอเทม Overflow
                    If GetPlayerMaxVital(index, MP) > Player(index).Vital(Vitals.MP) + Item(itemNum).AddMP Then
                        Player(index).Vital(Vitals.MP) = Player(index).Vital(Vitals.MP) + Item(itemNum).AddMP
                    Else
                        Player(index).Vital(Vitals.MP) = GetPlayerMaxVital(index, MP)
                    End If
                    
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemNum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, MP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add exp
                If Item(itemNum).AddEXP > 0 Then
                    SetPlayerExp index, GetPlayerExp(index) + Item(itemNum).AddEXP
                    CheckPlayerLevelUp index
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendEXP index
                End If
                Call SendAnimation(GetPlayerMap(index), Item(itemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                Call TakeInvItem(index, Player(index).Inv(invNum).num, 0)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            
            Case ITEM_TYPE_KEY
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "คุณมีค่า Status ไม่พอในการใช้ไอเทมนี้.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "ต้องการเลเวล " & Item(itemNum).LevelReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                
                ClassReq = Item(itemNum).ClassReq
    
                ' make sure the classreq > 0
                If ClassReq > 0 Then ' 0 = no req
                    If ClassReq <> GetPlayerClass(index) Then
                        Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้ไอเทมนี้.", BrightRed)
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "ต้องการระดับ " & Item(itemNum).AccessReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If

                Select Case GetPlayerDir(index)
                    Case DIR_UP

                        If GetPlayerY(index) > 0 Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(index) < Map(GetPlayerMap(index)).MaxY Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT

                        If GetPlayerX(index) > 0 Then
                            x = GetPlayerX(index) - 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT

                        If GetPlayerX(index) < Map(GetPlayerMap(index)).MaxX Then
                            x = GetPlayerX(index) + 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If itemNum = Map(GetPlayerMap(index)).Tile(x, y).Data1 Then
                        TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                        SendMapKey index, x, y, 1
                        Call MapMsg(GetPlayerMap(index), "ประตูถูกปลดล็อคแล้ว.", White)
                        
                        Call SendAnimation(GetPlayerMap(index), Item(itemNum).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(index)).Tile(x, y).Data2 = 1 Then
                            Call TakeInvItem(index, itemNum, 0)
                            Call PlayerMsg(index, "คุณแจถูกทำลายแล้ว.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            Case ITEM_TYPE_SPELL
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "คุณมีค่า Status ไม่พอในการใช้ไอเทมนี้.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "ต้องการเลเวล " & Item(itemNum).LevelReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                
                ClassReq = Item(itemNum).ClassReq
    
                ' make sure the classreq > 0
                If ClassReq > 0 Then ' 0 = no req
                    If ClassReq <> GetPlayerClass(index) Then
                        Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้ไอเทมนี้.", BrightRed)
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "ต้องการระดับ " & Item(itemNum).AccessReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(itemNum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq

                        If i <= GetPlayerLevel(index) Then
                            i = FindOpenSpellSlot(index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(index, n) Then
                                    Call SetPlayerSpell(index, i, n)
                                    Call SendAnimation(GetPlayerMap(index), Item(itemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                    Call TakeInvItem(index, itemNum, 0)
                                    Call PlayerMsg(index, "คุณได้เรียนรู้สกิล " & Trim$(Spell(n).Name) & " สำเร็จแล้ว.", BrightGreen)
                                    Player(index).skillEXP(i) = 0
                                    Player(index).skillLV(i) = 0
                                    Call SendPlayerSpells(index)
                                    SendPlayerData index
                                Else
                                    Call PlayerMsg(index, "คุณได้เรียนรู้สกิลนี้ไปแล้ว.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(index, "คุณไม่สามารถเรียนรู้ทักษะเพิ่มเติมได้อีกแล้ว.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "ต้องการเลเวล " & i & " ในการเรียนสกิลนี้.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(index, "สกิลนี้สามารถเรียนได้เฉพาะอาชีพ " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & " เท่านั้น.", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
                
                 Case ITEM_TYPE_SUMMON
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "คุณมีค่า Status ไม่พอในการใช้ไอเทมนี้.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "ต้องการเลเวล " & Item(itemNum).LevelReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement

                ClassReq = Item(itemNum).ClassReq
    
                ' make sure the classreq > 0
                If ClassReq > 0 Then ' 0 = no req
                    If ClassReq <> GetPlayerClass(index) Then
                        Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้ไอเทมนี้.", BrightRed)
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "ต้องการระดับ " & Item(itemNum).AccessReq & " ในการใช้ไอเทมนี้.", BrightRed
                    Exit Sub
                End If
                
                ' ปิดใช้ระบบสัตว์เลี้ยงชั่วคราว
                If IsPlaying(index) Then
                    PlayerMsg index, "ระบบสัตว์เลี้ยงถูกปิดการใช้งานชั่วคราว.", BrightRed
                    Exit Sub
                End If
                
                Call SpawnPet(index, GetPlayerMap(index), itemNum)
                PetFollowOwner index
                
                Case ITEM_TYPE_SCRIPT
                
                ' price is script
                Call CustomScript(index, Item(GetPlayerInvItemNum(index, invNum)).price)
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
                Call TakeInvItem(index, itemNum, 0)
                
                Case ITEM_TYPE_RECIPE
                
                ' Get the recipe information
                Item1 = Item(GetPlayerInvItemNum(index, invNum)).Data1
                Item2 = Item(GetPlayerInvItemNum(index, invNum)).Data2
                Result = Item(GetPlayerInvItemNum(index, invNum)).Data3
    
                ' Perform Recipe checks
                If Item1 <= 0 Then
                    Call PlayerMsg(index, "นี้เป็นสูตรที่ไม่สมบูรณ์...", White)
                    Exit Sub
                End If
                
                If Item2 <= 0 Then
                    Call PlayerMsg(index, "นี้เป็นสูตรที่ไม่สมบูรณ์...", White)
                    Exit Sub
                End If
                
                If Result <= 0 Then
                    Call PlayerMsg(index, "นี้เป็นสูตรที่ไม่สมบูรณ์...", White)
                    Exit Sub
                End If
                
                If GetPlayerEquipment(index, Weapon) <= 0 Then
                    Call PlayerMsg(index, "คุณไม่ได้สวมใส่เครื่องมือในการผลิตไอเทมชิ้นนี้ !", White)
                    Exit Sub
                End If
            
            If Craft(0) <> 1 Then
                Call PlayerMsg(index, "ต้องยืนอยู่ในเขตของเตาหลอมเพื่อสร้างไอเทม !", BrightRed)
                Exit Sub
                End If
                
                If Item(GetPlayerEquipment(index, Weapon)).Tool = Item(GetPlayerInvItemNum(index, invNum)).ToolReq Then
                    ' Give the resulting item
                        If HasItem(index, Item1) Then
                            If HasItem(index, Item2) Then
                                Call TakeInvItem(index, Item1, 1)
                                Call TakeInvItem(index, Item2, 1)
                                Call TakeInvItem(index, itemNum, 1)
                                randD = rand(1, 100)
                                ' มีโอกาศล้มเหลวเมื่อผลิต 50%
                                If randD >= 50 Then
                                    Call GiveInvItem(index, Result, 1)
                                    Call PlayerMsg(index, "คุณได้สร้างไอเทม " & Trim(Item(Result).Name) & " เสร็จสมบูรณ์แล้ว.", White)
                                Else
                                    Call PlayerMsg(index, "การสร้างไอเทม " & Trim(Item(Result).Name) & " ล้มเหลว.", BrightRed)
                                End If
                            Else
                                Call PlayerMsg(index, "คุณมีส่วนผสมยังไม่ครบ.", White)
                                Exit Sub
                            End If
                        Else
                            Call PlayerMsg(index, "คุณมีส่วนผสมยังไม่ครบ.", White)
                            Exit Sub
                        End If
                    Else
                        Call PlayerMsg(index, "ไม่สามารถใช้เครื่องมือนี้ในการผลิตได้.", White)
                        Exit Sub
                    End If
        End Select
    End If
End Sub

Function CheckCash(ByVal index As Long, ByVal CashItemNum As Long, ByVal CashAmount As Long) As Boolean

    Dim i As Long

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) = 1 Then
            If Item(1).Type = ITEM_TYPE_CURRENCY Then
                If CashAmount < 1 = GetPlayerInvItemValue(index, i) Then
                    CheckCash = True
                Else
                    CheckCash = True
                End If
            End If
        End If
    Next
End Function

'Checkpoint
Public Sub SetCheckpoint(ByVal index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    PlayerMsg index, "คุณได้ทำการบันทึกจุดเกิดแล้ว !", BrightGreen
    Call PutVar(App.Path & "\Data\accounts\checkpoints.ini", "*SERVEROPTIONS*", "" & GetPlayerName(index), "" & 1)
    Call PutVar(App.Path & "\Data\accounts\checkpoints.ini", "" & GetPlayerName(index), "MAPNUM", "" & mapnum)
    Call PutVar(App.Path & "\Data\accounts\checkpoints.ini", "" & GetPlayerName(index), "X", "" & x)
    Call PutVar(App.Path & "\Data\accounts\checkpoints.ini", "" & GetPlayerName(index), "Y", "" & y)
End Sub

Public Sub WarpToCheckpoint(ByVal index As Long)
    Dim mapnum As Integer
    Dim x As Integer
    Dim y As Integer
    
    mapnum = GetVar(App.Path & "\Data\accounts\checkpoints.ini", "" & GetPlayerName(index), "MAPNUM")
    x = GetVar(App.Path & "\Data\accounts\checkpoints.ini", "" & GetPlayerName(index), "X")
    y = GetVar(App.Path & "\Data\accounts\checkpoints.ini", "" & GetPlayerName(index), "Y")
    Call PlayerWarp(index, mapnum, x, y)
End Sub

Public Function ValCheckPoint(ByVal index As Long) As Boolean
    If GetVar(App.Path & "\Data\accounts\checkpoints.ini", "*SERVEROPTIONS*", "" & GetPlayerName(index)) = 1 Then
        ValCheckPoint = True
    Else
        ValCheckPoint = False
    End If
End Function

Sub CheckDoor(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Door_num As Long
    Dim i As Long
    Dim n As Long
    Dim key As Long
    Dim tmpIndex As Long
    
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_DOOR Then
        Door_num = Map(GetPlayerMap(index)).Tile(x, y).Data1



        If Door_num > 0 Then
            If Doors(Door_num).DoorType = 0 Then
                If Player(index).PlayerDoors(Door_num).state = 0 Then
                    If Doors(Door_num).UnlockType = 0 Then
                        For i = 1 To MAX_INV
                            key = GetPlayerInvItemNum(index, i)
                            If Doors(Door_num).key = key Then
                                TakeInvItem index, key, 1
                                If TempPlayer(index).inParty > 0 Then
                                    For n = 1 To MAX_PARTY_MEMBERS
                                        tmpIndex = Party(TempPlayer(index).inParty).Member(n)
                                        If tmpIndex > 0 Then
                                            Player(tmpIndex).PlayerDoors(Door_num).state = 1
                                            SendPlayerData (tmpIndex)
                                            If index <> tmpIndex Then
                                                PlayerMsg tmpIndex, "ปาร์ตี้ของคุณได้ปลดล็อคประตู.", BrightBlue
                                            Else
                                                PlayerMsg tmpIndex, "คุณได้ใช้กุญแจปลดล็อคประตู.", BrightBlue
                                            End If
                                        End If
                                    Next
                                    
                                Else
                                    Player(index).PlayerDoors(Door_num).state = 1
                                    PlayerMsg index, "ใช้กุญแจ 1 ดอกเพื่อปลดล็อคประตู.", BrightBlue
                                    SendPlayerData (index)
                                End If
                                Exit Sub
                            End If
                        Next
                        PlayerMsg index, "คุณไม่มีกุญแจเพื่อปลดล็อคประตู.", BrightBlue
                    ElseIf Doors(Door_num).UnlockType = 1 Then
                        If Doors(Door_num).state = 0 Then
                            PlayerMsg index, "คุณไม่มีสวิตช์เพื่อปลดล็อคประตู.", BrightBlue
                        End If
                    ElseIf Doors(Door_num).UnlockType = 2 Then
                        PlayerMsg index, "ประตูนี้ไม่ได้ล็อค", BrightBlue
                    End If
                    
                Else
                    PlayerMsg index, "ประตูไม่ได้ล็อคอยู่.", BrightBlue
                End If
            ElseIf Doors(Door_num).DoorType = 1 Then
                If Player(index).PlayerDoors(Door_num).state = 0 Then
                                If TempPlayer(index).inParty > 0 Then
                                    For n = 1 To MAX_PARTY_MEMBERS
                                        tmpIndex = Party(TempPlayer(index).inParty).Member(n)
                                        If tmpIndex > 0 Then
                                            Player(tmpIndex).PlayerDoors(Door_num).state = 1
                                            Player(tmpIndex).PlayerDoors(Doors(Door_num).Switch).state = 1
                                            SendPlayerData (tmpIndex)
                                            If index <> tmpIndex Then
                                                PlayerMsg tmpIndex, "ปาร์ตี้ของคุณได้เปิดสวิชต์และปลดล็อคประตูแล้ว.", BrightBlue
                                            Else
                                                PlayerMsg tmpIndex, "คุณได้เปิดสวิชต์และปลดล็อคประตูแล้ว.", BrightBlue
                                            End If
                                        End If
                                    Next
                                Else
                                    Player(index).PlayerDoors(Door_num).state = 1
                                    Player(index).PlayerDoors(Doors(Door_num).Switch).state = 1
                                    PlayerMsg index, "คุณได้เปิดสวิชต์และปลดล็อคประตูแล้ว.", BrightBlue
                                    SendPlayerData (index)
                                End If
                    
                Else
                                If TempPlayer(index).inParty > 0 Then
                                    For n = 1 To MAX_PARTY_MEMBERS
                                        tmpIndex = Party(TempPlayer(index).inParty).Member(n)
                                        If tmpIndex > 0 Then
                                            Player(tmpIndex).PlayerDoors(Door_num).state = 0
                                            Player(tmpIndex).PlayerDoors(Doors(Door_num).Switch).state = 0
                                            SendPlayerData (tmpIndex)
                                            If index <> tmpIndex Then
                                                PlayerMsg tmpIndex, "ปาร์ตี้ของคุณได้ปิดสวิชต์และล็อคประตูแล้ว.", BrightBlue
                                            Else
                                                PlayerMsg tmpIndex, "คุณได้ปิดสวิชต์และล็อคประตูแล้ว.", BrightBlue
                                            End If
                                        End If
                                    Next
                                Else
                                    Player(index).PlayerDoors(Door_num).state = 0
                                    Player(index).PlayerDoors(Doors(Door_num).Switch).state = 0
                                    PlayerMsg index, "คุณได้ปิดสวิชต์และล็อคประตูแล้ว.", BrightBlue
                                    SendPlayerData (index)
                                End If
                End If
            End If
        End If
    End If
End Sub

Function FindItem(ByVal index As Long, ByVal itemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemNum Then
            FindItem = i
            Exit Function
        End If

    Next

End Function

Sub SetPlayerdagger(ByVal index As Long, ByVal dagger As Long)
    Player(index).WieldDagger = Player(index).WieldDagger - Player(index).WieldDagger
End Sub

Function IsPlayerBusy(ByVal index As Long, ByVal OtherPlayer As Long) As Boolean
    ' Make sure they're not busy doing something else
    If IsPlaying(OtherPlayer) Then
        If TempPlayer(OtherPlayer).InBank Or TempPlayer(OtherPlayer).InShop > 0 Or TempPlayer(OtherPlayer).InTrade > 0 Or (TempPlayer(OtherPlayer).partyInvite > 0 And TempPlayer(OtherPlayer).partyInvite <> index) Or (TempPlayer(OtherPlayer).TradeRequest > 0 And TempPlayer(OtherPlayer).TradeRequest <> index) Or (TempPlayer(OtherPlayer).tmpGuildInviteId > 0 And TempPlayer(OtherPlayer).tmpGuildInviteId <> index) Then
            IsPlayerBusy = True
            PlayerMsg index, GetPlayerName(OtherPlayer) & " กำลังอยู่ในสถานะใด ๆ กับผู้เล่นอื่นอยู่!", BrightRed
            Exit Function
        End If
    End If
End Function

