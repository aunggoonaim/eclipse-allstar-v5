Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal mapNum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If itemnum < 1 Or itemnum > MAX_ITEMS Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapNum)
    Call SpawnItemSlot(i, itemnum, ItemVal, mapNum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemnum < 0 Or itemnum > MAX_ITEMS Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If itemnum >= 0 And itemnum <= MAX_ITEMS Then
            MapItem(mapNum, i).playerName = playerName
            MapItem(mapNum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(mapNum, i).canDespawn = canDespawn
            MapItem(mapNum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(mapNum, i).Num = itemnum
            MapItem(mapNum, i).Value = ItemVal
            MapItem(mapNum, i).x = x
            MapItem(mapNum, i).y = y
            ' send to map
            SendSpawnItemToMap mapNum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal mapNum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(mapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(mapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(mapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapNum).Tile(x, y).Data1, 1, mapNum, x, y)
                Else
                    Call SpawnItem(Map(mapNum).Tile(x, y).Data1, Map(mapNum).Tile(x, y).Data2, mapNum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal mapNpcNum As Long, ByVal mapNum As Long, Optional ByVal SetX As Long, Optional ByVal SetY As Long)
    Dim Buffer As clsBuffer
    Dim npcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or mapNum <= 0 Or mapNum > MAX_MAPS Then Exit Sub
    npcNum = Map(mapNum).NPC(mapNpcNum)

    If npcNum > 0 Then
    
        MapNpc(mapNum).NPC(mapNpcNum).Num = npcNum
        MapNpc(mapNum).NPC(mapNpcNum).target = 0
        MapNpc(mapNum).NPC(mapNpcNum).targetType = 0 ' clear
        
        MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
        MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(npcNum, Vitals.MP)
        
        MapNpc(mapNum).NPC(mapNpcNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(mapNum).MaxX
            For y = 0 To Map(mapNum).MaxY
                If Map(mapNum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(mapNum).Tile(x, y).Data1 = mapNpcNum Then
                        MapNpc(mapNum).NPC(mapNpcNum).x = x
                        MapNpc(mapNum).NPC(mapNpcNum).y = y
                        MapNpc(mapNum).NPC(mapNpcNum).Dir = Map(mapNum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                
                If SetX = 0 And SetY = 0 Then
                    x = Random(0, Map(mapNum).MaxX)
                    y = Random(0, Map(mapNum).MaxY)
                Else
                    x = SetX
                    y = SetY
                End If
    
                If x > Map(mapNum).MaxX Then x = Map(mapNum).MaxX
                If y > Map(mapNum).MaxY Then y = Map(mapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapNum, x, y) Then
                    MapNpc(mapNum).NPC(mapNpcNum).x = x
                    MapNpc(mapNum).NPC(mapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(mapNum).MaxX
                For y = 0 To Map(mapNum).MaxY

                    If NpcTileIsOpen(mapNum, x, y) Then
                        MapNpc(mapNum).NPC(mapNpcNum).x = x
                        MapNpc(mapNum).NPC(mapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).Num
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).Dir
            Buffer.WriteByte MapNpc(mapNum).NPC(mapNpcNum).IsPet
            Buffer.WriteString MapNpc(mapNum).NPC(mapNpcNum).PetData.Name
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).PetData.Owner
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        End If
        
        SendMapNpcVitals mapNum, mapNpcNum
    End If

End Sub

Public Function NpcTileIsOpen(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(mapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = mapNum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(mapNum).NPC(LoopI).Num > 0 Then
            If MapNpc(mapNum).NPC(LoopI).x = x Then
                If MapNpc(mapNum).NPC(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(mapNum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(mapNum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(mapNum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, mapNum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Function CanNpcMove(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(mapNum).NPC(mapNpcNum).x
    y = MapNpc(mapNum).NPC(mapNpcNum).y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapNum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapNum).MaxY Then
                n = Map(mapNum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapNum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x - 1) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapNum).MaxX Then
                n = Map(mapNum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapNum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).NPC(i).Num > 0) And (MapNpc(mapNum).NPC(i).x = MapNpc(mapNum).NPC(mapNpcNum).x + 1) And (MapNpc(mapNum).NPC(i).y = MapNpc(mapNum).NPC(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(mapNum).NPC(mapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNpc(mapNum).NPC(mapNpcNum).y = MapNpc(mapNum).NPC(mapNpcNum).y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapNum).NPC(mapNpcNum).y = MapNpc(mapNum).NPC(mapNpcNum).y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapNum).NPC(mapNpcNum).x = MapNpc(mapNum).NPC(mapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapNum).NPC(mapNpcNum).x = MapNpc(mapNum).NPC(mapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub

Sub NpcDir(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapNum).NPC(mapNpcNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapNpcNum
    Buffer.WriteLong Dir
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapNum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = mapNum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal mapNum As Long)
    Dim y As Long
    Dim x As Long
    TempTile(mapNum).DoorTimer = 0
    ReDim TempTile(mapNum).DoorOpen(0 To Map(mapNum).MaxX, 0 To Map(mapNum).MaxY)

    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY
            TempTile(mapNum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapNum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY

            If Map(mapNum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapNum).ResourceData(0 To Resource_Count)
                ResourceCache(mapNum).ResourceData(Resource_Count).x = x
                ResourceCache(mapNum).ResourceData(Resource_Count).y = y
                ResourceCache(mapNum).ResourceData(Resource_Count).cur_health = Resource(Map(mapNum).Tile(x, y).Data1).health
            End If

        Next
    Next

    ResourceCache(mapNum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    SendInventory index
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(index, oldSlot)
    NewNum = GetPlayerSpell(index, newSlot)
    SetPlayerSpell index, oldSlot, NewNum
    SetPlayerSpell index, newSlot, OldNum
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0
        PlayerMsg index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SendWornEquipment index
        SendMapEquipment index
        SendStats index
        ' send vitals
        Call SendVital(index, Vitals.HP)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
    Else
        PlayerMsg index, "Your inventory is full.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal index As Long)
Dim partyNum As Long, i As Long

    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
            ' check if leader
            If Party(partyNum).Leader = index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(i) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                index = Party(partyNum).Member(i)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(index).inParty > 0 Then
        partyNum = TempPlayer(index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = index
                    ' let them know
                    PlayerMsg index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = index
        ' let them know
        PlayerMsg index, "Invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if already in a party
    If TempPlayer(index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) = 0 Then
                'add to the party
                Party(partyNum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                SendPartyVitals partyNum, targetPlayer
                ' let everyone know they've joined
                PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = index
        Party(partyNum).Member(1) = index
        Party(partyNum).Member(2) = targetPlayer
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, index
        SendPartyVitals partyNum, targetPlayer
        ' let them know it's created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(index).inParty = partyNum
        TempPlayer(targetPlayer).inParty = partyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal targetPlayer As Long)
    PlayerMsg index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, x As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(x) = Party(partyNum).Member(x + 1)
                    Party(partyNum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partyNum
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal exp As Long, ByVal index As Long)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long

    ' check if it's worth sharing
    If Not exp >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP index, exp
        Exit Sub
    End If
    
    ' find out the equal share
    expShare = exp \ Party(partyNum).MemberCount
    leftOver = exp Mod Party(partyNum).MemberCount
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' give them their share
                GivePlayerEXP tmpIndex, expShare
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(RAND(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver
End Sub

Public Sub GivePlayerEXP(ByVal index As Long, ByVal exp As Long)
    ' give the exp
    Call SetPlayerExp(index, GetPlayerExp(index) + exp)
    SendEXP index
    SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    ' check if we've leveled
    CheckPlayerLevelUp index
End Sub

' projectiles
Public Sub HandleProjecTile(ByVal index As Long, ByVal PlayerProjectile As Long)
Dim x As Long, y As Long, i As Long

    ' check for subscript out of range
    If index < 1 Or index > MAX_PLAYERS Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
        
    ' check to see if it's time to move the Projectile
    If GetTickCount > TempPlayer(index).ProjecTile(PlayerProjectile).TravelTime Then
        With TempPlayer(index).ProjecTile(PlayerProjectile)
            ' set next travel time and the current position and then set the actual direction based on RMXP arrow tiles.
            Select Case .Direction
                ' down
                Case DIR_DOWN
                    .y = .y + 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(index) + .Range) + 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' up
                Case DIR_UP
                    .y = .y - 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(index) - .Range) - 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' right
                Case DIR_RIGHT
                    .x = .x + 1
                    ' check if they reached max range
                    If .x = (GetPlayerX(index) + .Range) + 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' left
                Case DIR_LEFT
                    .x = .x - 1
                    ' check if they reached maxrange
                    If .x = (GetPlayerX(index) - .Range) - 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
            End Select
            .TravelTime = GetTickCount + .Speed
        End With
    End If
    
    x = TempPlayer(index).ProjecTile(PlayerProjectile).x
    y = TempPlayer(index).ProjecTile(PlayerProjectile).y
    
    ' check if left map
    If x > Map(Player(index).Map).MaxX Or y > Map(Player(index).Map).MaxY Or x < 0 Or y < 0 Then
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if hit player
    For i = 1 To Player_HighIndex
        ' make sure they're actually playing
        If IsPlaying(i) Then
            ' check coordinates
            If x = Player(i).x And y = GetPlayerY(i) Then
                ' make sure it's not the attacker
                If Not x = Player(index).x Or Not y = GetPlayerY(index) Then
                    ' check if player can attack
                    If CanPlayerAttackPlayer(index, i, False, True) = True Then
                        ' attack the player and kill the project tile
                        PlayerAttackPlayer index, i, TempPlayer(index).ProjecTile(PlayerProjectile).Damage
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    Else
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    ' check for npc hit
    For i = 1 To MAX_MAP_NPCS
        If x = MapNpc(GetPlayerMap(index)).NPC(i).x And y = MapNpc(GetPlayerMap(index)).NPC(i).y Then
            ' they're hit, remove it and deal that damage ;)
            PlayerAttackNpc index, i, TempPlayer(index).ProjecTile(PlayerProjectile).Damage * Round((Player(index).Level * 1.14)) + Round((GetPlayerStat(index, Stats.Agility) * 0.6))
            ClearProjectile index, PlayerProjectile
            Exit Sub
        End If
    Next
    
    ' hit a block
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        ' hit a block, clear it.
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
End Sub


'makes the pet follow its owner
Sub PetFollowOwner(ByVal index As Long)
    If TempPlayer(index).TempPetSlot < 1 Then Exit Sub
    
    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).targetType = 1
    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).target = index
End Sub

'makes the pet wander around the map
Sub PetWander(ByVal index As Long)
    If TempPlayer(index).TempPetSlot < 1 Then Exit Sub

    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).targetType = TARGET_TYPE_NONE
    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).target = 0
End Sub

'Clear the npc from the map
Sub PetDisband(ByVal index As Long, ByVal mapNum As Long)
    Dim i As Long
    Dim j As Long

    If TempPlayer(index).TempPetSlot < 1 Then Exit Sub

    'Cache the Pets for players logging on [Remove Number from array]
    'THIS IS KINDA SLOW (EVEN WITHOUT TESTING, LOL), MAY HAVE TO CONVERT TO LINKED LIST FOR SPEED
    For i = 1 To PetMapCache(mapNum).UpperBound
        If PetMapCache(mapNum).Pet(i) = TempPlayer(index).TempPetSlot Then
            If PetMapCache(mapNum).UpperBound > 1 Then
                For j = PetMapCache(mapNum).UpperBound To i Step -1
                    PetMapCache(mapNum).Pet(j - 1) = PetMapCache(mapNum).Pet(j)
                Next
            Else
                PetMapCache(mapNum).Pet(1) = 0
            End If
            
            PetMapCache(mapNum).UpperBound = PetMapCache(mapNum).UpperBound - 1
            Exit For
        End If
    Next
    
    Call ClearSingleMapNpc(TempPlayer(index).TempPetSlot, mapNum)
    Map(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot) = 0
    TempPlayer(index).TempPetSlot = 0

    're-warp the players on the map
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then
                Call PlayerWarp(i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i))
                SendPlayerData index
            End If
        End If
    Next
End Sub


Sub SpawnPet(ByVal index As Long, ByVal mapNum As Long, npcNum As Long)
    Dim PlayerMap As Long
    Dim i As Integer
    Dim PetSlot As Byte
    
    'Prevent multiple pets for the same owner
    If TempPlayer(index).TempPetSlot > 0 Then Exit Sub
    
    PlayerMap = GetPlayerMap(index)
    PetSlot = 0
    
    For i = 1 To MAX_MAP_NPCS
        'If Map(PlayerMap).Npc(i) = 0 Then
        If MapNpc(PlayerMap).NPC(i).SpawnWait = 0 And MapNpc(PlayerMap).NPC(i).Num = 0 Then
            PetSlot = i
            Exit For
        End If
    Next
    
    If PetSlot = 0 Then
        Call PlayerMsg(index, "The map is too crowded for you to call on your pet!", Red)
        Exit Sub
    End If

    'create the pet for the map
    Map(PlayerMap).NPC(PetSlot) = npcNum
    MapNpc(PlayerMap).NPC(PetSlot).Num = npcNum
    'set its Pet Data
    MapNpc(PlayerMap).NPC(PetSlot).IsPet = YES
    MapNpc(PlayerMap).NPC(PetSlot).PetData.Name = GetPlayerName(index) & "'s " & NPC(npcNum).Name
    MapNpc(PlayerMap).NPC(PetSlot).PetData.Owner = index
    
    'If Pet doesn't exist with player, link it to the player
    If Player(index).Pet.SpriteNum <> npcNum Then
        Player(index).Pet.SpriteNum = npcNum
        Player(index).Pet.Name = GetPlayerName(index) & "'s " & NPC(npcNum).Name
    End If
    
    TempPlayer(index).TempPetSlot = PetSlot
       
    'cache the map for sending
    Call MapCache_Create(PlayerMap)

    'Cache the Pets for players logging on [Add new Number to array]
    PetMapCache(PlayerMap).UpperBound = PetMapCache(PlayerMap).UpperBound + 1
    PetMapCache(PlayerMap).Pet(PetMapCache(PlayerMap).UpperBound) = PetSlot
    
    If PetMapCache(Player(index).Map).UpperBound > 0 Then
        For i = 1 To PetMapCache(Player(index).Map).UpperBound
            Call NPCCache_Create(index, Player(index).Map, PetMapCache(Player(index).Map).Pet(i))
        Next
    End If

    Select Case GetPlayerDir(index)
        Case DIR_UP
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index), GetPlayerY(index) - 1)
        Case DIR_DOWN
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index), GetPlayerY(index) + 1)
        Case DIR_LEFT
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index) + 1, GetPlayerY(index))
        Case DIR_RIGHT
            Call SpawnNpc(PetSlot, PlayerMap, GetPlayerX(index), GetPlayerY(index) - 1)
    End Select
    
    're-warp the players on the map
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then
                Call PlayerWarp(i, PlayerMap, GetPlayerX(i), GetPlayerY(i))
            End If
        End If
    Next
    
End Sub
