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

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapnum, i).num = 0 Then
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

Sub SpawnItem(ByVal itemNum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If itemNum < 1 Or itemNum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapnum)
    Call SpawnItemSlot(i, itemNum, ItemVal, mapnum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemNum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemNum < 0 Or itemNum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If itemNum >= 0 And itemNum <= MAX_ITEMS Then
            MapItem(mapnum, i).playerName = playerName
            MapItem(mapnum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(mapnum, i).canDespawn = canDespawn
            MapItem(mapnum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(mapnum, i).num = itemNum
            MapItem(mapnum, i).Value = ItemVal
            MapItem(mapnum, i).x = x
            MapItem(mapnum, i).y = y
            ' send to map
            SendSpawnItemToMap mapnum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(mapnum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(mapnum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(mapnum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, 1, mapnum, x, y)
                Else
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, Map(mapnum).Tile(x, y).Data2, mapnum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal mapNpcNum As Long, ByVal mapnum As Long, Optional ByVal SetX As Long, Optional ByVal SetY As Long)
    Dim Buffer As clsBuffer
    Dim npcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Sub
    npcNum = Map(mapnum).NPC(mapNpcNum)

    If npcNum > 0 Then
    
        MapNpc(mapnum).NPC(mapNpcNum).num = npcNum
        MapNpc(mapnum).NPC(mapNpcNum).Target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0 ' clear
        MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(npcNum, Vitals.MP)
        
        MapNpc(mapnum).NPC(mapNpcNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(mapnum).MaxX
            For y = 0 To Map(mapnum).MaxY
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(mapnum).Tile(x, y).Data1 = mapNpcNum Then
                        MapNpc(mapnum).NPC(mapNpcNum).x = x
                        MapNpc(mapnum).NPC(mapNpcNum).y = y
                        MapNpc(mapnum).NPC(mapNpcNum).Dir = Map(mapnum).Tile(x, y).Data2
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
                    x = Random(0, Map(mapnum).MaxX)
                    y = Random(0, Map(mapnum).MaxY)
                Else
                    x = SetX
                    y = SetY
                End If
    
                If x > Map(mapnum).MaxX Then x = Map(mapnum).MaxX
                If y > Map(mapnum).MaxY Then y = Map(mapnum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapnum, x, y) Then
                    MapNpc(mapnum).NPC(mapNpcNum).x = x
                    MapNpc(mapnum).NPC(mapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(mapnum).MaxX
                For y = 0 To Map(mapnum).MaxY

                    If NpcTileIsOpen(mapnum, x, y) Then
                        MapNpc(mapnum).NPC(mapNpcNum).x = x
                        MapNpc(mapnum).NPC(mapNpcNum).y = y
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
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).num
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            Buffer.WriteByte MapNpc(mapnum).NPC(mapNpcNum).IsPet
            Buffer.WriteString MapNpc(mapnum).NPC(mapNpcNum).PetData.Name
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        End If
        
        SendMapNpcVitals mapnum, mapNpcNum
    End If

End Sub

Public Sub SpawnMapEventsFor(index As Long, mapnum As Long)
Dim i As Long, x As Long, y As Long, z As Long, spawncurrentevent As Boolean, p As Long
Dim Buffer As clsBuffer
    
    TempPlayer(index).EventMap.CurrentEvents = 0
    ReDim TempPlayer(index).EventMap.EventPages(0)
    
    If Map(mapnum).EventCount <= 0 Then Exit Sub
    For i = 1 To Map(mapnum).EventCount
        If Map(mapnum).Events(i).PageCount > 0 Then
            For z = Map(mapnum).Events(i).PageCount To 1 Step -1
                With Map(mapnum).Events(i).Pages(z)
                    spawncurrentevent = True
                    
                    If .chkVariable = 1 Then
                        If Player(index).Variables(.VariableIndex) < .VariableCondition Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSwitch = 1 Then
                        If Player(index).Switches(.SwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkHasItem = 1 Then
                        If HasItem(index, .HasItemIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSelfSwitch = 1 Then
                        If Map(mapnum).Events(i).SelfSwitches(.SelfSwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If spawncurrentevent = True Or (spawncurrentevent = False And z = 1) Then
                        'spawn the event... send data to player
                        TempPlayer(index).EventMap.CurrentEvents = TempPlayer(index).EventMap.CurrentEvents + 1
                        ReDim Preserve TempPlayer(index).EventMap.EventPages(TempPlayer(index).EventMap.CurrentEvents)
                        With TempPlayer(index).EventMap.EventPages(TempPlayer(index).EventMap.CurrentEvents)
                            If Map(mapnum).Events(i).Pages(z).GraphicType = 1 Then
                                Select Case Map(mapnum).Events(i).Pages(z).GraphicY
                                    Case 0
                                        .Dir = DIR_DOWN
                                    Case 1
                                        .Dir = DIR_LEFT
                                    Case 2
                                        .Dir = DIR_RIGHT
                                    Case 3
                                        .Dir = DIR_UP
                                End Select
                            Else
                                .Dir = 0
                            End If
                            .GraphicNum = Map(mapnum).Events(i).Pages(z).Graphic
                            .GraphicType = Map(mapnum).Events(i).Pages(z).GraphicType
                            .GraphicX = Map(mapnum).Events(i).Pages(z).GraphicX
                            .GraphicY = Map(mapnum).Events(i).Pages(z).GraphicY
                            .GraphicX2 = Map(mapnum).Events(i).Pages(z).GraphicX2
                            .GraphicY2 = Map(mapnum).Events(i).Pages(z).GraphicY2
                            Select Case Map(mapnum).Events(i).Pages(z).MoveSpeed
                                Case 0
                                    .movementspeed = 2
                                Case 1
                                    .movementspeed = 3
                                Case 2
                                    .movementspeed = 4
                                Case 3
                                    .movementspeed = 6
                                Case 4
                                    .movementspeed = 12
                                Case 5
                                    .movementspeed = 24
                            End Select
                            If Map(mapnum).Events(i).Global Then
                                .x = TempEventMap(mapnum).Events(i).x
                                .y = TempEventMap(mapnum).Events(i).y
                                .Dir = TempEventMap(mapnum).Events(i).Dir
                                .MoveRouteStep = TempEventMap(mapnum).Events(i).MoveRouteStep
                            Else
                                .x = Map(mapnum).Events(i).x
                                .y = Map(mapnum).Events(i).y
                                .MoveRouteStep = 0
                            End If
                            .Position = Map(mapnum).Events(i).Pages(z).Position
                            .eventID = i
                            .pageID = z
                            If spawncurrentevent = True Then
                                .Visible = 1
                            Else
                                .Visible = 0
                            End If
                            
                            .MoveType = Map(mapnum).Events(i).Pages(z).MoveType
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(mapnum).Events(i).Pages(z).MoveRouteCount
                                ReDim .MoveRoute(0 To Map(mapnum).Events(i).Pages(z).MoveRouteCount)
                                If Map(mapnum).Events(i).Pages(z).MoveRouteCount > 0 Then
                                    For p = 0 To Map(mapnum).Events(i).Pages(z).MoveRouteCount
                                        .MoveRoute(p) = Map(mapnum).Events(i).Pages(z).MoveRoute(p)
                                    Next
                                End If
                            End If
                            
                            .RepeatMoveRoute = Map(mapnum).Events(i).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(mapnum).Events(i).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(mapnum).Events(i).Pages(z).MoveFreq
                            .MoveSpeed = Map(mapnum).Events(i).Pages(z).MoveSpeed
                            
                            .WalkingAnim = Map(mapnum).Events(i).Pages(z).WalkAnim
                            .WalkThrough = Map(mapnum).Events(i).Pages(z).WalkThrough
                            .FixedDir = Map(mapnum).Events(i).Pages(z).DirFix
                            
                        End With
                        GoTo nextevent
                    End If
                End With
            Next
        End If
nextevent:
    Next
    
    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
        For i = 1 To TempPlayer(index).EventMap.CurrentEvents
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnEvent
            Buffer.WriteLong i
            With TempPlayer(index).EventMap.EventPages(i)
                Buffer.WriteString Map(GetPlayerMap(index)).Events(i).Name
                Buffer.WriteLong .Dir
                Buffer.WriteLong .GraphicNum
                Buffer.WriteLong .GraphicType
                Buffer.WriteLong .GraphicX
                Buffer.WriteLong .GraphicX2
                Buffer.WriteLong .GraphicY
                Buffer.WriteLong .GraphicY2
                Buffer.WriteLong .movementspeed
                Buffer.WriteLong .x
                Buffer.WriteLong .y
                Buffer.WriteLong .Position
                Buffer.WriteLong .Visible
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).WalkAnim
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).DirFix
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).WalkThrough
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).ShowName
            End With
            SendDataTo index, Buffer.ToArray
            Set Buffer = Nothing
        Next
    End If
End Sub

Public Function NpcTileIsOpen(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(mapnum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = mapnum Then
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

        If MapNpc(mapnum).NPC(LoopI).num > 0 Then
            If MapNpc(mapnum).NPC(LoopI).x = x Then
                If MapNpc(mapnum).NPC(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapnum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, mapnum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Sub SpawnAllMapGlobalEvents()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnGlobalEvents(i)
    Next

End Sub

Sub SpawnGlobalEvents(ByVal mapnum As Long)
    Dim i As Long, z As Long
    
    If Map(mapnum).EventCount > 0 Then
        TempEventMap(mapnum).EventCount = 0
        ReDim TempEventMap(mapnum).Events(0)
        For i = 1 To Map(mapnum).EventCount
            TempEventMap(mapnum).EventCount = TempEventMap(mapnum).EventCount + 1
            ReDim Preserve TempEventMap(mapnum).Events(0 To TempEventMap(mapnum).EventCount)
            If Map(mapnum).Events(i).PageCount > 0 Then
                If Map(mapnum).Events(i).Global = 1 Then
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).x = Map(mapnum).Events(i).x
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).y = Map(mapnum).Events(i).y
                    If Map(mapnum).Events(i).Pages(1).GraphicType = 1 Then
                        Select Case Map(mapnum).Events(i).Pages(1).GraphicY
                            Case 0
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_DOWN
                            Case 1
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_LEFT
                            Case 2
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_RIGHT
                            Case 3
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_UP
                        End Select
                    Else
                        TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_DOWN
                    End If
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).active = 1
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveType = Map(mapnum).Events(i).Pages(1).MoveType
                    
                    If TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveType = 2 Then
                        TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRouteCount = Map(mapnum).Events(i).Pages(1).MoveRouteCount
                        ReDim TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRoute(0 To Map(mapnum).Events(i).Pages(1).MoveRouteCount)
                        For z = 0 To Map(mapnum).Events(i).Pages(1).MoveRouteCount
                            TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRoute(z) = Map(mapnum).Events(i).Pages(1).MoveRoute(z)
                        Next
                    End If
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).RepeatMoveRoute = Map(mapnum).Events(i).Pages(1).RepeatMoveRoute
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).IgnoreIfCannotMove = Map(mapnum).Events(i).Pages(1).IgnoreMoveRoute
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveFreq = Map(mapnum).Events(i).Pages(1).MoveFreq
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveSpeed = Map(mapnum).Events(i).Pages(1).MoveSpeed
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).WalkThrough = Map(mapnum).Events(i).Pages(1).WalkThrough
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).FixedDir = Map(mapnum).Events(i).Pages(1).DirFix
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).WalkingAnim = Map(mapnum).Events(i).Pages(1).WalkAnim
                    
                End If
            End If
        Next
    End If

End Sub

Function CanNpcMove(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(mapnum).NPC(mapNpcNum).x
    y = MapNpc(mapnum).NPC(mapNpcNum).y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapnum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).NPC(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(mapNpcNum).x) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                n = Map(mapnum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).NPC(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(mapNpcNum).x) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapnum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(mapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(mapNpcNum).x - 1) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                n = Map(mapnum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(mapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).NPC(i).num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(mapNpcNum).x + 1) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(mapnum).NPC(mapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNpc(mapnum).NPC(mapNpcNum).y = MapNpc(mapnum).NPC(mapNpcNum).y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapnum).NPC(mapNpcNum).y = MapNpc(mapnum).NPC(mapNpcNum).y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapnum).NPC(mapNpcNum).x = MapNpc(mapnum).NPC(mapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapnum).NPC(mapNpcNum).x = MapNpc(mapnum).NPC(mapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub

Sub NpcDir(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapnum).NPC(mapNpcNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapNpcNum
    Buffer.WriteLong Dir
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
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

Sub ClearTempTile(ByVal mapnum As Long)
    Dim y As Long
    Dim x As Long
    TempTile(mapnum).DoorTimer = 0
    ReDim TempTile(mapnum).DoorOpen(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            TempTile(mapnum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapnum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            If Map(mapnum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapnum).ResourceData(0 To Resource_Count)
                ResourceCache(mapnum).ResourceData(Resource_Count).x = x
                ResourceCache(mapnum).ResourceData(Resource_Count).y = y
                ResourceCache(mapnum).ResourceData(Resource_Count).cur_health = Resource(Map(mapnum).Tile(x, y).Data1).health
            End If

        Next
    Next

    ResourceCache(mapnum).Resource_Count = Resource_Count
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

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long) ' Move spellslot
    Dim OldNum As Long
    Dim NewNum As Long
    Dim i As Long
    Dim oldExp As Long
    Dim oldLv As Byte
    Dim newExp As Long
    Dim newLv As Byte

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(index, oldSlot)
    NewNum = GetPlayerSpell(index, newSlot)
    SetPlayerSpell index, oldSlot, NewNum
    SetPlayerSpell index, newSlot, OldNum
    
    For i = 1 To MAX_PLAYER_SPELLS
        If i = oldSlot Then
            oldExp = Player(index).skillEXP(i)
            oldLv = Player(index).skillLV(i)
        End If
        
        If i = newSlot Then
            newExp = Player(index).skillEXP(i)
            newLv = Player(index).skillLV(i)
        End If
    Next
    
    For i = 1 To MAX_PLAYER_SPELLS
        If i = oldSlot Then
            Player(index).skillEXP(i) = newExp
            Player(index).skillLV(i) = newLv
        End If
        
        If i = newSlot Then
            Player(index).skillEXP(i) = oldExp
            Player(index).skillLV(i) = oldLv
        End If
    Next
    
    SendPlayerData index
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
    ' Bind Working
        If Item(GetPlayerEquipment(index, EqSlot)).BindType = 2 Then Exit Sub
        
        GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0
        PlayerMsg index, "คุณได้ถอดไอเทม " & CheckGrammar(Item(GetPlayerEquipment(index, EqSlot)).Name), Yellow
        ' send the sound
            If GetPlayerEquipment(index, Shield) <> Player(index).WieldDagger Then
                Call SetPlayerdagger(index, 0)
            End If
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
        PlayerMsg index, "ช่องเก็บของเต็ม ไม่สามารถถอดไอเทมได้.", BrightRed
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
        If Caps Then CheckGrammar = "" & Word Else CheckGrammar = "" & Word
    Else
        If Caps Then CheckGrammar = "" & Word Else CheckGrammar = "" & Word
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

Public Function rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    rand = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal index As Long)
Dim PartyNum As Long, i As Long

    PartyNum = TempPlayer(index).inParty
    If PartyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers PartyNum
        ' make sure there's more than 2 people
        If Party(PartyNum).MemberCount > 2 Then
        
            ' check if leader
            If Party(PartyNum).Leader = index Then

                ' leave party
                PartyMsg PartyNum, GetPlayerName(index) & " ได้ออกจากปาร์ตี้.", Pink
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(PartyNum).Member(i) = index Then
                        Party(PartyNum).Member(i) = 0
                        TempPlayer(index).inParty = 0
                        TempPlayer(index).partyInvite = 0
                        Exit For
                        End If
                Next
                
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(PartyNum).Member(i) > 0 And Party(PartyNum).Member(i) <> index Then
                        Party(PartyNum).Leader = Party(PartyNum).Member(i)
                        PartyMsg PartyNum, "หัวหน้าของปาร์ตี้ถูกเปลี่ยนตำแหน่ง.", BrightCyan

                        'PartyMsg partyNum, GetPlayerName(i) & " ได้เป็นหัวหน้าปาร์ตี้คนใหม่.", BrightCyan
                        Exit For
                    End If
                Next
                
                ' recount party
                Party_CountMembers PartyNum
                ' set update to all
                SendPartyUpdate PartyNum
                ' send clear to player
                SendPartyUpdateTo index
                
            Else
                ' not the leader, just leave
                PartyMsg PartyNum, GetPlayerName(index) & " ได้ออกจากปาร์ตี้.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(PartyNum).Member(i) = index Then
                        Party(PartyNum).Member(i) = 0
                        TempPlayer(index).inParty = 0
                        TempPlayer(index).partyInvite = 0
                        Exit For
                    End If
                Next
                
                ' recount party
                Party_CountMembers PartyNum
                ' set update to all
                SendPartyUpdate PartyNum
                ' send clear to player
                SendPartyUpdateTo index
                
            End If
        Else
            ' find out how many members we have
            Party_CountMembers PartyNum
            ' only 2 people, disband
            PartyMsg PartyNum, "ปาร์ตี้ถูกยุบ.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                index = Party(PartyNum).Member(i)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).partyInvite = 0
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                End If
            Next
            ' clear out the party itself
            ClearParty PartyNum
        End If
    End If
    
End Sub


Public Sub Party_Invite(ByVal index As Long, ByVal OtherPlayer As Long)
    Dim PartyNum As Long, i As Long
    
    ' Make sure they're not in a party
    If TempPlayer(OtherPlayer).inParty > 0 Then
        ' They're already in a party
        PlayerMsg index, "ผู้เล่นนี้ อยู่ในปาร์ตี้หรือมีปาร์ตี้อยู่แล้ว !", BrightRed
        Exit Sub
    End If
    
    ' make sure they're not busy
    If TempPlayer(OtherPlayer).partyInvite > 0 Or TempPlayer(OtherPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg index, "ผู้เล่นนี้อยู่ในระหว่างรอการตอบรับคำขอใด ๆ อยู่.", BrightRed
        ' exit out early
        Exit Sub
    End If
    
    ' Check if there doing another action
    If IsPlayerBusy(index, OtherPlayer) Then Exit Sub
    
    ' Check if we're in a party
    If TempPlayer(index).inParty > 0 Then
        PartyNum = TempPlayer(index).inParty
        ' Make sure we're the leader
        If Party(PartyNum).Leader = index Then
            ' Got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(PartyNum).Member(i) = 0 Then
                    ' Send the invitation
                    SendPartyInvite OtherPlayer, index
                    
                    ' Set the invite target
                    TempPlayer(OtherPlayer).partyInvite = index
                    
                    ' Let them know
                    PlayerMsg index, "ส่งคำชวนเข้าร่วมปาร์ตี้สำเร็จ.", Pink
                    Exit Sub
                End If
            Next
            
            ' No room
            PlayerMsg index, "ปาร์ตี้เต็ม !", BrightRed
            Exit Sub
        Else
            ' Not the leader
            PlayerMsg index, "คุณไม่ใช่หัวหน้าปาร์ตี้ !", BrightRed
            Exit Sub
        End If
    Else
        ' Not in a party - doesn't matter
        SendPartyInvite OtherPlayer, index
        
        ' Set the invite target
        TempPlayer(OtherPlayer).partyInvite = index
        
        ' Let them know
        PlayerMsg index, "ส่งคำชวนเข้าร่วมปาร์ตี้สำเร็จ.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal OtherPlayer As Long)
    Dim PartyNum As Byte, i As Long, n As Long

    ' Check if already in a party
    If TempPlayer(index).inParty > 0 Then
        ' Get the PartyNumber
        PartyNum = TempPlayer(index).inParty
        ' Got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(PartyNum).Member(i) = 0 Then
                ' Clear party invite
                TempPlayer(OtherPlayer).partyInvite = 0
                ' Add to the party
                Party(PartyNum).Member(i) = OtherPlayer
                ' Recount party
                Party_CountMembers PartyNum
                ' Send update to all - including new player
                SendPartyUpdate PartyNum
                'SendPartyVitals PartyNum, OtherPlayer
                
                For n = 1 To MAX_PARTY_MEMBERS
                    SendPartyVitals PartyNum, n
                Next
                
                ' Let everyone know they've joined
                PartyMsg PartyNum, GetPlayerName(OtherPlayer) & " ได้เข้าร่วมปาร์ตี้.", Pink
                
                ' Add them in
                TempPlayer(OtherPlayer).inParty = PartyNum
                Exit Sub
            End If
        Next
        
        ' No empty slots - let them know
        PlayerMsg index, "ปาร์ตี้เต็ม !", BrightRed
        PlayerMsg OtherPlayer, "ปาร์ตี้เต็ม !", BrightRed
        Exit Sub
    Else
        ' Not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' Find blank party
            If Not Party(i).Leader > 0 Then
                PartyNum = i
                Exit For
            End If
        Next
        
        ' Create the party
        Party(PartyNum).MemberCount = 2
        Party(PartyNum).Leader = index
        Party(PartyNum).Member(1) = index
        Party(PartyNum).Member(2) = OtherPlayer
        SendPartyUpdate PartyNum
        SendPartyVitals PartyNum, index
        SendPartyVitals PartyNum, OtherPlayer
        
        ' Let them know it's created
        PartyMsg PartyNum, "สร้างปาร์ตี้สำเร็จ.", BrightGreen
        PartyMsg PartyNum, GetPlayerName(index) & " ได้เข้าร่วมปาร์ตี้.", Pink
        PartyMsg PartyNum, GetPlayerName(OtherPlayer) & " ได้เข้าร่วมปาร์ตี้.", Pink
        
        ' Clear the invitation
        TempPlayer(OtherPlayer).partyInvite = 0
       
       ' Add them to the party
        TempPlayer(OtherPlayer).inParty = PartyNum
        TempPlayer(index).inParty = PartyNum
    End If
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal OtherPlayer As Long)
    
    'If IsPlaying(Index) Then
    '    PlayerMsg Index, GetPlayerName(OtherPlayer) & " ได้ออกจากปาร์ตี้ !", BrightRed
    'End If
    
    PlayerMsg index, GetPlayerName(OtherPlayer) & " ได้ปฏิเสธการเข้าร่วมปาร์ตี้.", BrightRed
    PlayerMsg OtherPlayer, "คุณได้ปฏิเสธการเข้าร่วมปาร์ตี้.", BrightRed
    
    ' Clear the invitation
    TempPlayer(OtherPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal PartyNum As Long)
    Dim i As Long, highIndex As Long, x As Long
    
    ' Find the high Index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(PartyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    
    ' Count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' We've got a blank member
        If Party(PartyNum).Member(i) = 0 Then
            ' Is it lower than the high Index?
            If i < highIndex Then
                ' Move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(PartyNum).Member(x) = Party(PartyNum).Member(x + 1)
                    Party(PartyNum).Member(x + 1) = 0
                Next
            Else
                ' Not lower - highIndex is count
                Party(PartyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        
        ' Check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(PartyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    
    ' If we're here it means that we need to re-count again
    Party_CountMembers PartyNum
    
End Sub

Public Sub Party_ShareExp(ByVal PartyNum As Long, ByVal exp As Long, ByVal index As Long)
    Dim ExpShare As Long, i As Long, tmpIndex As Long
    
    ' Find out the equal share
    ExpShare = (exp / 2) + (exp \ Party(PartyNum).MemberCount)
    
    If ExpShare < 1 Then
        ExpShare = 1
    End If
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(PartyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                'If GetPlayerMap(tmpIndex) = GetPlayerMap(index) Then
                    ' give them their share
                    If Player(tmpIndex).Level < MAX_LEVELS Then
                        GivePlayerEXP tmpIndex, ExpShare
                        Call PlayerMsg(tmpIndex, "คุณได้รับ " & ExpShare & "  exp จากปาร์ตี้.", Yellow)
                    Else
                        Call PlayerMsg(tmpIndex, "คุณไม่ได้รับ exp จากปาร์ตี้ เนื่องจากมีเลเวลสูงสุดแล้ว.", Yellow)
                    End If
                'End If
            End If
        End If
    Next
    
    'PartyMsg PartyNum, "คุณได้รับ " & ExpShare & " exp จากปาร์ตี้.", Pink

End Sub

Public Sub GivePlayerEXP(ByVal index As Long, ByVal exp As Long)
    
    ' give the exp
    Call SetPlayerExp(index, GetPlayerExp(index) + exp)
    SendEXP index
    
    If exp > 0 Then
        SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32) - 8
    Else
        SendActionMsg GetPlayerMap(index), "!", Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32) - 8
    End If
    
    ' check if we've leveled
    CheckPlayerLevelUp index
    
End Sub

' projectiles
Public Sub HandleProjecTile(ByVal index As Long, ByVal PlayerProjectile As Long)
Dim x As Long, y As Long, i As Long
Dim Damage As Long
Dim npcNum As Long
Dim mapnum As Long
Dim BlockAmount As Long
Dim DEFP As Long, DEFNPC As Long
Dim NDEF As Boolean

Damage = 0
NDEF = False

    ' check for subscript out of range
    If index < 1 Or index > MAX_PLAYERS Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
        
    ' check to see if it's time to move the Projectile
    If GetTickCount > TempPlayer(index).Projectile(PlayerProjectile).TravelTime Then
        With TempPlayer(index).Projectile(PlayerProjectile)
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
    
    x = TempPlayer(index).Projectile(PlayerProjectile).x
    y = TempPlayer(index).Projectile(PlayerProjectile).y
    
    ' check if left map
    If x > Map(GetPlayerMap(index)).MaxX Or y > Map(GetPlayerMap(index)).MaxY Or x < 0 Or y < 0 Then
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if hit player
    
    'Projectile scaling formula
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
                
                If CanPlayerDodge(i) Then
                    SendActionMsg mapnum, "หลบไร้ผล !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                    SendAnimation mapnum, DODGE_ANIM, (Player(i).x), (Player(i).y)
               '     ClearProjectile index, PlayerProjectile
               '     Exit Sub
                End If
                
                Damage = TempPlayer(index).Projectile(PlayerProjectile).Damage + GetPlayerDamage(index)
                ' randomise from 1 to max hit
                DEFP = GetPlayerDef(i)
                
                ' ระบบเจาะเกราะ
                If GetPlayerEquipment(index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).NDEF > 0 Then
                        NDEF = True
                    End If
                End If
                
                ' x1.2 Critical ! +ระบบเพิ่มความแรงคริติคอล
                If CanPlayerCrit(index) Then
                    If GetPlayerEquipment(index, Weapon) > 0 Then
                        Damage = Damage * GetPlayerCritDamage(index, False)
                    ElseIf GetPlayerEquipment(index, Shield) > 0 Then
                        Damage = Damage * GetPlayerCritDamage(index, True)
                    Else
                        Damage = Damage * GetPlayerCritDamage(index, False)
                    End If
                    
                    SendActionMsg mapnum, "คริติคอล !", Yellow, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                    SendAnimation mapnum, CRIT_ANIM, (Player(i).x), (Player(i).y)
                Else
                    ' ระบบเจาะเกราะ
                    If NDEF = True Then
                        Damage = Damage - (DEFP - ((DEFP * Item(GetPlayerEquipment(index, Weapon)).NDEF) / 100))
                    Else
                        Damage = Damage - DEFP
                    End If
                End If
                
                If CanPlayerBlock(i) Then
                    SendActionMsg mapnum, "ป้องกัน !", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                    SendAnimation mapnum, PARRY_ANIM, (Player(i).x), (Player(i).y)
                    ClearProjectile index, PlayerProjectile
                    Exit Sub
                End If
                
                If Damage > 0 Then
                    Call PlayerReflectPlayer(index, i, Damage, 0)
                    
                    ' ระบบ Vampire
                If GetPlayerEquipment(index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).Vampire > 0 Then
            
                    ' แก้ไขบัคดูดเลือดเกิน !!
                    If GetPlayerMaxVital(index, HP) > Player(index).Vital(Vitals.HP) + Int((Damage * (Item(GetPlayerEquipment(index, Weapon)).Vampire / 100))) Then
                            Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Int((Damage * (Item(GetPlayerEquipment(index, Weapon)).Vampire / 100)))
                        Else
                            Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, HP)
                     End If
                        
                        ' send vitals to party if in one
                        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                        SendActionMsg GetPlayerMap(index), "+" & Int((Damage * (Item(GetPlayerEquipment(index, Weapon)).Vampire / 100))), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                        SendAnimation mapnum, Vampire_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
                        SendVital index, HP
                    
                    End If
                End If
        
                Else
                    SendAnimation mapnum, PARRY_ANIM, (Player(i).x), (Player(i).y)
                    SendActionMsg mapnum, "อ่อนหัด !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                    ' Call PlayerMsg(index, "การโจมตีเบาเกินไป.", BrightRed)
                End If
                
                ClearProjectile index, PlayerProjectile
                Exit Sub
                        Exit Sub
                    Else
                        Call PlayerMsg(index, "ไม่สามารถโจมตีศัตรูที่ตายแล้วได้.", BrightRed)
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    ' check for npc hit
    For i = 1 To MAX_MAP_NPCS
        mapnum = GetPlayerMap(index)
        npcNum = MapNpc(mapnum).NPC(i).num
        If x = MapNpc(GetPlayerMap(index)).NPC(i).x And y = MapNpc(GetPlayerMap(index)).NPC(i).y Then
            ' they're hit, remove it and deal that damage ;)
            
            If CanPlayerAttackNpc(index, i, True) Then
            
                If CanNpcDodge(mapnum, npcNum, i) Then
                    SendActionMsg mapnum, "หลบไร้ผล !", White, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                    SendAnimation mapnum, DODGE_ANIM, (MapNpc(mapnum).NPC(i).x), (MapNpc(mapnum).NPC(i).y)
              '      ClearProjectile index, PlayerProjectile
              '      Exit Sub
                End If
                
                Damage = TempPlayer(index).Projectile(PlayerProjectile).Damage + GetPlayerDamage(index)
                ' randomise from 1 to max hit
                DEFNPC = NPC(npcNum).Def
                
                ' ระบบเจาะเกราะ
                 If GetPlayerEquipment(index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).NDEF > 0 Then
                        NDEF = True
                    End If
                End If
                
                ' x1.2 Critical ! +ระบบเพิ่มความแรงคริติคอล
                If CanPlayerCrit(index) Then
                    If GetPlayerEquipment(index, Weapon) > 0 Then
                        Damage = Damage * GetPlayerCritDamage(index, False)
                    ElseIf GetPlayerEquipment(index, Shield) > 0 Then
                        Damage = Damage * GetPlayerCritDamage(index, True)
                    Else
                        Damage = Damage * GetPlayerCritDamage(index, False)
                    End If
                        SendActionMsg mapnum, "คริติคอล !", BrightGreen, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                        SendAnimation mapnum, CRIT_ANIM, (MapNpc(mapnum).NPC(i).x), (MapNpc(mapnum).NPC(i).y)
                Else
                    ' ระบบเจาะเกราะ
                    If NDEF = True Then
                        Damage = Damage - (DEFNPC - ((DEFNPC * Item(GetPlayerEquipment(index, Weapon)).NDEF) / 100))
                    Else
                        Damage = Damage - DEFNPC
                    End If
                End If
                
                If CanNpcParry(npcNum) Then
                    SendActionMsg mapnum, "ป้องกัน !", Yellow, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                    SendAnimation mapnum, PARRY_ANIM, (MapNpc(mapnum).NPC(i).x), (MapNpc(mapnum).NPC(i).y)
                    ClearProjectile index, PlayerProjectile
                    Exit Sub
                End If
                
                If Damage > 0 Then
                    Call PlayerAttackNpc(index, i, Damage)
                    
                    ' ระบบ Vampire
                If GetPlayerEquipment(index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).Vampire > 0 Then
            
                    ' แก้ไขบัคดูดเลือดเกิน !!
                    If GetPlayerMaxVital(index, HP) > Player(index).Vital(Vitals.HP) + Int((Damage * (Item(GetPlayerEquipment(index, Weapon)).Vampire / 100))) Then
                            Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Int((Damage * (Item(GetPlayerEquipment(index, Weapon)).Vampire / 100)))
                        Else
                            Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, HP)
                        End If
                
                        ' send vitals to party if in one
                        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                        SendActionMsg GetPlayerMap(index), "+" & Int((Damage * (Item(GetPlayerEquipment(index, Weapon)).Vampire / 100))), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                        SendAnimation mapnum, Vampire_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
                        SendVital index, HP
                    End If
                End If
                    
                Else
                    SendActionMsg mapnum, "อ่อนหัด !", BrightRed, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                    ' Call PlayerMsg(index, "การโจมตีเบาเกินไป.", BrightRed)
                    SendAnimation mapnum, PARRY_ANIM, (MapNpc(mapnum).NPC(i).x), (MapNpc(mapnum).NPC(i).y)
                End If
                
                ClearProjectile index, PlayerProjectile
                Exit Sub
            Else
                'Call PlayerMsg(index, "ไม่สามารถโจมตีศัตรูที่ตายแล้วได้.", BrightRed)
                'ClearProjectile index, PlayerProjectile
                Exit Sub
            End If
        End If
    Next
    
    'Projectile scaling formula
    
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
    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).Target = index
End Sub

'makes the pet wander around the map
Sub PetWander(ByVal index As Long)
    If TempPlayer(index).TempPetSlot < 1 Then Exit Sub

    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).targetType = TARGET_TYPE_NONE
    MapNpc(GetPlayerMap(index)).NPC(TempPlayer(index).TempPetSlot).Target = 0
End Sub

'Clear the npc from the map
Sub PetDisband(ByVal index As Long, ByVal mapnum As Long)
    Dim i As Long
    Dim j As Long
    Dim n As Integer

    If TempPlayer(index).TempPetSlot < 1 Then Exit Sub

    ' จับเวลากดปุ่ม

    'Cache the Pets for players logging on [Remove Number from array]
    PetMapCache(mapnum).Pet(Player(index).Pet.CNum) = 0
    
    Call ClearSingleMapNpc(TempPlayer(index).TempPetSlot, mapnum)
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
    
    For n = 1 To Player_HighIndex
        If IsPlaying(n) Then
            If GetPlayerMap(n) = GetPlayerMap(index) Then
                Call PlayerWarp(n, GetPlayerMap(index), GetPlayerX(n), GetPlayerY(n))
            End If
        End If
    Next
    
End Sub


Sub SpawnPet(ByVal index As Long, ByVal mapnum As Long, npcNum As Long)
    Dim PlayerMap As Long
    Dim i As Integer
    Dim PetSlot As Byte
    
    'Prevent multiple pets for the same owner
    If TempPlayer(index).TempPetSlot > 0 Then Exit Sub
    
    PlayerMap = GetPlayerMap(index)
    PetSlot = 0
    
    For i = 1 To MAX_MAP_NPCS
        'If Map(PlayerMap).Npc(i) = 0 Then
        If MapNpc(PlayerMap).NPC(i).SpawnWait = 0 And MapNpc(PlayerMap).NPC(i).num = 0 Then
            PetSlot = i
            Exit For
        End If
    Next
    
    If PetSlot = 0 Then
        Call PlayerMsg(index, "ไม่สามารถเรียกสัตว์เลี้ยงในแผนที่นี้ได้ หรือ Slot สัตว์เลี้ยงของแผนที่เต็มแล้ว !", Red)
        Exit Sub
    End If

    'create the pet for the map
    Map(PlayerMap).NPC(PetSlot) = npcNum
    MapNpc(PlayerMap).NPC(PetSlot).num = npcNum
    'set its Pet Data
    MapNpc(PlayerMap).NPC(PetSlot).IsPet = YES
    MapNpc(PlayerMap).NPC(PetSlot).PetData.Name = GetPlayerName(index) & " " & NPC(npcNum).Name
    MapNpc(PlayerMap).NPC(PetSlot).PetData.Owner = index
    
    'If Pet doesn't exist with player, link it to the player
    If Player(index).Pet.SpriteNum <> npcNum Then
        Player(index).Pet.SpriteNum = npcNum
        Player(index).Pet.Name = GetPlayerName(index) & " " & NPC(npcNum).Name
    End If
    
    TempPlayer(index).TempPetSlot = PetSlot
       
    'cache the map for sending
    Call MapCache_Create(PlayerMap)

    'Cache the Pets for players logging on [Add new Number to array]
    
    For i = 1 To 10
        If PetMapCache(PlayerMap).Pet(i) = 0 Then
            PetMapCache(PlayerMap).Pet(i) = PetSlot
            Player(index).Pet.CNum = i
            Exit For
        End If
    Next
    For i = 1 To 10
        If PetMapCache(PlayerMap).Pet(i) > 0 Then
            Call NPCCache_Create(index, Player(index).Map, PetMapCache(Player(index).Map).Pet(i))
        End If
    Next

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

TempPlayer(index).havePet = True
    
End Sub

Function CanEventMove(index As Long, ByVal mapnum As Long, x As Long, y As Long, eventID As Long, WalkThrough As Long, ByVal Dir As Byte, Optional globalevent As Boolean = False) As Boolean
    Dim i As Long
    Dim n As Long, z As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    CanEventMove = True
    
    

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapnum).Tile(x, y - 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If
                
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = y - 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x) And (MapNpc(mapnum).NPC(i).y = y - 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x) And (TempEventMap(mapnum).Events(z).y = y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_UP + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                n = Map(mapnum).Tile(x, y + 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = y + 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x) And (MapNpc(mapnum).NPC(i).y = y + 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x) And (TempEventMap(mapnum).Events(z).y = y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_DOWN + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapnum).Tile(x - 1, y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x - 1) And (GetPlayerY(i) = y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x - 1) And (MapNpc(mapnum).NPC(i).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x - 1) And (TempEventMap(mapnum).Events(z).y = y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x - 1) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_LEFT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                n = Map(mapnum).Tile(x + 1, y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x + 1) And (GetPlayerY(i) = y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(i).x = x + 1) And (MapNpc(mapnum).NPC(i).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x + 1) And (TempEventMap(mapnum).Events(z).y = y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x + 1) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_RIGHT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

    End Select

End Function

Sub EventDir(playerindex As Long, ByVal mapnum As Long, ByVal eventID As Long, ByVal Dir As Long, Optional globalevent As Boolean = False)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent Then
        If Map(mapnum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(mapnum).Events(eventID).Dir = Dir
    Else
        If Map(mapnum).Events(eventID).Pages(TempPlayer(playerindex).EventMap.EventPages(eventID).pageID).DirFix = 0 Then TempPlayer(playerindex).EventMap.EventPages(eventID).Dir = Dir
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventDir
    Buffer.WriteLong eventID
    If globalevent Then
        Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
    Else
        Buffer.WriteLong TempPlayer(playerindex).EventMap.EventPages(eventID).Dir
    End If
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub EventMove(index As Long, mapnum As Long, ByVal eventID As Long, ByVal Dir As Long, movementspeed As Long, Optional globalevent As Boolean = False)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent Then
        If Map(mapnum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(mapnum).Events(eventID).Dir = Dir
    Else
        If Map(mapnum).Events(eventID).Pages(TempPlayer(index).EventMap.EventPages(eventID).pageID).DirFix = 0 Then TempPlayer(index).EventMap.EventPages(eventID).Dir = Dir
    End If

    Select Case Dir
        Case DIR_UP
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).y = TempEventMap(mapnum).Events(eventID).y - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).y = TempPlayer(index).EventMap.EventPages(eventID).y - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
            
        Case DIR_DOWN
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).y = TempEventMap(mapnum).Events(eventID).y + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).y = TempPlayer(index).EventMap.EventPages(eventID).y + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
        Case DIR_LEFT
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).x = TempEventMap(mapnum).Events(eventID).x - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).x = TempPlayer(index).EventMap.EventPages(eventID).x - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
        Case DIR_RIGHT
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).x = TempEventMap(mapnum).Events(eventID).x + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).x = TempPlayer(index).EventMap.EventPages(eventID).x + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
    End Select

End Sub
