Attribute VB_Name = "modEventLogic"
Option Explicit
Public Sub RemoveDeadEvents()
Dim i As Long, mapnum As Long, Buffer As clsBuffer, x As Long, id As Long, page As Long
    'See if we should remove any events....
    For i = 1 To Player_HighIndex
        If TempPlayer(i).EventMap.CurrentEvents > 0 Then
            mapnum = GetPlayerMap(i)
            For x = 1 To TempPlayer(i).EventMap.CurrentEvents
                id = TempPlayer(i).EventMap.EventPages(x).eventID
                page = TempPlayer(i).EventMap.EventPages(x).pageID
                If Map(mapnum).Events(id).PageCount >= page Then
                
                    'See if there is any reason to delete this event....
                    'In other words, go back through conditions and make sure they all check up.
                    If TempPlayer(i).EventMap.EventPages(x).Visible = 1 Then
                        If Map(mapnum).Events(id).Pages(page).chkHasItem = 1 Then
                            If HasItem(i, Map(mapnum).Events(id).Pages(page).HasItemIndex) = 0 Then
                                TempPlayer(i).EventMap.EventPages(x).Visible = 0
                            End If
                        End If
                        
                        If Map(mapnum).Events(id).Pages(page).chkSelfSwitch = 1 Then
                            If Map(mapnum).Events(id).Pages(page).SelfSwitchCompare = 0 Then
                                If Map(mapnum).Events(id).SelfSwitches(Map(mapnum).Events(id).Pages(page).SelfSwitchIndex) = 0 Then
                                    TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                End If
                            Else
                                If Map(mapnum).Events(id).SelfSwitches(Map(mapnum).Events(id).Pages(page).SelfSwitchIndex) = 1 Then
                                    TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                End If
                            End If
                        End If
                        
                        If Map(mapnum).Events(id).Pages(page).chkVariable = 1 Then
                            Select Case Map(mapnum).Events(id).Pages(page).VariableCompare
                                Case 0
                                    If Player(i).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) <> Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 1
                                    If Player(i).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) < Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 2
                                    If Player(i).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) > Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 3
                                    If Player(i).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) <= Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 4
                                    If Player(i).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) >= Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                                Case 5
                                    If Player(i).Variables(Map(mapnum).Events(id).Pages(page).VariableIndex) = Map(mapnum).Events(id).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                    End If
                            End Select
                        End If
                        
                        If Map(mapnum).Events(id).Pages(page).chkSwitch = 1 Then
                            If Map(mapnum).Events(id).Pages(page).SwitchCompare = 1 Then
                                If Player(i).Switches(Map(mapnum).Events(id).Pages(page).SwitchIndex) = 1 Then
                                    TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                End If
                            Else
                                If Player(i).Switches(Map(mapnum).Events(id).Pages(page).SwitchIndex) = 0 Then
                                    TempPlayer(i).EventMap.EventPages(x).Visible = 0
                                End If
                            End If
                        End If
                        
                        If TempPlayer(i).EventMap.EventPages(x).Visible = 0 Then
                            Set Buffer = New clsBuffer
                            Buffer.WriteLong SSpawnEvent
                            Buffer.WriteLong id
                            With TempPlayer(i).EventMap.EventPages(x)
                                Buffer.WriteString Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Name
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
                                Buffer.WriteLong Map(mapnum).Events(id).Pages(page).WalkAnim
                                Buffer.WriteLong Map(mapnum).Events(id).Pages(page).DirFix
                                Buffer.WriteLong Map(mapnum).Events(id).Pages(page).WalkThrough
                                Buffer.WriteLong Map(mapnum).Events(id).Pages(page).ShowName
                            End With
                            SendDataTo i, Buffer.ToArray
                            Set Buffer = Nothing
                        End If
                    End If
                End If
            Next
        End If
    Next
    
End Sub

Public Sub SpawnNewEvents()
Dim Buffer As clsBuffer, pageID As Long, id As Long, compare As Long, i As Long, mapnum As Long, x As Long, z As Long, spawnevent As Boolean, p As Long
    'That was only removing events... now we gotta worry about spawning them again, luckily, it is almost the same exact thing, but backwards!
    For i = 1 To Player_HighIndex
        If TempPlayer(i).EventMap.CurrentEvents > 0 Then
            mapnum = GetPlayerMap(i)
            For x = 1 To TempPlayer(i).EventMap.CurrentEvents
                id = TempPlayer(i).EventMap.EventPages(x).eventID
                pageID = TempPlayer(i).EventMap.EventPages(x).pageID
                
                'See if there is any reason to delete this event....
                'In other words, go back through conditions and make sure they all check up.
                For z = Map(mapnum).Events(id).PageCount To 1 Step -1
                        
                    spawnevent = True
                        
                    If Map(mapnum).Events(id).Pages(z).chkHasItem = 1 Then
                        If HasItem(i, Map(mapnum).Events(id).Pages(z).HasItemIndex) = 0 Then
                            spawnevent = False
                        End If
                    End If
                        
                    If Map(mapnum).Events(id).Pages(z).chkSelfSwitch = 1 Then
                        If Map(mapnum).Events(id).Pages(z).SelfSwitchCompare = 0 Then
                            compare = 1
                        Else
                            compare = 0
                        End If
                        If Map(mapnum).Events(id).Global = 1 Then
                            If Map(mapnum).Events(id).SelfSwitches(Map(mapnum).Events(id).Pages(z).SelfSwitchIndex) <> compare Then
                                spawnevent = False
                            End If
                        Else
                            If TempPlayer(i).EventMap.EventPages(id).SelfSwitches(Map(mapnum).Events(id).Pages(z).SelfSwitchIndex) <> compare Then
                                spawnevent = False
                            End If
                        End If
                    End If
                        
                    If Map(mapnum).Events(id).Pages(z).chkVariable = 1 Then
                        Select Case Map(mapnum).Events(id).Pages(z).VariableCompare
                            Case 0
                                If Player(i).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) = Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 1
                                If Player(i).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) >= Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 2
                                If Player(i).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) <= Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 3
                                If Player(i).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) > Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 4
                                If Player(i).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) < Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                            Case 5
                                If Player(i).Variables(Map(mapnum).Events(id).Pages(z).VariableIndex) <> Map(mapnum).Events(id).Pages(z).VariableCondition Then
                                    spawnevent = False
                                End If
                        End Select
                    End If
                        
                    If Map(mapnum).Events(id).Pages(z).chkSwitch = 1 Then
                        If Map(mapnum).Events(id).Pages(z).SwitchCompare = 0 Then
                            If Player(i).Switches(Map(mapnum).Events(id).Pages(z).SwitchIndex) = 0 Then
                                spawnevent = False
                            End If
                        Else
                            If Player(i).Switches(Map(mapnum).Events(id).Pages(z).SwitchIndex) = 1 Then
                                spawnevent = False
                            End If
                        End If
                    End If
                        
                    If spawnevent = True Then
                        If TempPlayer(i).EventMap.EventPages(x).Visible = 1 Then
                            If z <= pageID Then
                                spawnevent = False
                            End If
                        End If
                    End If
                        
                    If spawnevent = True Then
                        With TempPlayer(i).EventMap.EventPages(x)
                            If Map(mapnum).Events(id).Pages(z).GraphicType = 1 Then
                                Select Case Map(mapnum).Events(id).Pages(z).GraphicY
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
                            .GraphicNum = Map(mapnum).Events(id).Pages(z).Graphic
                            .GraphicType = Map(mapnum).Events(id).Pages(z).GraphicType
                            .GraphicX = Map(mapnum).Events(id).Pages(z).GraphicX
                            .GraphicY = Map(mapnum).Events(id).Pages(z).GraphicY
                            .GraphicX2 = Map(mapnum).Events(id).Pages(z).GraphicX2
                            .GraphicY2 = Map(mapnum).Events(id).Pages(z).GraphicY2
                            Select Case Map(mapnum).Events(id).Pages(z).MoveSpeed
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
                            .Position = Map(mapnum).Events(id).Pages(z).Position
                            .eventID = id
                            .pageID = z
                            .Visible = 1
                                
                            .MoveType = Map(mapnum).Events(id).Pages(z).MoveType
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(mapnum).Events(id).Pages(z).MoveRouteCount
                                ReDim .MoveRoute(0 To Map(mapnum).Events(id).Pages(z).MoveRouteCount)
                                For p = 0 To Map(mapnum).Events(id).Pages(z).MoveRouteCount
                                    .MoveRoute(p) = Map(mapnum).Events(id).Pages(z).MoveRoute(p)
                                Next
                            End If
                                
                            .RepeatMoveRoute = Map(mapnum).Events(id).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(mapnum).Events(id).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(mapnum).Events(id).Pages(z).MoveFreq
                            .MoveSpeed = Map(mapnum).Events(id).Pages(z).MoveSpeed
                                
                            .WalkThrough = Map(mapnum).Events(id).Pages(z).WalkThrough
                            .WalkingAnim = Map(mapnum).Events(id).Pages(z).WalkAnim
                            .FixedDir = Map(mapnum).Events(id).Pages(z).DirFix
                            
                        End With
                        
                        
                        
                        Set Buffer = New clsBuffer
                        Buffer.WriteLong SSpawnEvent
                        Buffer.WriteLong id
                        With TempPlayer(i).EventMap.EventPages(x)
                            Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Name
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
                            Buffer.WriteLong Map(mapnum).Events(id).Pages(z).WalkAnim
                            Buffer.WriteLong Map(mapnum).Events(id).Pages(z).DirFix
                            Buffer.WriteLong Map(mapnum).Events(id).Pages(z).WalkThrough
                            Buffer.WriteLong Map(mapnum).Events(id).Pages(z).ShowName
                        End With
                        SendDataTo i, Buffer.ToArray
                        Set Buffer = Nothing
                        GoTo nextevent
                    End If
                Next
nextevent:
            Next
        End If
    Next
    
End Sub

Public Sub ProcessEventMovement()
Dim rand As Long, x As Long, i As Long, playerID As Long, eventID As Long, WalkThrough As Long, isglobal As Boolean, mapnum As Long, actualmovespeed As Long, Buffer As clsBuffer, z As Long, sendupdate As Boolean
    'Process Movement if needed for each player/each map/each event....
    For i = 1 To MAX_MAPS
        If PlayersOnMap(i) Then
            'Manage Global Events First, then all the others.....
            If TempEventMap(i).EventCount > 0 Then
                For x = 1 To TempEventMap(i).EventCount
                    If TempEventMap(i).Events(x).active = 1 Then
                        If TempEventMap(i).Events(x).MoveTimer <= GetTickCount Then
                            'Real event! Lets process it!
                            Select Case TempEventMap(i).Events(x).MoveType
                                Case 0
                                    'Nothing, fixed position
                                Case 1 'Random, move randomly if possible...
                                    rand = Random(0, 3)
                                    If CanEventMove(0, i, TempEventMap(i).Events(x).x, TempEventMap(i).Events(x).y, x, TempEventMap(i).Events(x).WalkThrough, rand, True) Then
                                        Select Case TempEventMap(i).Events(x).MoveSpeed
                                            Case 0
                                                EventMove 0, i, x, rand, 2, True
                                            Case 1
                                                EventMove 0, i, x, rand, 3, True
                                            Case 2
                                                EventMove 0, i, x, rand, 4, True
                                            Case 3
                                                EventMove 0, i, x, rand, 6, True
                                            Case 4
                                                EventMove 0, i, x, rand, 12, True
                                            Case 5
                                                EventMove 0, i, x, rand, 24, True
                                        End Select
                                    Else
                                        EventDir 0, i, x, rand, True
                                    End If
                                Case 2 'Move Route - later
                                    With TempEventMap(i).Events(x)
                                        isglobal = True
                                        mapnum = i
                                        playerID = 0
                                        eventID = x
                                        WalkThrough = TempEventMap(i).Events(x).WalkThrough
                                        If .MoveRouteCount > 0 Then
                                            If .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 1 Then
                                                .MoveRouteStep = 0
                                            ElseIf .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 0 Then
                                                GoTo donotprocessmoveroute
                                            End If
                                            .MoveRouteStep = .MoveRouteStep + 1
                                            Select Case .MoveSpeed
                                                Case 0
                                                    actualmovespeed = 2
                                                Case 1
                                                    actualmovespeed = 3
                                                Case 2
                                                    actualmovespeed = 4
                                                Case 3
                                                    actualmovespeed = 6
                                                Case 4
                                                    actualmovespeed = 12
                                                Case 5
                                                    actualmovespeed = 24
                                            End Select
                                            Select Case .MoveRoute(.MoveRouteStep).index
                                                Case 1
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_UP, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_UP, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 2
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_DOWN, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_DOWN, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 3
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_LEFT, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_LEFT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 4
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_RIGHT, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_RIGHT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 5
                                                    z = Random(0, 3)
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 6
                                                    If isglobal = False Then
                                                        z = CanEventMoveTowardsPlayer(playerID, mapnum, eventID)
                                                        If z >= 5 Then
                                                            'No
                                                        Else
                                                            'i is the direct, lets go...
                                                            If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 7
                                                    If isglobal = False Then
                                                        z = CanEventMoveAwayFromPlayer(playerID, mapnum, eventID)
                                                        If z >= 5 Then
                                                            'No
                                                        Else
                                                            'i is the direct, lets go...
                                                            If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 8
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, .Dir, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, .Dir, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 9
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                    End Select
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 10
                                                    .MoveTimer = GetTickCount + 100
                                                Case 11
                                                    .MoveTimer = GetTickCount + 500
                                                Case 12
                                                    .MoveTimer = GetTickCount + 1000
                                                Case 13
                                                    EventDir playerID, mapnum, eventID, DIR_UP, isglobal
                                                Case 14
                                                    EventDir playerID, mapnum, eventID, DIR_DOWN, isglobal
                                                Case 15
                                                    EventDir playerID, mapnum, eventID, DIR_LEFT, isglobal
                                                Case 16
                                                    EventDir playerID, mapnum, eventID, DIR_RIGHT, isglobal
                                                Case 17
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_DOWN
                                                        Case DIR_LEFT
                                                            z = DIR_UP
                                                        Case DIR_DOWN
                                                            z = DIR_LEFT
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 18
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_LEFT
                                                        Case DIR_RIGHT
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_RIGHT
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 19
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 20
                                                    z = Random(0, 3)
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 21
                                                    If isglobal = False Then
                                                        z = GetDirToPlayer(playerID, mapnum, eventID)
                                                        EventDir playerID, mapnum, eventID, z, isglobal
                                                    End If
                                                Case 22
                                                    If isglobal = False Then
                                                        z = GetDirAwayFromPlayer(playerID, mapnum, eventID)
                                                        EventDir playerID, mapnum, eventID, z, isglobal
                                                    End If
                                                Case 23
                                                    .MoveSpeed = 0
                                                Case 24
                                                    .MoveSpeed = 1
                                                Case 25
                                                    .MoveSpeed = 2
                                                Case 26
                                                    .MoveSpeed = 3
                                                Case 27
                                                    .MoveSpeed = 4
                                                Case 28
                                                    .MoveSpeed = 5
                                                Case 29
                                                    .MoveFreq = 0
                                                Case 30
                                                    .MoveFreq = 1
                                                Case 31
                                                    .MoveFreq = 2
                                                Case 32
                                                    .MoveFreq = 3
                                                Case 33
                                                    .MoveFreq = 4
                                                Case 34
                                                    .WalkingAnim = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 35
                                                    .WalkingAnim = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 36
                                                    .FixedDir = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 37
                                                    .FixedDir = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 38
                                                    .WalkThrough = 1
                                                Case 39
                                                    .WalkThrough = 0
                                                Case 40
                                                    .Position = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 41
                                                    .Position = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 42
                                                    .Position = 2
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 43
                                                    .GraphicType = .MoveRoute(.MoveRouteStep).Data1
                                                    .GraphicNum = .MoveRoute(.MoveRouteStep).Data2
                                                    .GraphicX = .MoveRoute(.MoveRouteStep).Data3
                                                    .GraphicX2 = .MoveRoute(.MoveRouteStep).data4
                                                    .GraphicY = .MoveRoute(.MoveRouteStep).data5
                                                    .GraphicY2 = .MoveRoute(.MoveRouteStep).data6
                                                    If .GraphicType = 1 Then
                                                        Select Case .GraphicY
                                                            Case 0
                                                                .Dir = DIR_DOWN
                                                            Case 1
                                                                .Dir = DIR_LEFT
                                                            Case 2
                                                                .Dir = DIR_RIGHT
                                                            Case 3
                                                                .Dir = DIR_UP
                                                        End Select
                                                    End If
                                                    'Need to Send Update to client
                                                    sendupdate = True
                                            End Select
                                            
                                            If sendupdate Then
                                                Set Buffer = New clsBuffer
                                                Buffer.WriteLong SSpawnEvent
                                                Buffer.WriteLong eventID
                                                With TempEventMap(i).Events(x)
                                                    Buffer.WriteString Map(GetPlayerMap(i)).Events(eventID).Name
                                                    Buffer.WriteLong .Dir
                                                    Buffer.WriteLong .GraphicNum
                                                    Buffer.WriteLong .GraphicType
                                                    Buffer.WriteLong .GraphicX
                                                    Buffer.WriteLong .GraphicX2
                                                    Buffer.WriteLong .GraphicY
                                                    Buffer.WriteLong .GraphicY2
                                                    Buffer.WriteLong .MoveSpeed
                                                    Buffer.WriteLong .x
                                                    Buffer.WriteLong .y
                                                    Buffer.WriteLong .Position
                                                    Buffer.WriteLong .active
                                                    Buffer.WriteLong .WalkingAnim
                                                    Buffer.WriteLong .FixedDir
                                                End With
                                                SendDataToMap i, Buffer.ToArray
                                                Set Buffer = Nothing
                                            End If
donotprocessmoveroute:
                                        End If
                                    End With
                            End Select
                            
                            Select Case TempEventMap(i).Events(x).MoveFreq
                                Case 0
                                    TempEventMap(i).Events(x).MoveTimer = GetTickCount + 4000
                                Case 1
                                    TempEventMap(i).Events(x).MoveTimer = GetTickCount + 2000
                                Case 2
                                    TempEventMap(i).Events(x).MoveTimer = GetTickCount + 1000
                                Case 3
                                    TempEventMap(i).Events(x).MoveTimer = GetTickCount + 500
                                Case 4
                                    TempEventMap(i).Events(x).MoveTimer = GetTickCount + 250
                            End Select
                        End If
                    End If
                Next
            End If
            'HOPEFULLY this will not reduce CPS too much!
        End If
        DoEvents
    Next
End Sub

Public Sub ProcessLocalEventMovement()
Dim rand As Long, x As Long, i As Long, playerID As Long, eventID As Long, WalkThrough As Long, isglobal As Boolean, mapnum As Long, actualmovespeed As Long, Buffer As clsBuffer, z As Long, sendupdate As Boolean
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            playerID = i
            If TempPlayer(i).EventMap.CurrentEvents > 0 Then
                For x = 1 To TempPlayer(i).EventMap.CurrentEvents
                    If Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Global = 1 Then GoTo nextevent1
                    If TempPlayer(i).EventMap.EventPages(x).Visible = 1 Then
                        If TempPlayer(i).EventMap.EventPages(x).MoveTimer <= GetTickCount Then
                            'Real event! Lets process it!
                            Select Case TempPlayer(i).EventMap.EventPages(x).MoveType
                                Case 0
                                    'Nothing, fixed position
                                Case 1 'Random, move randomly if possible...
                                    rand = Random(0, 3)
                                    playerID = i
                                    If CanEventMove(i, GetPlayerMap(i), TempPlayer(i).EventMap.EventPages(x).x, TempPlayer(i).EventMap.EventPages(x).y, x, TempPlayer(i).EventMap.EventPages(x).WalkThrough, rand, False) Then
                                        Select Case TempPlayer(i).EventMap.EventPages(x).MoveSpeed
                                            Case 0
                                                EventMove i, GetPlayerMap(i), x, rand, 2, False
                                            Case 1
                                                EventMove i, GetPlayerMap(i), x, rand, 3, False
                                            Case 2
                                                EventMove i, GetPlayerMap(i), x, rand, 4, False
                                            Case 3
                                                EventMove i, GetPlayerMap(i), x, rand, 6, False
                                            Case 4
                                                EventMove i, GetPlayerMap(i), x, rand, 12, False
                                            Case 5
                                                EventMove i, GetPlayerMap(i), x, rand, 24, False
                                        End Select
                                    Else
                                        EventDir 0, GetPlayerMap(i), x, rand, True
                                    End If
                                Case 2 'Move Route - later!
                                    With TempPlayer(i).EventMap.EventPages(x)
                                        isglobal = False
                                        mapnum = GetPlayerMap(i)
                                        playerID = i
                                        eventID = .eventID
                                        WalkThrough = .WalkThrough
                                        If .MoveRouteCount > 0 Then
                                            If .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 1 Then
                                                .MoveRouteStep = 0
                                            ElseIf .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 0 Then
                                                GoTo donotprocessmoveroute1
                                            End If
                                            .MoveRouteStep = .MoveRouteStep + 1
                                            Select Case .MoveSpeed
                                                Case 0
                                                    actualmovespeed = 2
                                                Case 1
                                                    actualmovespeed = 3
                                                Case 2
                                                    actualmovespeed = 4
                                                Case 3
                                                    actualmovespeed = 6
                                                Case 4
                                                    actualmovespeed = 12
                                                Case 5
                                                    actualmovespeed = 24
                                            End Select
                                            Select Case .MoveRoute(.MoveRouteStep).index
                                                Case 1
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_UP, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_UP, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 2
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_DOWN, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_DOWN, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 3
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_LEFT, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_LEFT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 4
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, DIR_RIGHT, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, DIR_RIGHT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 5
                                                    z = Random(0, 3)
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 6
                                                    If isglobal = False Then
                                                        z = CanEventMoveTowardsPlayer(playerID, mapnum, eventID)
                                                        If z >= 5 Then
                                                            'No
                                                        Else
                                                            'i is the direct, lets go...
                                                            If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 7
                                                    If isglobal = False Then
                                                        z = CanEventMoveAwayFromPlayer(playerID, mapnum, eventID)
                                                        If z >= 5 Then
                                                            'No
                                                        Else
                                                            'i is the direct, lets go...
                                                            If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                                EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 8
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, .Dir, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, .Dir, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 9
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                    End Select
                                                    If CanEventMove(playerID, mapnum, .x, .y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, mapnum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 10
                                                    .MoveTimer = GetTickCount + 100
                                                Case 11
                                                    .MoveTimer = GetTickCount + 500
                                                Case 12
                                                    .MoveTimer = GetTickCount + 1000
                                                Case 13
                                                    EventDir playerID, mapnum, eventID, DIR_UP, isglobal
                                                Case 14
                                                    EventDir playerID, mapnum, eventID, DIR_DOWN, isglobal
                                                Case 15
                                                    EventDir playerID, mapnum, eventID, DIR_LEFT, isglobal
                                                Case 16
                                                    EventDir playerID, mapnum, eventID, DIR_RIGHT, isglobal
                                                Case 17
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_DOWN
                                                        Case DIR_LEFT
                                                            z = DIR_UP
                                                        Case DIR_DOWN
                                                            z = DIR_LEFT
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 18
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_LEFT
                                                        Case DIR_RIGHT
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_RIGHT
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 19
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                    End Select
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 20
                                                    z = Random(0, 3)
                                                    EventDir playerID, mapnum, eventID, z, isglobal
                                                Case 21
                                                    If isglobal = False Then
                                                        z = GetDirToPlayer(playerID, mapnum, eventID)
                                                        EventDir playerID, mapnum, eventID, z, isglobal
                                                    End If
                                                Case 22
                                                    If isglobal = False Then
                                                        z = GetDirAwayFromPlayer(playerID, mapnum, eventID)
                                                        EventDir playerID, mapnum, eventID, z, isglobal
                                                    End If
                                                Case 23
                                                    .MoveSpeed = 0
                                                Case 24
                                                    .MoveSpeed = 1
                                                Case 25
                                                    .MoveSpeed = 2
                                                Case 26
                                                    .MoveSpeed = 3
                                                Case 27
                                                    .MoveSpeed = 4
                                                Case 28
                                                    .MoveSpeed = 5
                                                Case 29
                                                    .MoveFreq = 0
                                                Case 30
                                                    .MoveFreq = 1
                                                Case 31
                                                    .MoveFreq = 2
                                                Case 32
                                                    .MoveFreq = 3
                                                Case 33
                                                    .MoveFreq = 4
                                                Case 34
                                                    .WalkingAnim = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 35
                                                    .WalkingAnim = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 36
                                                    .FixedDir = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 37
                                                    .FixedDir = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 38
                                                    .WalkThrough = 1
                                                Case 39
                                                    .WalkThrough = 0
                                                Case 40
                                                    .Position = 0
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 41
                                                    .Position = 1
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 42
                                                    .Position = 2
                                                    'Need to send update to client
                                                    sendupdate = True
                                                Case 43
                                                    .GraphicType = .MoveRoute(.MoveRouteStep).Data1
                                                    .GraphicNum = .MoveRoute(.MoveRouteStep).Data2
                                                    .GraphicX = .MoveRoute(.MoveRouteStep).Data3
                                                    .GraphicX2 = .MoveRoute(.MoveRouteStep).data4
                                                    .GraphicY = .MoveRoute(.MoveRouteStep).data5
                                                    .GraphicY2 = .MoveRoute(.MoveRouteStep).data6
                                                    If .GraphicType = 1 Then
                                                        Select Case .GraphicY
                                                            Case 0
                                                                .Dir = DIR_DOWN
                                                            Case 1
                                                                .Dir = DIR_LEFT
                                                            Case 2
                                                                .Dir = DIR_RIGHT
                                                            Case 3
                                                                .Dir = DIR_UP
                                                        End Select
                                                    End If
                                                    'Need to Send Update to client
                                                    sendupdate = True
                                            End Select
                                            
                                            If sendupdate Then
                                                Set Buffer = New clsBuffer
                                                Buffer.WriteLong SSpawnEvent
                                                Buffer.WriteLong eventID
                                                With TempPlayer(playerID).EventMap.EventPages(eventID)
                                                    Buffer.WriteString Map(GetPlayerMap(playerID)).Events(eventID).Name
                                                    Buffer.WriteLong .Dir
                                                    Buffer.WriteLong .GraphicNum
                                                    Buffer.WriteLong .GraphicType
                                                    Buffer.WriteLong .GraphicX
                                                    Buffer.WriteLong .GraphicX2
                                                    Buffer.WriteLong .GraphicY
                                                    Buffer.WriteLong .GraphicY2
                                                    Buffer.WriteLong .MoveSpeed
                                                    Buffer.WriteLong .x
                                                    Buffer.WriteLong .y
                                                    Buffer.WriteLong .Position
                                                    Buffer.WriteLong .Visible
                                                    Buffer.WriteLong .WalkingAnim
                                                    Buffer.WriteLong .FixedDir
                                                End With
                                                SendDataTo playerID, Buffer.ToArray
                                                Set Buffer = Nothing
                                            End If
                                        End If
                                    End With
                            End Select

donotprocessmoveroute1:
                            Select Case TempPlayer(playerID).EventMap.EventPages(x).MoveFreq
                                Case 0
                                    TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 4000
                                Case 1
                                    TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 2000
                                Case 2
                                    TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 1000
                                Case 3
                                    TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 500
                                Case 4
                                    TempPlayer(playerID).EventMap.EventPages(x).MoveTimer = GetTickCount + 250
                            End Select
                        End If
                    End If
nextevent1:
                Next
            End If
        End If
        DoEvents
    Next
End Sub

Public Sub ProcessEventCommands()
Dim Buffer As clsBuffer, i As Long, x As Long, z As Long, removeeventprocess As Boolean, w As Long, v As Long, p As Long
    'Now, we process the damn things for commands :P
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            For x = 1 To TempPlayer(i).EventMap.CurrentEvents
                If TempPlayer(i).EventMap.EventPages(x).Visible Then
                    If Map(Player(i).Map).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Pages(TempPlayer(i).EventMap.EventPages(x).pageID).Trigger = 2 Then 'Parallel Process baby!
                        If TempPlayer(i).EventProcessingCount > 0 Then
                            For z = 1 To TempPlayer(i).EventProcessingCount
                                If TempPlayer(i).EventProcessing(z).eventID = TempPlayer(i).EventMap.EventPages(x).eventID And TempPlayer(i).EventMap.EventPages(x).pageID = TempPlayer(i).EventProcessing(z).pageID Then
                                    'Exit For
                                Else
                                    If z = TempPlayer(i).EventProcessingCount Then
                                        If Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Pages(TempPlayer(i).EventMap.EventPages(x).pageID).CommandListCount > 0 Then
                                            'start new event processing
                                            TempPlayer(i).EventProcessingCount = TempPlayer(i).EventProcessingCount + 1
                                            ReDim Preserve TempPlayer(i).EventProcessing(TempPlayer(i).EventProcessingCount)
                                            With TempPlayer(i).EventProcessing(TempPlayer(i).EventProcessingCount)
                                                .ActionTimer = GetTickCount
                                                .CurList = 1
                                                .CurSlot = 1
                                                .eventID = TempPlayer(i).EventMap.EventPages(x).eventID
                                                .pageID = TempPlayer(i).EventMap.EventPages(x).pageID
                                                .WaitingForResponse = 0
                                                ReDim .ListLeftOff(0 To Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Pages(TempPlayer(i).EventMap.EventPages(x).pageID).CommandListCount)
                                            End With
                                        End If
                                        Exit For
                                        
                                    End If
                                End If
                            Next
                        Else
                            If Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Pages(TempPlayer(i).EventMap.EventPages(x).pageID).CommandListCount > 0 Then
                                'Clearly need to start it!
                                TempPlayer(i).EventProcessingCount = 1
                                ReDim Preserve TempPlayer(i).EventProcessing(TempPlayer(i).EventProcessingCount)
                                With TempPlayer(i).EventProcessing(TempPlayer(i).EventProcessingCount)
                                    .ActionTimer = GetTickCount
                                    .CurList = 1
                                    .CurSlot = 1
                                    .eventID = TempPlayer(i).EventMap.EventPages(x).eventID
                                    .pageID = TempPlayer(i).EventMap.EventPages(x).pageID
                                    .WaitingForResponse = 0
                                    ReDim .ListLeftOff(0 To Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(x).eventID).Pages(TempPlayer(i).EventMap.EventPages(x).pageID).CommandListCount)
                                End With
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
    
    'That is it for starting parallel processes :D now we just have to make the code that actually processes the events to their fullest
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).EventProcessingCount > 0 Then
restartloop:
                For x = 1 To TempPlayer(i).EventProcessingCount
                    With TempPlayer(i).EventProcessing(x)
                        If TempPlayer(i).EventProcessingCount = 0 Then Exit Sub
                        removeeventprocess = False
                        If .WaitingForResponse = 2 Then
                            If TempPlayer(i).InShop = 0 Then
                                .WaitingForResponse = 0
                            End If
                        End If
                        If .WaitingForResponse = 3 Then
                            If TempPlayer(i).InBank = False Then
                                .WaitingForResponse = 0
                            End If
                        End If
                        If .WaitingForResponse = 0 Then
                            If .ActionTimer <= GetTickCount Then
restartlist:
                                If .ListLeftOff(.CurList) > 0 Then
                                    .CurSlot = .ListLeftOff(.CurList) + 1
                                End If
                                If .CurList > Map(Player(i).Map).Events(.eventID).Pages(.pageID).CommandListCount Then
                                    'Get rid of this event, it is bad
                                    removeeventprocess = True
                                    GoTo endprocess
                                End If
                                If .CurSlot > Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).CommandCount Then
                                    If .CurList = 1 Then
                                        'Get rid of this event, it is bad
                                        removeeventprocess = True
                                        GoTo endprocess
                                    Else
                                        .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).ParentList
                                        .CurSlot = 1
                                        GoTo restartlist
                                    End If
                                End If
                                'If we are still here, then we are good to process shit :D
                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).index
                                    Case EventType.evAddText
                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                            Case 0
                                                PlayerMsg i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                            Case 1
                                                MapMsg GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                            Case 2
                                                GlobalMsg Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        End Select
                                    Case EventType.evShowText
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SEventChat
                                        Buffer.WriteLong .eventID
                                        Buffer.WriteLong .pageID
                                        Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1
                                        Buffer.WriteLong 0
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).CommandCount > .CurSlot Then
                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evShowText Or Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evShowChoices Then
                                                Buffer.WriteLong 1
                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evCondition Then
                                                Buffer.WriteLong 2
                                            Else
                                                Buffer.WriteLong 0
                                            End If
                                        Else
                                            Buffer.WriteLong 2
                                        End If
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                        .WaitingForResponse = 1
                                    Case EventType.evShowChoices
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SEventChat
                                        Buffer.WriteLong .eventID
                                        Buffer.WriteLong .pageID
                                        Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1
                                        If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text2)) > 0 Then
                                            w = 1
                                            If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text3)) > 0 Then
                                                w = 2
                                                If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text4)) > 0 Then
                                                    w = 3
                                                    If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text5)) > 0 Then
                                                        w = 4
                                                    End If
                                                End If
                                            End If
                                        End If
                                        Buffer.WriteLong w
                                        For v = 1 To w
                                            Select Case v
                                                Case 1
                                                    Buffer.WriteString Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text2)
                                                Case 2
                                                    Buffer.WriteString Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text3)
                                                Case 3
                                                    Buffer.WriteString Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text4)
                                                Case 4
                                                    Buffer.WriteString Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text5)
                                            End Select
                                        Next
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).CommandCount > .CurSlot Then
                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evShowText Or Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evShowChoices Then
                                                Buffer.WriteLong 1
                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot + 1).index = EventType.evCondition Then
                                                Buffer.WriteLong 2
                                            Else
                                                Buffer.WriteLong 0
                                            End If
                                        Else
                                            Buffer.WriteLong 2
                                        End If
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                        .WaitingForResponse = 1
                                    Case EventType.evPlayerVar
                                        Player(i).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                    Case EventType.evPlayerSwitch
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            Player(i).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = 1
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            Player(i).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = 0
                                        End If
                                    Case EventType.evSelfSwitch
                                        If Map(GetPlayerMap(i)).Events(.eventID).Global = 1 Then
                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 1
                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 0
                                            End If
                                        Else
                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 1
                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 0
                                            End If
                                        End If
                                    Case EventType.evCondition
                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Condition
                                            Case 0
                                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                    Case 0
                                                        If Player(i).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 1
                                                        If Player(i).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) >= Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 2
                                                        If Player(i).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) <= Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 3
                                                        If Player(i).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) > Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 4
                                                        If Player(i).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) < Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 5
                                                        If Player(i).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) <> Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                End Select
                                            Case 1
                                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                    Case 0
                                                        If Player(i).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = 1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 1
                                                        If Player(i).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = 0 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                End Select
                                            Case 2
                                                If HasItem(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) > 0 Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                            Case 3
                                                If Player(i).Class = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                            Case 4
                                                If HasSpell(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = True Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                            Case 5
                                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                    Case 0
                                                        If GetPlayerLevel(i) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 1
                                                        If GetPlayerLevel(i) >= Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 2
                                                        If GetPlayerLevel(i) <= Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 3
                                                        If GetPlayerLevel(i) > Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 4
                                                        If GetPlayerLevel(i) < Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 5
                                                        If GetPlayerLevel(i) <> Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                End Select
                                            Case 6
                                                If Map(GetPlayerMap(i)).Events(.eventID).Global = 1 Then
                                                    Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                        Case 0 'Self Switch is true
                                                            If Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 1 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                        Case 1  'self switch is false
                                                            If Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 0 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                    End Select
                                                Else
                                                    Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                        Case 0 'Self Switch is true
                                                            If TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 1 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                        Case 1  'self switch is false
                                                            If TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 0 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                    End Select
                                                End If
                                        End Select
                                        GoTo endprocess
                                    Case EventType.evExitProcess
                                        removeeventprocess = True
                                        GoTo endprocess
                                    Case EventType.evChangeItems
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            If FindItem(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) > 0 Then
                                                Call SetPlayerInvItemValue(i, FindItem(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3)
                                            End If
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            GiveInvItem i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, True
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 2 Then
                                            TakeInvItem i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                        End If
                                        SendInventory i
                                    Case EventType.evRestoreHP
                                        SetPlayerVital i, HP, GetPlayerMaxVital(i, HP)
                                        SendVital i, HP
                                    Case EventType.evRestoreMP
                                        SetPlayerVital i, MP, GetPlayerMaxVital(i, MP)
                                        SendVital i, MP
                                    Case EventType.evLevelUp
                                        SetPlayerLevel i, GetPlayerLevel(i) + 1
                                        SetPlayerExp i, 0
                                        SendEXP i
                                        SendPlayerData i
                                    Case EventType.evChangeLevel
                                        SetPlayerLevel i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SetPlayerExp i, 0
                                        SendEXP i
                                        SendPlayerData i
                                    Case EventType.evChangeSkills
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            If FindOpenSpellSlot(i) > 0 Then
                                                If HasSpell(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = False Then
                                                    SetPlayerSpell i, FindOpenSpellSlot(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                Else
                                                    'Error, already knows spell
                                                End If
                                            Else
                                                'Error, no room for spells
                                            End If
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            If HasSpell(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1) = True Then
                                                For p = 1 To MAX_PLAYER_SPELLS
                                                    If Player(i).Spell(p) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 Then
                                                        SetPlayerSpell i, p, 0
                                                    End If
                                                Next
                                            End If
                                        End If
                                        SendPlayerSpells i
                                    Case EventType.evChangeClass
                                        Player(i).Class = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SendPlayerData i
                                    Case EventType.evChangeSprite
                                        Player(i).Sprite = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SendPlayerData i
                                    Case EventType.evChangeSex
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 0 Then
                                            Player(i).Sex = SEX_MALE
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 1 Then
                                            Player(i).Sex = SEX_FEMALE
                                        End If
                                        SendPlayerData i
                                    Case EventType.evChangePK
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 0 Then
                                            Player(i).PK = NO
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 1 Then
                                            Player(i).PK = YES
                                        End If
                                        SendPlayerData i
                                    Case EventType.evWarpPlayer
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).data4 = 0 Then
                                            PlayerWarp i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                        Else
                                            Player(i).Dir = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).data4 - 1
                                            PlayerWarp i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                        End If
                                        
                                    Case EventType.evSetMoveRoute
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 <= Map(GetPlayerMap(i)).EventCount Then
                                            If Map(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Global = 1 Then
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveType = 2
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).IgnoreIfCannotMove = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).RepeatMoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteCount = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRouteCount
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRoute
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteStep = 0
                                            Else
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveType = 2
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).IgnoreIfCannotMove = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).RepeatMoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteCount = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRouteCount
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).MoveRoute
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteStep = 0
                                            End If
                                        End If
                                    Case EventType.evPlayAnimation
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, GetPlayerX(i), GetPlayerY(i), TARGET_TYPE_PLAYER, i
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            If Map(GetPlayerMap(i)).Events(.eventID).Global = 1 Then
                                                SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).x, Map(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).y
                                            Else
                                                SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).x, TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3).y
                                            End If
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 2 Then
                                            SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).data4
                                        End If
                                    Case EventType.evCustomScript
                                        'Runs Through Cases for a script
                                        Call CustomScript(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                    Case EventType.evPlayBGM
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SPlayBGM
                                        Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                    Case EventType.evFadeoutBGM
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SFadeoutBGM
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                    Case EventType.evPlaySound
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SPlaySound
                                        Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Text1
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                    Case EventType.evStopSound
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SStopSound
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                    Case EventType.evSetAccess
                                        Player(i).Access = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SendPlayerData i
                                    Case EventType.evOpenShop
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 > 0 Then ' shop exists?
                                            If Len(Trim$(Shop(Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1).Name)) > 0 Then ' name exists?
                                                SendOpenShop i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                TempPlayer(i).InShop = Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1 ' stops movement and the like
                                                .WaitingForResponse = 2
                                            End If
                                        End If
                                    Case EventType.evOpenBank
                                        SendBank i
                                        TempPlayer(i).InBank = True
                                        .WaitingForResponse = 3
                                    Case EventType.evGiveExp
                                        GivePlayerEXP i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.pageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                End Select
                                .CurSlot = .CurSlot + 1
                            End If
                        End If
endprocess:
                    End With
                    If removeeventprocess = True Then
                        TempPlayer(i).EventProcessingCount = TempPlayer(i).EventProcessingCount - 1
                        If TempPlayer(i).EventProcessingCount <= 0 Then
                            ReDim TempPlayer(i).EventProcessing(0)
                            GoTo restartloop:
                        Else
                            For z = x To TempPlayer(i).EventProcessingCount - 1
                                TempPlayer(i).EventProcessing(x) = TempPlayer(i).EventProcessing(x + 1)
                            Next
                            ReDim Preserve TempPlayer(i).EventProcessing(TempPlayer(i).EventProcessingCount)
                            GoTo restartloop
                        End If
                    End If
                Next
            End If
        End If
    Next
End Sub

Public Sub UpdateEventLogic()
Dim i As Long, x As Long, y As Long, z As Long, mapnum As Long, id As Long
Dim page As Long, Buffer As clsBuffer, spawnevent As Boolean, p As Long, rand As Long, isglobal As Boolean, actualmovespeed As Long, playerID As Long, WalkThrough As Long, eventID As Long, sendupdate As Boolean, removeeventprocess As Boolean, w As Long, v As Long
    'Check Removing and Adding of Events (Did switches change or something?)
    RemoveDeadEvents
    SpawnNewEvents
    ProcessEventMovement
    ProcessLocalEventMovement
    ProcessEventCommands
    
End Sub

Sub SendSwitchesAndVariables(index As Long, Optional everyone As Boolean = False)
Dim Buffer As clsBuffer, i As Long

Set Buffer = New clsBuffer
Buffer.WriteLong SSwitchesAndVariables

For i = 1 To MAX_SWITCHES
    Buffer.WriteString Switches(i)
Next

For i = 1 To MAX_VARIABLES
    Buffer.WriteString Variables(i)
Next

If everyone Then
    SendDataToAll Buffer.ToArray
Else
    SendDataTo index, Buffer.ToArray
End If

Set Buffer = Nothing
End Sub

Sub SendMapEventData(index As Long)
Dim Buffer As clsBuffer, i As Long, x As Long, y As Long, z As Long, mapnum As Long, w As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapEventData
    mapnum = GetPlayerMap(index)
    'Event Data
    Buffer.WriteLong Map(mapnum).EventCount
        
    If Map(mapnum).EventCount > 0 Then
        For i = 1 To Map(mapnum).EventCount
            With Map(mapnum).Events(i)
                Buffer.WriteString .Name
                Buffer.WriteLong .Global
                Buffer.WriteLong .x
                Buffer.WriteLong .y
                Buffer.WriteLong .PageCount
            End With
            If Map(mapnum).Events(i).PageCount > 0 Then
                For x = 1 To Map(mapnum).Events(i).PageCount
                    With Map(mapnum).Events(i).Pages(x)
                        Buffer.WriteLong .chkVariable
                        Buffer.WriteLong .VariableIndex
                        Buffer.WriteLong .VariableCondition
                        Buffer.WriteLong .VariableCompare
                            
                        Buffer.WriteLong .chkSwitch
                        Buffer.WriteLong .SwitchIndex
                        Buffer.WriteLong .SwitchCompare
                        
                        Buffer.WriteLong .chkHasItem
                        Buffer.WriteLong .HasItemIndex
                            
                        Buffer.WriteLong .chkSelfSwitch
                        Buffer.WriteLong .SelfSwitchIndex
                        Buffer.WriteLong .SelfSwitchCompare
                            
                        Buffer.WriteLong .GraphicType
                        Buffer.WriteLong .Graphic
                        Buffer.WriteLong .GraphicX
                        Buffer.WriteLong .GraphicY
                        Buffer.WriteLong .GraphicX2
                        Buffer.WriteLong .GraphicY2
                        
                        Buffer.WriteLong .MoveType
                        Buffer.WriteLong .MoveSpeed
                        Buffer.WriteLong .MoveFreq
                        Buffer.WriteLong .MoveRouteCount
                        
                        Buffer.WriteLong .IgnoreMoveRoute
                        Buffer.WriteLong .RepeatMoveRoute
                            
                        If .MoveRouteCount > 0 Then
                            For y = 1 To .MoveRouteCount
                                Buffer.WriteLong .MoveRoute(y).index
                                Buffer.WriteLong .MoveRoute(y).Data1
                                Buffer.WriteLong .MoveRoute(y).Data2
                                Buffer.WriteLong .MoveRoute(y).Data3
                                Buffer.WriteLong .MoveRoute(y).data4
                                Buffer.WriteLong .MoveRoute(y).data5
                                Buffer.WriteLong .MoveRoute(y).data6
                            Next
                        End If
                            
                        Buffer.WriteLong .WalkAnim
                        Buffer.WriteLong .DirFix
                        Buffer.WriteLong .WalkThrough
                        Buffer.WriteLong .ShowName
                        Buffer.WriteLong .Trigger
                        Buffer.WriteLong .CommandListCount
                        
                        Buffer.WriteLong .Position
                    End With
                        
                    If Map(mapnum).Events(i).Pages(x).CommandListCount > 0 Then
                        For y = 1 To Map(mapnum).Events(i).Pages(x).CommandListCount
                            Buffer.WriteLong Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount
                            Buffer.WriteLong Map(mapnum).Events(i).Pages(x).CommandList(y).ParentList
                            If Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                For z = 1 To Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount
                                    With Map(mapnum).Events(i).Pages(x).CommandList(y).Commands(z)
                                        Buffer.WriteLong .index
                                        Buffer.WriteString .Text1
                                        Buffer.WriteString .Text2
                                        Buffer.WriteString .Text3
                                        Buffer.WriteString .Text4
                                        Buffer.WriteString .Text5
                                        Buffer.WriteLong .Data1
                                        Buffer.WriteLong .Data2
                                        Buffer.WriteLong .Data3
                                        Buffer.WriteLong .data4
                                        Buffer.WriteLong .data5
                                        Buffer.WriteLong .data6
                                        Buffer.WriteLong .ConditionalBranch.CommandList
                                        Buffer.WriteLong .ConditionalBranch.Condition
                                        Buffer.WriteLong .ConditionalBranch.Data1
                                        Buffer.WriteLong .ConditionalBranch.Data2
                                        Buffer.WriteLong .ConditionalBranch.Data3
                                        Buffer.WriteLong .ConditionalBranch.ElseCommandList
                                        Buffer.WriteLong .MoveRouteCount
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                Buffer.WriteLong .MoveRoute(w).index
                                                Buffer.WriteLong .MoveRoute(w).Data1
                                                Buffer.WriteLong .MoveRoute(w).Data2
                                                Buffer.WriteLong .MoveRoute(w).Data3
                                                Buffer.WriteLong .MoveRoute(w).data4
                                                Buffer.WriteLong .MoveRoute(w).data5
                                                Buffer.WriteLong .MoveRoute(w).data6
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    'End Event Data
    SendDataTo index, Buffer.ToArray
    Set Buffer = Nothing
End Sub
