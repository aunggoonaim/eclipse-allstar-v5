Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
Dim I As Long
    If Index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(Index)
                Case 1 ' Elf
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 10 + 200
                Case 2 ' Man
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 10 + 260
                Case 3 ' Orc
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 10 + 280
                Case 4 ' Dwarf
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 10 + 300
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 15 + 150
            End Select
                        For I = 1 To 10
                If TempPlayer(Index).Buffs(I) = BUFF_ADD_HP Then
                    GetPlayerMaxVital = GetPlayerMaxVital + TempPlayer(Index).BuffValue(I)

                End If
                If TempPlayer(Index).Buffs(I) = BUFF_SUB_HP Then
                    GetPlayerMaxVital = GetPlayerMaxVital - TempPlayer(Index).BuffValue(I)
                End If
            Next

        Case MP
            Select Case GetPlayerClass(Index)
                Case 1 ' Elf
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 10 + 250
                Case 2 ' Man
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 10 + 240
                Case 3 ' Orc
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 10 + 160
                Case 4 ' Dwarf
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 10 + 190
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 5 + 25
            End Select
            For I = 1 To 10
                If TempPlayer(Index).Buffs(I) = BUFF_ADD_MP Then
                    GetPlayerMaxVital = GetPlayerMaxVital + TempPlayer(Index).BuffValue(I)
                End If
                If TempPlayer(Index).Buffs(I) = BUFF_SUB_MP Then
                    GetPlayerMaxVital = GetPlayerMaxVital - TempPlayer(Index).BuffValue(I)
                End If
            Next



    End Select
End Function

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim I As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            I = (GetPlayerStat(Index, Stats.Willpower) * 0.8) + 6
        Case MP
            I = (GetPlayerStat(Index, Stats.Willpower) / 4) + 12.5
    End Select

    If I < 2 Then I = 2
    GetPlayerVitalRegen = I
End Function

Function GetPlayerDamage(ByVal Index As Long) As Long
Dim I As Long
    Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(Index, Weapon)
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) * Item(weaponNum).Data2 + (GetPlayerLevel(Index) / 5)
    Else
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) + (GetPlayerLevel(Index) / 5)
    End If
For I = 1 To 10
        If TempPlayer(Index).Buffs(I) = BUFF_ADD_ATK Then
            GetPlayerDamage = GetPlayerDamage + TempPlayer(Index).BuffValue(I)
        End If
        If TempPlayer(Index).Buffs(I) = BUFF_SUB_ATK Then
            GetPlayerDamage = GetPlayerDamage - TempPlayer(Index).BuffValue(I)
        End If
    Next

End Function

Function GetPlayerDef(ByVal Index As Long) As Long
     Dim I As Long
    Dim DefNum As Long
    Dim Def As Long
    
    GetPlayerDef = 0
    Def = 0
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    
    If GetPlayerEquipment(Index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(Index, Armor)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(Index, Helmet) > 0 Then
        DefNum = GetPlayerEquipment(Index, Helmet)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(Index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(Index, Shield)
        Def = Def + Item(DefNum).Data2
    End If
    
   If Not GetPlayerEquipment(Index, Armor) > 0 And Not GetPlayerEquipment(Index, Helmet) > 0 And Not GetPlayerEquipment(Index, Shield) > 0 Then
        GetPlayerDef = 0.085 * GetPlayerStat(Index, Endurance) + (GetPlayerLevel(Index) / 5)
    Else
        GetPlayerDef = 0.085 * GetPlayerStat(Index, Endurance) * Def + (GetPlayerLevel(Index) / 5)
    End If
    For I = 1 To 10
        If TempPlayer(Index).Buffs(I) = BUFF_ADD_DEF Then
            GetPlayerDef = GetPlayerDef + TempPlayer(Index).BuffValue(I)
        End If
        If TempPlayer(Index).Buffs(I) = BUFF_SUB_DEF Then
            GetPlayerDef = GetPlayerDef - TempPlayer(Index).BuffValue(I)
        End If
    Next


End Function

Function GetNpcMaxVital(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim x As Long

    ' Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = NPC(npcNum).HP
        Case MP
            GetNpcMaxVital = 30 + (NPC(npcNum).stat(Intelligence) * 10) + 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim I As Long

    'Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            I = (NPC(npcNum).stat(Stats.Willpower) * 0.8) + 6
        Case MP
            I = (NPC(npcNum).stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = I

End Function

Function GetNpcDamage(ByVal npcNum As Long) As Long
    GetNpcDamage = 0.085 * 5 * NPC(npcNum).stat(Stats.Strength) * NPC(npcNum).Damage + (NPC(npcNum).Level / 5)
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(Index, Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False

    rate = GetPlayerStat(Index, Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerParry = False

    rate = GetPlayerStat(Index, Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcBlock = False

    rate = 0
    ' TODO : make it based on shield lol
End Function

Public Function CanNpcCrit(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = NPC(npcNum).stat(Stats.Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = NPC(npcNum).stat(Stats.Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = NPC(npcNum).stat(Stats.Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal mapNpcNum As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, mapNpcNum) Then
    
        mapNum = GetPlayerMap(Index)
        npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(npcNum) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(npcNum) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(mapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (NPC(npcNum).stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, mapNpcNum, Damage)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal mapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapNum As Long
    Dim npcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).NPC(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapNum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
    
        ' exit out early
        If IsSpell Then
             If npcNum > 0 Then
                If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    Dim petowner As Long
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If
        

        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If npcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    NpcX = MapNpc(mapNum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapNum).NPC(mapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(mapNum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapNum).NPC(mapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(mapNum).NPC(mapNpcNum).x + 1
                    NpcY = MapNpc(mapNum).NPC(mapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(mapNum).NPC(mapNpcNum).x - 1
                    NpcY = MapNpc(mapNum).NPC(mapNpcNum).y
            End Select

            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then

TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
TempPlayer(Attacker).target = mapNpcNum
SendTarget Attacker

If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        CanPlayerAttackNpc = True
                        End If

                        If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                            Call CheckTasks(Attacker, QUEST_TYPE_GOTALK, npcNum)
                            Call CheckTasks(Attacker, QUEST_TYPE_GOGIVE, npcNum)
                            Call CheckTasks(Attacker, QUEST_TYPE_GOGET, npcNum)
                            If NPC(npcNum).Quest = YES Then
                                If CanStartQuest(Attacker, NPC(npcNum).QuestNum) Then
                                    'if can start show the request message (chat1)
                                    QuestMessage Attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).Chat(1)), NPC(npcNum).QuestNum
                                    Exit Function
                                End If
                                If QuestInProgress(Attacker, NPC(npcNum).QuestNum) Then
                                    'if the quest is in progress show the meanwhile message (chat2)
                                    PlayerMsg Attacker, Trim$(NPC(npcNum).Name) + ": " + Trim$(Quest(NPC(npcNum).QuestNum).Chat(2)), BrightGreen
                                    'QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).Chat(2)), 0
                                    Exit Function
                                End If
                            End If
                                                      If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                            If NPC(npcNum).Convo = True Then
                                InitChat Attacker, mapNum, mapNpcNum
                                Exit Function
                            End If
                        End If
                            If NPC(npcNum).Quest = NO Then
                                If NPC(npcNum).Convo = False Then
                                If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                    PlayerMsg Attacker, Trim$(NPC(npcNum).Name) & ": " & Trim$(NPC(npcNum).AttackSay), White
                                Else
                                    InitChat Attacker, mapNum, mapNpcNum
                                End If
                                Exit Function
                            End If
                        End If
                        End If
                    End If
                End If
            End If
        End If
End Function
Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n As Long
    Dim I As Long
    Dim STR As Long
    Dim Def As Long
    Dim mapNum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapNum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
    ' projectiles
    If npcNum < 1 Then Exit Sub
    Name = Trim$(NPC(npcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Attacker, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellnum = 0 Then Call SendAnimation(mapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y)
            End If
        End If

        ' Calculate exp to give attacker
        exp = NPC(npcNum).exp

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
        Else
            ' no party - keep exp for self
            GivePlayerEXP Attacker, exp
        End If
        
        'Drop the goods if they get it
        n = Int(Rnd * NPC(npcNum).DropChance) + 1

        If n = 1 Then
            Call SpawnItem(NPC(npcNum).DropItem, NPC(npcNum).DropItemValue, mapNum, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y)
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapNum).NPC(mapNpcNum).Num = 0
        MapNpc(mapNum).NPC(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) = 0
        
        'Checks if NPC was a pet
        If MapNpc(mapNum).NPC(mapNpcNum).IsPet = YES Then
            Call PetDisband(MapNpc(mapNum).NPC(mapNpcNum).PetData.Owner, mapNum) 'The pet was killed
        End If
        
        ' clear DoTs and HoTs
        For I = 1 To MAX_DOTS
            With MapNpc(mapNum).NPC(mapNpcNum).DoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapNum).NPC(mapNpcNum).HoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        Call CheckTasks(Attacker, QUEST_TYPE_GOSLAY, npcNum)
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong mapNpcNum
        SendDataToMap mapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And IsConnected(I) Then
                If Player(I).Map = mapNum Then
                    If TempPlayer(I).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(I).target = mapNpcNum Then
                            TempPlayer(I).target = 0
                            TempPlayer(I).targetType = TARGET_TYPE_NONE
                            SendTarget I
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg mapNum, "-" & Damage, BrightRed, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Attacker, MapNpc(mapNum).NPC(mapNpcNum).x, MapNpc(mapNum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellnum = 0 Then Call SendAnimation(mapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(mapNum).NPC(mapNpcNum).targetType = 1 ' player
        MapNpc(mapNum).NPC(mapNpcNum).target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(mapNum).NPC(I).Num = MapNpc(mapNum).NPC(mapNpcNum).Num Then
                    MapNpc(mapNum).NPC(I).target = Attacker
                    MapNpc(mapNum).NPC(I).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapNum).NPC(mapNpcNum).stopRegen = True
        MapNpc(mapNum).NPC(mapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunNPC mapNpcNum, mapNum, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Npc mapNum, mapNpcNum, spellnum, Attacker
            End If
        End If
        
        SendMapNpcVitals mapNum, mapNpcNum
    End If

    If spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = GetTickCount
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal Index As Long)
Dim mapNum As Long, npcNum As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(mapNpcNum, Index) Then
        mapNum = GetPlayerMap(Index)
        npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (Player(Index).x * 32), (Player(Index).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(Index) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (Player(Index).x * 32), (Player(Index).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(Index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(Index, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (MapNpc(mapNum).NPC(mapNpcNum).x * 32), (MapNpc(mapNum).NPC(mapNpcNum).y * 32)
        End If
        
        Damage = Damage - GetPlayerDef(Index)

        If Damage > 0 Then
            Call NpcAttackPlayer(mapNpcNum, Index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim mapNum As Long
    Dim npcNum As Long
    Dim petowner As Long

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).NPC(mapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    'check if the NPC attacking us is actually our pet.
'We don't want a rebellion on our hands now do we?
        

    mapNum = GetPlayerMap(Index)
    npcNum = MapNpc(mapNum).NPC(mapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(mapNum).NPC(mapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(mapNum).NPC(mapNpcNum).AttackTimer = GetTickCount
   

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If npcNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(mapNum).NPC(mapNpcNum).y) And (GetPlayerX(Index) = MapNpc(mapNum).NPC(mapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(mapNum).NPC(mapNpcNum).y) And (GetPlayerX(Index) = MapNpc(mapNum).NPC(mapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(mapNum).NPC(mapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(mapNum).NPC(mapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(mapNum).NPC(mapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(mapNum).NPC(mapNpcNum).x) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal mapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim exp As Long
    Dim mapNum As Long
    Dim I As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).NPC(mapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    mapNum = GetPlayerMap(Victim)
    Name = Trim$(NPC(MapNpc(mapNum).NPC(mapNpcNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong mapNpcNum
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(mapNum).NPC(mapNpcNum).stopRegen = True
    MapNpc(mapNum).NPC(mapNpcNum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(mapNum).NPC(mapNpcNum).Num
        
        ' kill player
        KillPlayer Victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(mapNum).NPC(mapNpcNum).target = 0
        MapNpc(mapNum).NPC(mapNpcNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        Call SendAnimation(mapNum, NPC(MapNpc(GetPlayerMap(Victim)).NPC(mapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(mapNum).NPC(mapNpcNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
    End If

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, Victim) Then
    
        mapNum = GetPlayerMap(Attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(Victim) Then
            SendActionMsg mapNum, "Dodge!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(Victim) Then
            SendActionMsg mapNum, "Parry!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(Victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(Victim, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        End If

        Damage = Damage - GetPlayerDef(Victim)

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, Victim, Damage)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

' projectiles
Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean

    If Not IsSpell And Not IsProjectile Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function

    If Not IsSpell And Not IsProjectile Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If
TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
TempPlayer(Attacker).target = Victim
SendTarget Attacker
    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim I As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        exp = (GetPlayerExp(Victim) \ 10)

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If

        If exp = 0 Then
            Call PlayerMsg(Victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - exp)
            SendEXP Victim
            Call PlayerMsg(Victim, "You lost " & exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And IsConnected(I) Then
                If Player(I).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(I).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(I).target = Victim Then
                            TempPlayer(I).target = 0
                            TempPlayer(I).targetType = TARGET_TYPE_NONE
                            SendTarget I
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
        
        Call CheckTasks(Attacker, QUEST_TYPE_GOKILL, Victim)
        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
        
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunPlayer Victim, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Player Victim, spellnum, Attacker
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal Index As Long, ByVal spellslot As Long)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapNum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellnum = GetPlayerSpell(Index, spellslot)
    mapNum = GetPlayerMap(Index)
    
    If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(Index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(Index).targetType
    target = TempPlayer(Index).target
    Range = Spell(spellnum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg Index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(mapNum).NPC(target).x, MapNpc(mapNum).NPC(target).y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(Index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapNum, Spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg mapNum, "Casting " & Trim$(Spell(spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        TempPlayer(Index).spellBuffer.Spell = spellslot
        TempPlayer(Index).spellBuffer.Timer = GetTickCount
        TempPlayer(Index).spellBuffer.target = TempPlayer(Index).target
        TempPlayer(Index).spellBuffer.tType = TempPlayer(Index).targetType
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim Dur As Long
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapNum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim I As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(Index, spellslot)
    mapNum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, spellnum) Then Exit Sub

    MPCost = Spell(spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    ' set the vital
    Vital = Spell(spellnum).Vital
   If Spell(spellnum).Type <> SPELL_TYPE_BUFF Then
        Vital = Spell(spellnum).Vital
        Vital = Round((Vital * 0.6)) * Round((Player(Index).Level * 1.14)) * Round((Stats.Intelligence + (Stats.Willpower / 2)))
    
        If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
            Vital = Vital + Round((GetPlayerStat(Index, Stats.Willpower) * 1.2))
        End If
    
        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEHP Then
            Vital = Vital + Round((GetPlayerStat(Index, Stats.Intelligence) * 1.2))
        End If
    End If
    
    If Spell(spellnum).Type = SPELL_TYPE_BUFF Then
        If Round(GetPlayerStat(Index, Stats.Willpower) / 5) > 1 Then
            Dur = Spell(spellnum).Duration * Round(GetPlayerStat(Index, Stats.Willpower) / 5)
        Else
            Dur = Spell(spellnum).Duration
        End If
    End If

    Range = Spell(spellnum).Range
    
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellnum).Type
                            Case SPELL_TYPE_BUFF
                        Call ApplyBuff(Index, Spell(spellnum).BuffType, Dur, Spell(spellnum).Vital)
                        SendAnimation GetPlayerMap(Index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                        ' send the sound
                        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, spellnum
                        DidCast = True

                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, Index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, Index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation mapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    PlayerWarp Index, Spell(spellnum).Map, Spell(spellnum).x, Spell(spellnum).y
                    SendAnimation GetPlayerMap(Index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(Index)
                y = GetPlayerY(Index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = MapNpc(mapNum).NPC(target).x
                    y = MapNpc(mapNum).NPC(target).y
                End If
                
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    SendClearSpellBuffer Index
                End If
            End If
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If I <> Index Then
                                If GetPlayerMap(I) = GetPlayerMap(Index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(I), GetPlayerY(I)) Then
                                        If CanPlayerAttackPlayer(Index, I, True) Then
                                            SendAnimation mapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, I
                                            PlayerAttackPlayer Index, I, Vital, spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(mapNum).NPC(I).Num > 0 Then
                            If MapNpc(mapNum).NPC(I).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapNum).NPC(I).x, MapNpc(mapNum).NPC(I).y) Then
                                    If CanPlayerAttackNpc(Index, I, True) Then
                                        SendAnimation mapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, I
                                        PlayerAttackNpc Index, I, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                    
                    DidCast = True
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(Index) Then
                                If isInRange(AoE, x, y, GetPlayerX(I), GetPlayerY(I)) Then
                                    SpellPlayer_Effect VitalType, increment, I, Vital, spellnum
                                End If
                            End If
                        End If
                    Next
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(mapNum).NPC(I).Num > 0 Then
                            If MapNpc(mapNum).NPC(I).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapNum).NPC(I).x, MapNpc(mapNum).NPC(I).y) Then
                                    SpellNpc_Effect VitalType, increment, I, Vital, spellnum, mapNum
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
            
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(target)
                y = GetPlayerY(target)
            Else
                x = MapNpc(mapNum).NPC(target).x
                y = MapNpc(mapNum).NPC(target).y
            End If
                
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
                Exit Sub
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer Index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc Index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    End If
                    
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    End If
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(Index, target, True) Then
                                SpellPlayer_Effect VitalType, increment, target, Vital, spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, target, Vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(Index, target, True) Then
                                SpellNpc_Effect VitalType, increment, target, Vital, spellnum, mapNum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, target, Vital, spellnum, mapNum
                        End If
                    End If
                    Case SPELL_TYPE_BUFF
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).BuffType <= BUFF_ADD_DEF And Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_NONE Or Spell(spellnum).BuffType > BUFF_NONE And Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Then
                            Call ApplyBuff(target, Spell(spellnum).BuffType, Dur, Spell(spellnum).Vital)
                            SendAnimation GetPlayerMap(Index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                            ' send the sound
                            SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, spellnum
                            DidCast = True
                        Else
                            PlayerMsg Index, "You can not debuff another player in a safe zone!", BrightRed
                        End If
                    End If


            End Select
    End Select
    
    If DidCast Then
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPCost)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
        
        TempPlayer(Index).SpellCD(spellslot) = GetTickCount + (Spell(spellnum).CDTime * 1000)
        Call SendCooldown(Index, spellslot)
        SendActionMsg mapNum, Trim$(Spell(spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal spellnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(Index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        
        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, spellnum
        
        If increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
            If Spell(spellnum).Duration > 0 Then
                AddHoT_Player Index, spellnum
            End If
        ElseIf Not increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
        End If
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal spellnum As Long, ByVal mapNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation mapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
        SendActionMsg mapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(mapNum).NPC(Index).x * 32, MapNpc(mapNum).NPC(Index).y * 32
        
        ' send the sound
        SendMapSound Index, MapNpc(mapNum).NPC(Index).x, MapNpc(mapNum).NPC(Index).y, SoundEntity.seSpell, spellnum
        
        If increment Then
            MapNpc(mapNum).NPC(Index).Vital(Vital) = MapNpc(mapNum).NPC(Index).Vital(Vital) + Damage
            If Spell(spellnum).Duration > 0 Then
                AddHoT_Npc mapNum, Index, spellnum
            End If
        ElseIf Not increment Then
            MapNpc(mapNum).NPC(Index).Vital(Vital) = MapNpc(mapNum).NPC(Index).Vital(Vital) - Damage
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal spellnum As Long, ByVal Caster As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(I)
            If .Spell = spellnum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal Index As Long, ByVal spellnum As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(I)
            If .Spell = spellnum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal mapNum As Long, ByVal Index As Long, ByVal spellnum As Long, ByVal Caster As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With MapNpc(mapNum).NPC(Index).DoT(I)
            If .Spell = spellnum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal mapNum As Long, ByVal Index As Long, ByVal spellnum As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With MapNpc(mapNum).NPC(Index).HoT(I)
            If .Spell = spellnum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal Index As Long, ByVal dotNum As Long)
    With TempPlayer(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, Index, True) Then
                    PlayerAttackPlayer .Caster, Index, Spell(.Spell).Vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal Index As Long, ByVal hotNum As Long)
    With TempPlayer(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg Player(Index).Map, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Player(Index).x * 32, Player(Index).y * 32
                Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + Spell(.Spell).Vital
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal mapNum As Long, ByVal Index As Long, ByVal dotNum As Long)
    With MapNpc(mapNum).NPC(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, Index, True) Then
                    PlayerAttackNpc .Caster, Index, Spell(.Spell).Vital, , True
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal mapNum As Long, ByVal Index As Long, ByVal hotNum As Long)
    With MapNpc(mapNum).NPC(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg mapNum, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(mapNum).NPC(Index).x * 32, MapNpc(mapNum).NPC(Index).y * 32
                MapNpc(mapNum).NPC(Index).Vital(Vitals.HP) = MapNpc(mapNum).NPC(Index).Vital(Vitals.HP) + Spell(.Spell).Vital
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal Index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Spell(spellnum).StunDuration
        TempPlayer(Index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned Index
        ' tell him he's stunned
        PlayerMsg Index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal mapNum As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(mapNum).NPC(Index).StunDuration = Spell(spellnum).StunDuration
        MapNpc(mapNum).NPC(Index).StunTimer = GetTickCount
    End If
End Sub

Function CanNpcAttackNpc(ByVal mapNum As Long, ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    Dim aNpcNum As Long
    Dim vNpcNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long
    Dim petowner As Long
    
    CanNpcAttackNpc = False

    ' Check for subscript out of range
    If Attacker <= 0 Or Attacker > MAX_MAP_NPCS Then
        Exit Function
    End If
    
    If Victim <= 0 Or Victim > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(mapNum).NPC(Attacker).Num <= 0 Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(mapNum).NPC(Victim).Num <= 0 Then
        Exit Function
    End If

    aNpcNum = MapNpc(mapNum).NPC(Attacker).Num
    vNpcNum = MapNpc(mapNum).NPC(Victim).Num
    
    If aNpcNum <= 0 Then Exit Function
    If vNpcNum <= 0 Then Exit Function
    
    
    ' Make sure the npcs arent already dead
    If MapNpc(mapNum).NPC(Attacker).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapNum).NPC(Victim).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(mapNum).NPC(Attacker).AttackTimer + 1000 Then
        Exit Function
    End If
    
    MapNpc(mapNum).NPC(Attacker).AttackTimer = GetTickCount
    
    AttackerX = MapNpc(mapNum).NPC(Attacker).x
    AttackerY = MapNpc(mapNum).NPC(Attacker).y
    VictimX = MapNpc(mapNum).NPC(Victim).x
    VictimY = MapNpc(mapNum).NPC(Victim).y

    ' Check if at same coordinates
    If (VictimY + 1 = AttackerY) And (VictimX = AttackerX) Then
        CanNpcAttackNpc = True
    Else

        If (VictimY - 1 = AttackerY) And (VictimX = AttackerX) Then
            CanNpcAttackNpc = True
        Else

            If (VictimY = AttackerY) And (VictimX + 1 = AttackerX) Then
                CanNpcAttackNpc = True
            Else

                If (VictimY = AttackerY) And (VictimX - 1 = AttackerX) Then
                    CanNpcAttackNpc = True
                End If
            End If
        End If
    End If

End Function

Sub NpcAttackNpc(ByVal mapNum As Long, ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    Dim aNpcNum As Long
    Dim vNpcNum As Long
    Dim n As Long
    Dim petowner As Long
    
    If Attacker <= 0 Or Attacker > MAX_MAP_NPCS Then Exit Sub
    If Victim <= 0 Or Victim > MAX_MAP_NPCS Then Exit Sub
    
    If Damage <= 0 Then Exit Sub
    
    aNpcNum = MapNpc(mapNum).NPC(Attacker).Num
    vNpcNum = MapNpc(mapNum).NPC(Victim).Num
    
    If aNpcNum <= 0 Then Exit Sub
    If vNpcNum <= 0 Then Exit Sub
    
    'set the victim's target to the pet attacking it
    MapNpc(mapNum).NPC(Victim).targetType = 2 'Npc
    MapNpc(mapNum).NPC(Victim).target = Attacker
    
    ' Send this packet so they can see the person attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong Attacker
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing

    If Damage >= MapNpc(mapNum).NPC(Victim).Vital(Vitals.HP) Then
        SendActionMsg mapNum, "-" & Damage, BrightRed, 1, (MapNpc(mapNum).NPC(Victim).x * 32), (MapNpc(mapNum).NPC(Victim).y * 32)
        SendBlood mapNum, MapNpc(mapNum).NPC(Victim).x, MapNpc(mapNum).NPC(Victim).y
        
        ' npc is dead.
        'Call GlobalMsg(CheckGrammar(Trim$(Npc(vNpcNum).Name), 1) & " has been killed by " & CheckGrammar(Trim$(Npc(aNpcNum).Name)) & "!", BrightRed)

        ' Set NPC target to 0
        MapNpc(mapNum).NPC(Attacker).target = 0
        MapNpc(mapNum).NPC(Attacker).targetType = 0
        'reset the targetter for the player
        
        If MapNpc(mapNum).NPC(Attacker).IsPet = YES Then
            TempPlayer(MapNpc(mapNum).NPC(Attacker).PetData.Owner).target = 0
            TempPlayer(MapNpc(mapNum).NPC(Attacker).PetData.Owner).targetType = TARGET_TYPE_NONE
            
            petowner = MapNpc(mapNum).NPC(Attacker).PetData.Owner
            
            SendTarget petowner
            
            'Give the player the pet owner some experience from the kill
            Call SetPlayerExp(petowner, GetPlayerExp(petowner) + NPC(MapNpc(mapNum).NPC(Victim).Num).exp)
            CheckPlayerLevelUp petowner
            SendActionMsg mapNum, "+" & NPC(MapNpc(mapNum).NPC(Victim).Num).exp & "Exp", White, 1, GetPlayerX(petowner) * 32, GetPlayerY(petowner) * 32
            SendEXP petowner
                      
        ElseIf MapNpc(mapNum).NPC(Victim).IsPet = YES Then
            'Get the pet owners' index
            petowner = MapNpc(mapNum).NPC(Victim).PetData.Owner
            'Set the NPC's target on the owner now
            MapNpc(mapNum).NPC(Attacker).targetType = 1 'player
            MapNpc(mapNum).NPC(Attacker).target = petowner
            'Disband the pet
            PetDisband petowner, GetPlayerMap(petowner)
        End If
               
        ' Drop the goods if they get it
        'For n = 1 To MAX_NPC_DROPS
        If NPC(vNpcNum).DropItem <> 0 Then
            If Rnd <= NPC(vNpcNum).DropChance Then
                Call SpawnItem(NPC(vNpcNum).DropItem, NPC(vNpcNum).DropItemValue, mapNum, MapNpc(mapNum).NPC(Victim).x, MapNpc(mapNum).NPC(Victim).y)
            End If
        End If
        'Next
        
        
        ' Reset victim's stuff so it dies in loop
        MapNpc(mapNum).NPC(Victim).Num = 0
        MapNpc(mapNum).NPC(Victim).SpawnWait = GetTickCount
        MapNpc(mapNum).NPC(Victim).Vital(Vitals.HP) = 0
               
        ' send npc death packet to map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong Victim
        SendDataToMap mapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        If petowner > 0 Then
            PetFollowOwner petowner
        End If
    Else
        ' npc not dead, just do the damage
        MapNpc(mapNum).NPC(Victim).Vital(Vitals.HP) = MapNpc(mapNum).NPC(Victim).Vital(Vitals.HP) - Damage
       
        ' Say damage
        SendActionMsg mapNum, "-" & Damage, BrightRed, 1, (MapNpc(mapNum).NPC(Victim).x * 32), (MapNpc(mapNum).NPC(Victim).y * 32)
        SendBlood mapNum, MapNpc(mapNum).NPC(Victim).x, MapNpc(mapNum).NPC(Victim).y
    End If
    
    'Send both Npc's Vitals to the client
    SendMapNpcVitals mapNum, Attacker
    SendMapNpcVitals mapNum, Victim

End Sub
