Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
Dim i As Long
Dim Per1a As Integer
Dim Per2a As Integer
Dim Per1s As Integer
Dim Per2s As Integer


    Per1a = 0
    Per2a = 0
    Per1s = 0
    Per2s = 0

    If index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(index)
                Case 1 ' มนุษย์
                    GetPlayerMaxVital = ((100 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, Strength) * 7) + 80) * 2
                Case 2 ' เอลฟ์
                    GetPlayerMaxVital = ((80 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, Strength) * 5) + 65) * 2
                Case 3 ' การ์เดี้ยน
                    GetPlayerMaxVital = ((300 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, Strength) * 6) + 25) * 2
                Case 4 ' เบอเซิร์ก
                    GetPlayerMaxVital = ((500 * (GetPlayerLevel(index) / 2)) + (GetPlayerStat(index, Strength) * 8) + 10) * 2
                Case 5 ' พาลาดิน
                    GetPlayerMaxVital = ((350 * (GetPlayerLevel(index) / 2)) + (GetPlayerStat(index, Strength) * 6) + 25) * 2
                Case 6 ' วิซาร์ด
                    GetPlayerMaxVital = ((200 * (GetPlayerLevel(index) / 2)) + (GetPlayerStat(index, Strength) * 6) + 20) * 2
                Case 7 ' ซามูไร
                    GetPlayerMaxVital = ((180 * (GetPlayerLevel(index) / 2)) + (GetPlayerStat(index, Strength) * 9) + 20) * 2
                Case 8 ' ฮันเตอร์
                    GetPlayerMaxVital = ((190 * (GetPlayerLevel(index) / 2)) + (GetPlayerStat(index, Strength) * 6) + 25) * 2
                Case 9 ' สไนเปอร์
                    GetPlayerMaxVital = ((180 * (GetPlayerLevel(index) / 2)) + (GetPlayerStat(index, Strength) * 5) + 20) * 2
                Case 10 ' แอสแซสซิน
                    GetPlayerMaxVital = ((220 * (GetPlayerLevel(index) / 2)) + (GetPlayerStat(index, Strength) * 7) + 25) * 2
                Case 11 ' ดาร์คลอร์ด
                    GetPlayerMaxVital = ((150 * (GetPlayerLevel(index) / 2)) + (GetPlayerStat(index, Strength) * 7) + 20) * 2
                Case Else ' คลาสอื่นที่นอกเหนือจากนี้
                    GetPlayerMaxVital = ((GetPlayerLevel(index) * 10)) + (GetPlayerStat(index, Strength) * 8) + 15
            End Select
            
            ' เพิ่ม HP ตาม สิ่งของที่สวมใส่
            
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).Add1 > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).Per1 > 0 Then
                            Per1a = Per1a + Item(GetPlayerEquipment(index, Weapon)).HPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital + Item(GetPlayerEquipment(index, Weapon)).HPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Armor) > 0 Then
                If Item(GetPlayerEquipment(index, Armor)).Add1 > 0 Then
                    If Item(GetPlayerEquipment(index, Armor)).Per1 > 0 Then
                            Per1a = Per1a + Item(GetPlayerEquipment(index, Armor)).HPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital + Item(GetPlayerEquipment(index, Armor)).HPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Helmet) > 0 Then
                If Item(GetPlayerEquipment(index, Helmet)).Add1 > 0 Then
                    If Item(GetPlayerEquipment(index, Helmet)).Per1 > 0 Then
                            Per1a = Per1a + Item(GetPlayerEquipment(index, Helmet)).HPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital + Item(GetPlayerEquipment(index, Helmet)).HPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Shield) > 0 Then
                If Item(GetPlayerEquipment(index, Shield)).Add1 > 0 Then
                    If Item(GetPlayerEquipment(index, Shield)).Per1 > 0 Then
                            Per1a = Per1a + Item(GetPlayerEquipment(index, Shield)).HPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital + Item(GetPlayerEquipment(index, Shield)).HPCase
                    End If
                End If
            End If
            
            ' ลด HP ตามสิ่งของที่สวมใส่
            
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).Sub1 > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).Per1 > 0 Then
                            Per1s = Per1s + Item(GetPlayerEquipment(index, Weapon)).HPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital - Item(GetPlayerEquipment(index, Weapon)).HPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Armor) > 0 Then
                If Item(GetPlayerEquipment(index, Armor)).Sub1 > 0 Then
                    If Item(GetPlayerEquipment(index, Armor)).Per1 > 0 Then
                            Per1s = Per1s + Item(GetPlayerEquipment(index, Armor)).HPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital - Item(GetPlayerEquipment(index, Armor)).HPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Helmet) > 0 Then
                If Item(GetPlayerEquipment(index, Helmet)).Sub1 > 0 Then
                    If Item(GetPlayerEquipment(index, Helmet)).Per1 > 0 Then
                            Per1s = Per1s + Item(GetPlayerEquipment(index, Helmet)).HPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital - Item(GetPlayerEquipment(index, Helmet)).HPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Shield) > 0 Then
                If Item(GetPlayerEquipment(index, Shield)).Sub1 > 0 Then
                    If Item(GetPlayerEquipment(index, Shield)).Per1 > 0 Then
                            Per1s = Per1s + Item(GetPlayerEquipment(index, Shield)).HPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital - Item(GetPlayerEquipment(index, Shield)).HPCase
                    End If
                End If
            End If
            
            If Per1a > 0 Then
                GetPlayerMaxVital = GetPlayerMaxVital + (GetPlayerMaxVital * (Per1a / 100))
            End If
            
            If Per1s > 0 Then
                GetPlayerMaxVital = (GetPlayerMaxVital * ((100 - Per1s) / 100))
            End If
                        
            If GetPlayerMaxVital < 0 Then
                GetPlayerMaxVital = 1
            End If
            
        Case MP
            Select Case GetPlayerClass(index)
                Case 1 ' มนุษย์
                    GetPlayerMaxVital = (50 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 3) + 5
                Case 2 ' เอลฟ์
                    GetPlayerMaxVital = (100 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 7) + 12
                Case 3 ' การ์เดี้ยน
                    GetPlayerMaxVital = (65 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 3) + 3
                Case 4 ' เบอเซิร์ก
                    GetPlayerMaxVital = (35 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 2) + 1
                Case 5 ' พาลาดิน
                    GetPlayerMaxVital = (150 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 3) + 5
                Case 6 ' วิซาร์ด
                    GetPlayerMaxVital = (120 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 8) + 16
                Case 7 ' ซามูไร
                    GetPlayerMaxVital = (100 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 7) + 12
                Case 8 ' ฮันเตอร์
                    GetPlayerMaxVital = (65 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 3) + 3
                Case 9 ' สไนเปอร์
                    GetPlayerMaxVital = (35 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 2) + 1
                Case 10 ' แอสแซสซิน
                    GetPlayerMaxVital = (150 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 3) + 5
                Case 11 ' ดาร์คลอร์ด
                    GetPlayerMaxVital = (120 * (GetPlayerLevel(index) / 5)) + (GetPlayerStat(index, intelligence) * 8) + 16
                Case Else ' Anything else - Default value
                    GetPlayerMaxVital = (GetPlayerLevel(index) / 2) + (GetPlayerStat(index, intelligence)) + 20
            End Select
            
            ' เพิ่ม MP ตาม สิ่งของที่สวมใส่
            
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).Add2 > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).Per2 > 0 Then
                            Per2a = Per2a + Item(GetPlayerEquipment(index, Weapon)).MPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital + Item(GetPlayerEquipment(index, Weapon)).MPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Armor) > 0 Then
                If Item(GetPlayerEquipment(index, Armor)).Add2 > 0 Then
                    If Item(GetPlayerEquipment(index, Armor)).Per2 > 0 Then
                            Per2a = Per2a + Item(GetPlayerEquipment(index, Armor)).MPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital + Item(GetPlayerEquipment(index, Armor)).MPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Helmet) > 0 Then
                If Item(GetPlayerEquipment(index, Helmet)).Add2 > 0 Then
                    If Item(GetPlayerEquipment(index, Helmet)).Per2 > 0 Then
                            Per2a = Per2a + Item(GetPlayerEquipment(index, Helmet)).MPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital + Item(GetPlayerEquipment(index, Helmet)).MPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Shield) > 0 Then
                If Item(GetPlayerEquipment(index, Shield)).Add2 > 0 Then
                    If Item(GetPlayerEquipment(index, Shield)).Per2 > 0 Then
                            Per2a = Per2a + Item(GetPlayerEquipment(index, Shield)).MPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital + Item(GetPlayerEquipment(index, Shield)).MPCase
                    End If
                End If
            End If
            
            ' ลด MP ตาม สิ่งของที่สวมใส่
            
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).Sub2 > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).Per2 > 0 Then
                            Per2s = Per2s + Item(GetPlayerEquipment(index, Weapon)).MPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital - Item(GetPlayerEquipment(index, Weapon)).MPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Armor) > 0 Then
                If Item(GetPlayerEquipment(index, Armor)).Sub2 > 0 Then
                    If Item(GetPlayerEquipment(index, Armor)).Per2 > 0 Then
                            Per2s = Per2s + Item(GetPlayerEquipment(index, Armor)).MPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital - Item(GetPlayerEquipment(index, Armor)).MPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Helmet) > 0 Then
                If Item(GetPlayerEquipment(index, Helmet)).Sub2 > 0 Then
                    If Item(GetPlayerEquipment(index, Helmet)).Per2 > 0 Then
                            Per2s = Per2s + Item(GetPlayerEquipment(index, Helmet)).MPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital - Item(GetPlayerEquipment(index, Helmet)).MPCase
                    End If
                End If
            End If
            
            If GetPlayerEquipment(index, Shield) > 0 Then
                If Item(GetPlayerEquipment(index, Shield)).Sub2 > 0 Then
                    If Item(GetPlayerEquipment(index, Shield)).Per2 > 0 Then
                            Per2s = Per2s + Item(GetPlayerEquipment(index, Shield)).MPCase
                        Else
                            GetPlayerMaxVital = GetPlayerMaxVital - Item(GetPlayerEquipment(index, Shield)).MPCase
                    End If
                End If
            End If
            
            If Per2a > 0 Then
                GetPlayerMaxVital = GetPlayerMaxVital + (GetPlayerMaxVital * (Per2a / 100))
            End If
            
            If Per2s > 0 Then
                GetPlayerMaxVital = (GetPlayerMaxVital * ((100 - Per2s) / 100))
            End If
            
            If GetPlayerMaxVital < 0 Then
                GetPlayerMaxVital = 1
            End If
            
    End Select
    
End Function

Function GetPlayerVitalRegen(ByVal index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (GetPlayerStat(index, Stats.Endurance) * 2) + (GetPlayerLevel(index) / 2) + (GetPlayerMaxVital(index, HP) * 0.01) + 1
        
            ' เพิ่ม RegenHP ตามสิ่งของที่สวมใส่
            
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).RegenHp > 0 Then
                    i = i + (Item(GetPlayerEquipment(index, Weapon)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(index, Armor) > 0 Then
                If Item(GetPlayerEquipment(index, Armor)).RegenHp > 0 Then
                    i = i + (Item(GetPlayerEquipment(index, Armor)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(index, Helmet) > 0 Then
                If Item(GetPlayerEquipment(index, Helmet)).RegenHp > 0 Then
                    i = i + (Item(GetPlayerEquipment(index, Helmet)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(index, Shield) > 0 Then
                If Item(GetPlayerEquipment(index, Shield)).RegenHp > 0 Then
                    i = i + (Item(GetPlayerEquipment(index, Shield)).RegenHp)
                End If
            End If
            
            If i <= 1 Then
                i = 1
            End If
        
        Case MP
            i = (GetPlayerStat(index, Stats.intelligence)) + (GetPlayerLevel(index) / 4) + (GetPlayerMaxVital(index, MP) * 0.01) + 1
    
            ' เพิ่ม RegenMP ตามสิ่งของที่สวมใส่
            
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).RegenMp > 0 Then
                    i = i + (Item(GetPlayerEquipment(index, Weapon)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(index, Armor) > 0 Then
                If Item(GetPlayerEquipment(index, Armor)).RegenMp > 0 Then
                    i = i + (Item(GetPlayerEquipment(index, Armor)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(index, Helmet) > 0 Then
                If Item(GetPlayerEquipment(index, Helmet)).RegenMp > 0 Then
                    i = i + (Item(GetPlayerEquipment(index, Helmet)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(index, Shield) > 0 Then
                If Item(GetPlayerEquipment(index, Shield)).RegenMp > 0 Then
                    i = i + (Item(GetPlayerEquipment(index, Shield)).RegenMp)
                End If
            End If
            
            If i <= 1 Then i = 1
            
            If i > 9999 Then i = 9999
            
    End Select
    
    GetPlayerVitalRegen = i
    
End Function

Function GetPlayerDamage(ByVal index As Long) As Long
    Dim weaponNum As Long
    Dim i As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, Weapon)
        If Not Item(weaponNum).Projectile.Pic > 0 Then
            GetPlayerDamage = (GetPlayerStat(index, Strength) * 3.5) + Item(weaponNum).Data2 + (GetPlayerLevel(index) * 2) '
        Else
            GetPlayerDamage = (GetPlayerStat(index, willpower) * 2.5) + Item(weaponNum).Data2 + (GetPlayerLevel(index) * 2) '
        End If
        
        GetPlayerDamage = rand(GetPlayerDamage * (Item(weaponNum).DmgLow / 100), GetPlayerDamage * (Item(weaponNum).DmgHigh / 100))
    Else
        GetPlayerDamage = (GetPlayerStat(index, Strength) * 3.5) + (GetPlayerLevel(index) * 2) '
        GetPlayerDamage = rand(GetPlayerDamage * 0.5, GetPlayerDamage)
    End If

End Function

Function GetPlayerDamageLHand(ByVal index As Long) As Long
    Dim weaponNum As Long
    Dim i As Long
    
    GetPlayerDamageLHand = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        weaponNum = GetPlayerEquipment(index, Shield)
        If Item(weaponNum).LHand > 0 Then
            GetPlayerDamageLHand = (GetPlayerStat(index, Strength) * 3.5) + Item(weaponNum).Data2 + (GetPlayerLevel(index) * 2) '
            GetPlayerDamageLHand = rand(GetPlayerDamageLHand * (Item(weaponNum).DmgLow / 100), GetPlayerDamageLHand * (Item(weaponNum).DmgHigh / 100))
        End If
    End If

End Function

Function GetPlayerCritDamage(ByVal index As Long, ByVal LHand As Boolean) As Double
    Dim weaponNum As Long
    Dim i As Long
    
    GetPlayerCritDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If LHand = False Then
    
        If GetPlayerEquipment(index, Weapon) > 0 Then
            If Item(GetPlayerEquipment(index, Weapon)).Projectile.Pic > 0 Then
                ' Damage = Damage + (Damage * (GetPlayerStat(index, willpower) / 100))
                GetPlayerCritDamage = 1.2 + (Item(GetPlayerEquipment(index, Weapon)).CritATK / 100) + (GetPlayerStat(index, willpower) / 100)
            Else
                ' Damage = Damage + (Damage * (GetPlayerStat(index, Strength) / 100))
                GetPlayerCritDamage = 1.2 + (Item(GetPlayerEquipment(index, Weapon)).CritATK / 100) + (GetPlayerStat(index, willpower) / 100)
            End If
        Else
            GetPlayerCritDamage = 1.2 + (GetPlayerStat(index, willpower) / 100)
        End If
    
    Else
    
        If GetPlayerEquipment(index, Shield) > 0 Then
            ' Damage = Damage + (Damage * (GetPlayerStat(index, Strength) / 100))
            GetPlayerCritDamage = 1.2 + (Item(GetPlayerEquipment(index, Shield)).CritATK / 100) + (GetPlayerStat(index, willpower) / 100)
        Else
            GetPlayerCritDamage = 1.2 + (GetPlayerStat(index, willpower) / 100)
        End If
    
    End If

End Function

Function GetPlayerMATK(ByVal index As Long) As Long
    Dim weaponNum As Long
    Dim i As Long
    
    GetPlayerMATK = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, Weapon)
        GetPlayerMATK = (GetPlayerStat(index, intelligence) * 4) + Item(weaponNum).MATK + (GetPlayerLevel(index) * 2) '  + (GetPlayerStat(index, Intelligence) * 1.5)
        GetPlayerMATK = rand(GetPlayerMATK * (Item(weaponNum).MagicLow / 100), GetPlayerMATK * (Item(weaponNum).MagicHigh / 100))
    Else
        GetPlayerMATK = (GetPlayerStat(index, intelligence) * 4) + (GetPlayerLevel(index) * 2) ' + (GetPlayerStat(index, Intelligence) * 1.5)
        GetPlayerMATK = rand(GetPlayerMATK * 0.5, GetPlayerMATK)
    End If
    
End Function

Function GetPlayerDef(ByVal index As Long) As Long
    Dim DefNum As Long
    Dim Def As Long
    Dim i As Long
    
    GetPlayerDef = 0
    Def = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(index, Armor)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Helmet) > 0 Then
        DefNum = GetPlayerEquipment(index, Helmet)
        Def = Def + Item(DefNum).Data2
    End If
    
    ' Fixed shield by allstar
    If GetPlayerEquipment(index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(index, Shield)
        Def = Def + Item(DefNum).Data2
    End If
    
    If Def <= 0 Then
        Def = 0
    End If
        
   If Not GetPlayerEquipment(index, Armor) > 0 And Not GetPlayerEquipment(index, Helmet) > 0 And Not GetPlayerEquipment(index, Shield) > 0 Then
        GetPlayerDef = (GetPlayerStat(index, Endurance) * 2) + (GetPlayerLevel(index) * 2) + Def
    Else
        GetPlayerDef = (GetPlayerStat(index, Endurance) * 2) + (GetPlayerLevel(index) * 2) + Def
    End If
    
    ' Check berserker
    If GetPlayerClass(index) = 4 Then ' Hulk Class None def
        GetPlayerDef = 0
    End If

End Function

Function GetPlayerMDEF(ByVal index As Long) As Long
    Dim DefNum As Long
    Dim MDEF As Long
    Dim i As Long
    
    GetPlayerMDEF = 0
    MDEF = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(index, Armor)
        MDEF = MDEF + Item(DefNum).MATK
    End If
    
    If GetPlayerEquipment(index, Helmet) > 0 Then
        DefNum = GetPlayerEquipment(index, Helmet)
        MDEF = MDEF + Item(DefNum).MATK
    End If
    
    ' Fixed shield get MDEF by allstar
     If GetPlayerEquipment(index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(index, Shield)
        MDEF = MDEF + Item(DefNum).MATK
     End If
    
    If MDEF <= 0 Then
        MDEF = 0
    End If

    GetPlayerMDEF = rand(GetPlayerStat(index, intelligence), (GetPlayerStat(index, intelligence) * 2)) + (GetPlayerLevel(index) * 2) + MDEF
    
    ' Check berserker
    If GetPlayerClass(index) = 4 Then ' Hulk Class None Mdef
        GetPlayerMDEF = 0
    End If

End Function

Function GetPlayerReflectDMG(ByVal index As Long) As Long
    Dim DMG As Long
    Dim DefNum As Long
    Dim i As Long
    
    DMG = 50
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(index, Weapon) > 0 Then
        DefNum = GetPlayerEquipment(index, Weapon)
        DMG = DMG + Item(DefNum).DmgReflect
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(index, Armor)
        DMG = DMG + Item(DefNum).DmgReflect
    End If
    
    If GetPlayerEquipment(index, Helmet) > 0 Then
        DefNum = GetPlayerEquipment(index, Helmet)
        DMG = DMG + Item(DefNum).DmgReflect
    End If
    
    ' Fixed shield get DMG by allstar
     If GetPlayerEquipment(index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(index, Shield)
        DMG = DMG + Item(DefNum).DmgReflect
     End If
    
    If DMG < 50 Then
        DMG = 50
    End If
    
    GetPlayerReflectDMG = (GetPlayerStat(index, Strength) / 10) + (GetPlayerLevel(index)) + DMG
    
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
            GetNpcMaxVital = NPC(npcNum).HP + NPC(npcNum).stat(Endurance)
        Case MP
            GetNpcMaxVital = 10 + (NPC(npcNum).stat(intelligence) * 5)
    End Select

End Function

Function GetNpcDEF(ByVal npcNum As Long) As Long
    Dim x As Long

    ' Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcDEF = 0
        Exit Function
    End If

    GetNpcDEF = NPC(npcNum).Def

End Function

Function GetNpcVitalRegen(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = rand((NPC(npcNum).RegenHp) / 2, (NPC(npcNum).RegenHp))
        Case MP
            i = rand((NPC(npcNum).RegenMp) / 2, (NPC(npcNum).RegenMp))
    End Select
    
    If i > 0 Then
        GetNpcVitalRegen = i
    Else
        GetNpcVitalRegen = 1
    End If

End Function

Function GetNpcDamage(ByVal npcNum As Long) As Long
    GetNpcDamage = (NPC(npcNum).stat(Stats.Strength) * 2) + NPC(npcNum).Damage
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal index As Long) As Boolean
Dim rate As Long
Dim stat As Long
Dim rndNum As Long

    CanPlayerBlock = False
    rate = 0
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        If Item(GetPlayerEquipment(index, Shield)).isDagger = False Then
            rate = Item(GetPlayerEquipment(index, Shield)).Reflect + (GetPlayerStat(index, Endurance) / 10)
        Else
            rate = (GetPlayerStat(index, Endurance) / 10)
        End If
    Else
         rate = (GetPlayerStat(index, Endurance) / 10)
    End If
    
    If GetPlayerEquipment(index, Helmet) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Helmet)).Reflect
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Armor)).Reflect
    End If
    
    If GetPlayerEquipment(index, Weapon) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Weapon)).Reflect
    End If
    
    rndNum = rand(1, 100)
    
    If rate > 80 Then rate = 80
    
    If rndNum <= rate Then
        CanPlayerBlock = True
    End If
    
End Function

Public Function CanPlayerLHand(ByVal index As Long) As Boolean

    CanPlayerLHand = False
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        If Item(GetPlayerEquipment(index, Shield)).LHand > 0 Then
            CanPlayerLHand = True
        End If
    End If
    
End Function

Public Function CanPlayerCrit(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False
    
    If GetPlayerEquipment(index, Weapon) > 0 Then
        rate = (GetPlayerStat(index, willpower) / 3) + Item(GetPlayerEquipment(index, Weapon)).CritRate
    Else
         rate = GetPlayerStat(index, willpower) / 5
    End If
    
    If rate > 80 Then rate = 80
    
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
    
End Function

Public Function CanPlayerAbsorbMagic(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerAbsorbMagic = False
    rate = 0
    
    If GetPlayerEquipment(index, Helmet) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Helmet)).AbsorbMagic
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Armor)).AbsorbMagic
    End If
    
    If GetPlayerEquipment(index, Weapon) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Weapon)).AbsorbMagic
    End If
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Shield)).AbsorbMagic
    End If
    
    rate = rate + GetPlayerStat(index, intelligence) / 10
    
    If rate > 80 Then rate = 80
    
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanPlayerAbsorbMagic = True
    End If
    
End Function

Public Function CanPlayerCritLHand(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCritLHand = False
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        rate = (GetPlayerStat(index, willpower) / 3) + Item(GetPlayerEquipment(index, Shield)).CritRate
    Else
         rate = GetPlayerStat(index, willpower) / 5
    End If
    
    If rate > 80 Then rate = 80
    
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanPlayerCritLHand = True
    End If
    
End Function

Public Function CanPlayerDodge(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False
    
    rate = GetPlayerStat(index, Agility) / 4
    
    If GetPlayerEquipment(index, Weapon) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Weapon)).Dodge
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Armor)).Dodge
    End If
    
    If GetPlayerEquipment(index, Helmet) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Helmet)).Dodge
    End If
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        rate = rate + Item(GetPlayerEquipment(index, Shield)).Dodge
    End If
    
    If rate > 80 Then rate = 80
    
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
    
    If TempPlayer(index).StunDuration > 0 Then CanPlayerDodge = False
    
End Function

Public Function CanNpcCrit(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    ' Fixed by Allstar
    ' rate = NPC(npcNum).stat(Stats.willpower) / 2
    rate = NPC(npcNum).CritRate
    
    If rate > 80 Then rate = 80
    
    rndNum = rand(1, 100)
    
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
    
End Function

Public Function CanNpcAbsorbMagic(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcAbsorbMagic = False

    ' Fixed by Allstar
    ' rate = NPC(npcNum).stat(Stats.willpower) / 2
    rate = NPC(npcNum).AbsorbMagic
    
    If rate > 80 Then rate = 80
    
    rndNum = rand(1, 100)
    
    If rndNum <= rate Then
        CanNpcAbsorbMagic = True
    End If
    
End Function

Public Function CanNpcDodge(ByVal mapnum As Long, ByVal npcNum As Long, ByVal mapNpcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = NPC(npcNum).Dodge
    If rate > 80 Then rate = 80
    
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).StunDuration > 0 Then CanNpcDodge = False
    
End Function

Public Function CanNpcParry(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = NPC(npcNum).Block
    If rate > 80 Then rate = 80
    
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If

End Function

' ตาบอด
Public Function CanNpcNoEye(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcNoEye = False

    rate = NPC(npcNum).stat(Stats.Endurance)
    If rate > 80 Then rate = 80
    
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanNpcNoEye = True
    End If

End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal mapNpcNum As Long)
Dim BlockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim Damage As Long
Dim DEFNPC As Long, NDEF As Boolean

    Damage = 0
    NDEF = False
    
    mapnum = GetPlayerMap(index)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, mapNpcNum) Then
        
        ' check if NPC can avoid the attack
        If CanNpcDodge(mapnum, npcNum, mapNpcNum) And Not CanPlayerCrit(index) Then
            SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
            SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
            'Call PlayerMsg(index, "Dodge : " & NpcDodge(mapnum, npcNum, mapNpcNum) & " %", Yellow)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(index)
        DEFNPC = NPC(npcNum).Def
        
        ' ระบบเจาะเกราะ
        If GetPlayerEquipment(index, Weapon) > 0 Then
            If Item(GetPlayerEquipment(index, Weapon)).NDEF > 0 Then
                NDEF = True
            End If
        End If
        
        ' x1.2 Critical ! +ระบบเพิ่มความแรงคริติคอล
        If CanPlayerCrit(index) Then
            Damage = Damage * GetPlayerCritDamage(index, False)
            'Call PlayerMsg(index, "Crit x " & GetPlayerCritDamage(index, False), Yellow)
            SendActionMsg mapnum, "คริติคอล !", BrightGreen, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
            SendAnimation mapnum, CRIT_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
        Else
            ' ระบบเจาะเกราะ
            If NDEF = True Then
                Damage = Damage - (DEFNPC - ((DEFNPC * Item(GetPlayerEquipment(index, Weapon)).NDEF) / 100))
            Else
                Damage = Damage - DEFNPC
            End If
        End If
            
        ' ระบบติดสะท้อนดาเมจ
        If CanNpcParry(npcNum) Then
            If Not CanPlayerDodge(index) Then
                SendActionMsg mapnum, "สะท้อน !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                Call PlayerMsg(index, "มอนส์เตอร์ " & Trim(NPC(npcNum).Name) & " ได้สะท้อนดาเมจกลับ.", BrightCyan)
                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
                Call NpcReflectPlayer(mapNpcNum, index, Damage * (NPC(npcNum).ReflectDmg / 100), 0)
                Exit Sub
            Else
                ' ถ้าผู้เล่นหลบการสะท้อน ?
                SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
                SendActionMsg mapnum, "หลบสะท้อน !", White, 1, (Player(index).x * 32), (Player(index).y * 32) - 16
            End If
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(index, mapNpcNum, Damage)
            
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
            SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
            SendActionMsg mapnum, "อ่อนหัด !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
            ' Call PlayerMsg(index, "การโจมตีของคุณไร้ผล.", BrightRed)
        End If
    End If
    
End Sub

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpcLHand(ByVal index As Long, ByVal mapNpcNum As Long)
Dim BlockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim Damage As Long
Dim DEFNPC As Long, NDEF As Boolean

    Damage = 0
    NDEF = False
    
    If Not CanPlayerLHand(index) Then Exit Sub
       
    ' Can we attack the npc?
    If CanPlayerAttackNpcLHand(index, mapNpcNum) Then
    
        mapnum = GetPlayerMap(index)
        npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(mapnum, npcNum, mapNpcNum) And Not CanPlayerCrit(index) Then
            SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
            SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamageLHand(index)
        DEFNPC = NPC(npcNum).Def
        
        ' ระบบเจาะเกราะ
        If GetPlayerEquipment(index, Shield) > 0 Then
            If Item(GetPlayerEquipment(index, Shield)).NDEF > 0 Then
                NDEF = True
            End If
        End If
        
        ' x1.2 Critical ! +ระบบเพิ่มความแรงคริติคอล
        If CanPlayerCritLHand(index) Then
            Damage = Damage * GetPlayerCritDamage(index, True)
            SendActionMsg mapnum, "คริติคอล !", BrightGreen, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
            SendAnimation mapnum, CRIT_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
        Else
            ' ระบบเจาะเกราะ
            If NDEF = True Then
                Damage = Damage - (DEFNPC - ((DEFNPC * Item(GetPlayerEquipment(index, Shield)).NDEF) / 100))
            Else
                Damage = Damage - DEFNPC
            End If
        End If
        
        ' ระบบติดสะท้อนดาเมจ
        If CanNpcParry(npcNum) Then
            If Not CanPlayerDodge(index) Then
                SendActionMsg mapnum, "สะท้อน !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                Call PlayerMsg(index, "มอนส์เตอร์ " & Trim(NPC(npcNum).Name) & " ได้สะท้อนดาเมจกลับ.", BrightCyan)
                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
                Call NpcReflectPlayer(mapNpcNum, index, Damage * (NPC(npcNum).ReflectDmg / 100), 1)
                Exit Sub
            Else
                ' ถ้าผู้เล่นหลบการสะท้อน ?
                SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
                SendActionMsg mapnum, "หลบสะท้อน !", White, 1, (Player(index).x * 32), (Player(index).y * 32) - 16
            End If
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpcLHand(index, mapNpcNum, Damage)
            
            ' ระบบ Vampire
        If GetPlayerEquipment(index, Shield) > 0 Then
            If Item(GetPlayerEquipment(index, Shield)).Vampire > 0 Then
            
                ' แก้ไขบัคดูดเลือดเกิน !!
                If GetPlayerMaxVital(index, HP) > Player(index).Vital(Vitals.HP) + Int((Damage * (Item(GetPlayerEquipment(index, Shield)).Vampire / 100))) Then
                    Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Int((Damage * (Item(GetPlayerEquipment(index, Shield)).Vampire / 100)))
                Else
                    Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, HP)
                End If
                
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                SendActionMsg GetPlayerMap(index), "+" & Int((Damage * (Item(GetPlayerEquipment(index, Shield)).Vampire / 100))), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                SendAnimation mapnum, Vampire_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
                SendVital index, HP
            End If
        End If
            
        Else
            SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
            SendActionMsg mapnum, "อ่อนหัด !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
            ' Call PlayerMsg(index, "การโจมตีของคุณไร้ผล.", BrightRed)
        End If
    End If
        
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal mapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim AttackSpeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).NPC(mapNpcNum).num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
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
        
     ' ตรวจสอบการมึน
    If TempPlayer(Attacker).StunDuration > 0 Then
        'Call PlayerMsg(Attacker, "คุณกำลังมึนงง.", BrightRed)
        'SendActionMsg GetPlayerMap(Attacker), "มึน !", BrightRed, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        Exit Function
    End If
        
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If Attacker = petowner Then
            'Call PlayerMsg(Attacker, "You can not attack your own pet.", BrightRed)
            Exit Function
        End If
    End If

If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
            Call PlayerMsg(Attacker, "ที่นี่คือเขตปลอดภัย ! คุณไม่สามารถโจมตีสัตว์เลี้ยงของผู้เล่นได้.", BrightRed)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerAccess(petowner) > ADMIN_MONITOR Then
            Call PlayerMsg(Attacker, "คุณไม่สามารถโจมตีสัตว์เลี้ยงของ " & GetPlayerName(petowner) & "'[GM] ได้ !", BrightRed)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
            Call PlayerMsg(Attacker, "GM ไม่สามารถโจมตีสัตว์เลี้ยงผู้เล่นอื่นได้.", BrightBlue)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerLevel(petowner) < 10 Then
            Call PlayerMsg(Attacker, "เจ้าของสัตว์เลี้ยง, " & GetPlayerName(petowner) & " มีเลเวลต่ำกว่า 10 ! ไม่สามารถโจมตีได้.", BrightRed)
            Exit Function
        End If
    End If

    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        If GetPlayerLevel(Attacker) < 10 Then
            Call PlayerMsg(Attacker, "คุณมีเลเวลต่ำกว่า 10, ไม่สามารถสั่งให้สัตว์เลี้ยงโจมตีใครได้ !", BrightRed)
            Exit Function
        End If
    End If
        
        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            AttackSpeed = ((2000 + Item(GetPlayerEquipment(Attacker, Weapon)).Speed) - Item(GetPlayerEquipment(Attacker, Weapon)).SpeedLow) - ((GetPlayerStat(Attacker, Stats.Agility) * 5))
        Else
            AttackSpeed = 2000 - ((GetPlayerStat(Attacker, Stats.Agility) * 5))
        End If
                
        ' Fixed bug attackspeed high
        If AttackSpeed < 100 Then
            AttackSpeed = 100
        End If

        If npcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + AttackSpeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x + 1
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x - 1
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y
            End Select

            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then

                TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
                TempPlayer(Attacker).Target = mapNpcNum
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
                                    PlayerMsg Attacker, Trim$(NPC(npcNum).Name) + " : " + Trim$(Quest(NPC(npcNum).QuestNum).Chat(2)), BrightGreen
                                    'QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).Chat(2)), 0
                                    Exit Function
                                End If
                            End If
                            
                        If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then

                        End If
                        
                            If NPC(npcNum).Quest = NO Then

                                If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                    PlayerMsg Attacker, Trim$(NPC(npcNum).Name) & " : " & Trim$(NPC(npcNum).AttackSay), White
                                    'SendActionMsg mapnum, Trim$(NPC(npcNum).AttackSay), White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
                                
                                Else

                                End If
                                Exit Function
                            End If

                        End If
                    End If
                End If
            End If
        End If
End Function

Public Function CanPlayerAttackNpcLHand(ByVal Attacker As Long, ByVal mapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).NPC(mapNpcNum).num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
    
        ' exit out early
        If IsSpell Then
             If npcNum > 0 Then
                If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    Dim petowner As Long
                    CanPlayerAttackNpcLHand = True
                    Exit Function
                End If
            End If
        End If
        
     ' ตรวจสอบการมึน
    If TempPlayer(Attacker).StunDuration > 0 Then
        'Call PlayerMsg(Attacker, "คุณกำลังมึนงง.", BrightRed)
        'SendActionMsg GetPlayerMap(Attacker), "มึน !", BrightRed, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        Exit Function
    End If
        
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If Attacker = petowner Then
            'Call PlayerMsg(Attacker, "You can not attack your own pet.", BrightRed)
            Exit Function
        End If
    End If


If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
            Call PlayerMsg(Attacker, "ที่นี่คือเขตปลอดภัย ! คุณไม่สามารถโจมตีสัตว์เลี้ยงของผู้เล่นได้.", BrightRed)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerAccess(petowner) > ADMIN_MONITOR Then
            Call PlayerMsg(Attacker, "คุณไม่สามารถโจมตีสัตว์เลี้ยงของ " & GetPlayerName(petowner) & "'[GM] ได้ !", BrightRed)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
            Call PlayerMsg(Attacker, "GM ไม่สามารถโจมตีสัตว์เลี้ยงผู้เล่นอื่นได้.", BrightBlue)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerLevel(petowner) < 10 Then
            Call PlayerMsg(Attacker, "เจ้าของสัตว์เลี้ยง, " & GetPlayerName(petowner) & " มีเลเวลต่ำกว่า 10 ! ไม่สามารถโจมตีได้.", BrightRed)
            Exit Function
        End If
    End If

    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        If GetPlayerLevel(Attacker) < 10 Then
            Call PlayerMsg(Attacker, "คุณมีเลเวลต่ำกว่า 10, ไม่สามารถสั่งให้สัตว์เลี้ยงโจมตีใครได้ !", BrightRed)
            Exit Function
        End If
    End If

        If npcNum > 0 Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x + 1
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x - 1
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y
            End Select

            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then

                TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
                TempPlayer(Attacker).Target = mapNpcNum
                SendTarget Attacker

                        If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            CanPlayerAttackNpcLHand = True
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
                                    PlayerMsg Attacker, Trim$(NPC(npcNum).Name) + " : " + Trim$(Quest(NPC(npcNum).QuestNum).Chat(2)), BrightGreen
                                    'QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).Chat(2)), 0
                                    Exit Function
                                End If
                            End If
                            
                        If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then

                        End If
                        
                            If NPC(npcNum).Quest = NO Then

                                If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                    PlayerMsg Attacker, Trim$(NPC(npcNum).Name) & " : " & Trim$(NPC(npcNum).AttackSay), White
                                    'SendActionMsg mapnum, Trim$(NPC(npcNum).AttackSay), White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
                                Else

                                End If
                                Exit Function
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
    Dim n, s, dRate As Long
    Dim i As Long
    Dim str As Long
    Dim Def As Long
    Dim DropRate, r As Double
    Dim mapnum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer
    Dim EXPRATE As Long
    Dim NoneExp As Boolean
    Dim Punch As Boolean
    
    Punch = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' projectiles
    If npcNum < 1 Then Exit Sub
    Name = Trim$(NPC(npcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    Else
        Punch = True
    End If
    
    ' สกิลติดตัวทำงานเมื่อถูกโจมตี?
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Spell(Player(Attacker).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Attacker).Spell(i)).PATK > 0 Then
                        If Spell(Player(Attacker).Spell(i)).PerSkill + (Spell(Player(Attacker).Spell(i)).S4 * Player(Attacker).skillLV(i)) >= rand(1, 100) Then
                            If Spell(Player(Attacker).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMove(Attacker, i, mapNpcNum, TARGET_TYPE_NPC)
                                'Call PlayerMsg(Attacker, "[Damage] : " & Player(Attacker).Spell(i), BrightGreen)
                            Else
                                Call CastSpellPassive(Attacker, i, mapNpcNum, TARGET_TYPE_NPC)
                                'Call PlayerMsg(Attacker, "[Heal] : " & Player(Attacker).Spell(i), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(Attacker), Trim$(Spell(Player(Attacker).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16
                            'Call PlayerMsg(Attacker, "[สกิลติดตัว] : " & Trim(Spell(Player(Attacker).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
    
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Sub
    
    'If Not CanPlayerAttackNpcLHand(Attacker, mapNpcNum) Then Exit Sub
    
    ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
    ' TempPlayer(Attacker).stopRegen = True
    ' TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) Then
        
        If CanPlayerLHand(Attacker) Then
            SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP), White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32) + 16, (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        Else
            SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP), White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        End If
        
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
            End If
        Else
            ' มีเสียงต่อยเมื่อไม่ใส่อาวุธ
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(mapnum, PUNCH_ANIM, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If

        ' Fixed animation on npc death when player spelled
        If spellnum > 0 Then
            Call SendAnimation(mapnum, Spell(spellnum).spellAnim, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
            SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        End If

        ' เช็คการคูณ Exp จากเซิฟเวอร์
        EXPRATE = frmServer.scrlExpRate.Value

        ' check spell level up ?
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Player(Attacker).skillLV(i) < MAX_SKILL_LEVEL Then
                    Player(Attacker).skillEXP(i) = Player(Attacker).skillEXP(i) + (NPC(npcNum).Level * (EXPRATE * 3))
                    Call CheckPlayerSkillUp(Attacker, i)
                    SendPlayerData Attacker
                Else
                    Player(Attacker).skillLV(i) = MAX_SKILL_LEVEL - 1
                    SendPlayerData Attacker
                End If
            End If
        Next
        
        SendActionMsg GetPlayerMap(Attacker), "EXP SKILL +" & NPC(npcNum).Level * (EXPRATE * 3), BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16
        
        ' เช็คการคูณ Exp จากเซิฟเวอร์
        EXPRATE = frmServer.scrlExpRate.Value
        ' Calculate exp to give attacker
        exp = (rand((NPC(npcNum).exp), (NPC(npcNum).EXP_max)) * EXPRATE)
        
        ' //////// สูตรคำนวน Exp แบบใหม่ !! //////////
        
        ' ถ้าไม่ใช่บอส
        If Not NPC(npcNum).BossNum > 0 Then
        
        ' ถ้าเวลต่างกับ npc เกิน 20 ให้ได้รับ Exp 10%
        
        If GetPlayerLevel(Attacker) > NPC(npcNum).Level + 20 Or GetPlayerLevel(Attacker) < NPC(npcNum).Level - 20 Then
            exp = exp * 0.1
            NoneExp = True
        Else
            NoneExp = False
        End If
        
        ' ถ้าเวลต่างกับ npc เกิน 10 ให้ได้รับ Exp ครึ่งเดียว
        If (GetPlayerLevel(Attacker) > NPC(npcNum).Level + 10 Or GetPlayerLevel(Attacker) < NPC(npcNum).Level - 10) And NoneExp = False Then
            exp = exp * 0.5
            ' Make sure we dont get less then 0
            If exp < 0 Then
                exp = 1
            End If
        End If
        
        End If
        
        ' /////////////////////////////////////////

        ' มีปาร์ตี้ ?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                'Call PlayerMsg(Attacker, "คุณได้รับ " & Exp & " Exp จากการสังหาร " & NPC(npcNum).Name, Yellow)
            Else
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                ' Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                Call SetPlayerExp(Attacker, 1)
                SendEXP Attacker
            End If
        Else
        ' ไม่มีปาร์ตี้
            If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                GivePlayerEXP Attacker, exp
                If exp > 0 Then
                    Call PlayerMsg(Attacker, "คุณได้รับ " & exp & " Exp จากการสังหาร " & NPC(npcNum).Name, Yellow)
                End If
            Else
                ' GivePlayerEXP Attacker, 0
                Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                Call SetPlayerExp(Attacker, 1)
                SendEXP Attacker
            End If
        End If
    
    ' npc ดรอปไอเทมเมื่อตาย
    For n = 1 To MAX_NPC_DROPS
        
        If NPC(npcNum).DropItem(n) = 0 Then Exit For

        r = Rnd
        DropRate = NPC(npcNum).DropChance(n) * frmServer.scrlDropRate.Value
        If DropRate > 1 Then DropRate = 1
        'Call PlayerMsg(Attacker, "ไอเทม : " & Trim(Item(NPC(npcNum).DropItem(n)).Name) & vbNewLine & "อัตราดรอป : " & Trim(DropRate * 100) & " ใน " & Trim(r * 100), Yellow)
        
        If (DropRate * 100) >= (r * 100) Then
            Call SpawnItem(NPC(npcNum).DropItem(n), NPC(npcNum).DropItemValue(n), mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If
        
    Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).NPC(mapNpcNum).num = 0
        MapNpc(mapnum).NPC(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' มีเสียง npc ตาย
        SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.seDie, 1
        
        'Checks if NPC was a pet
        If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
            Call PetDisband(MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner, mapnum) 'The pet was killed
        End If
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).NPC(mapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).NPC(mapNpcNum).HoT(i)
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
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = mapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
    
        ' ถ้า Npc ไม่ตายและได้รับดาเมจอยู่
        
        ' Kick System
        If n > 0 Then
            If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > 0 Then
                If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > rand(1, 100) Then
                    MapNpc(mapnum).NPC(mapNpcNum).StunDuration = 2
                    MapNpc(mapnum).NPC(mapNpcNum).StunTimer = GetTickCount
                    SendActionMsg mapnum, "ติดมึน !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) + 16
                    SendAnimation mapnum, Stun_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
                End If
            End If
        End If
        
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        
        If CanPlayerLHand(Attacker) Then
            SendActionMsg mapnum, "-" & Damage, White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32) + 16, (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        Else
            SendActionMsg mapnum, "-" & Damage, White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        End If
        
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
            End If
        Else
            If Not overTime Then
                ' มีเสียงต่อยเมื่อไม่ใส่อาวุธ
                SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
                Call SendAnimation(mapnum, PUNCH_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        If Damage > MapNpc(mapnum).NPC(mapNpcNum).GetDamage Or MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0 Or MapNpc(mapnum).NPC(mapNpcNum).Target = 0 Then
            MapNpc(mapnum).NPC(mapNpcNum).targetType = 1 ' player
            MapNpc(mapnum).NPC(mapNpcNum).Target = Attacker
            MapNpc(mapnum).NPC(mapNpcNum).GetDamage = Damage
        End If

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).NPC(i).num = MapNpc(mapnum).NPC(mapNpcNum).num Then
                    MapNpc(mapnum).NPC(i).Target = Attacker
                    MapNpc(mapnum).NPC(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        ' MapNpc(mapnum).NPC(mapNpcNum).stopRegen = True
        ' MapNpc(mapnum).NPC(mapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunNPC mapNpcNum, mapnum, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Npc mapnum, mapNpcNum, spellnum, Attacker
            End If
        End If

        SendMapNpcVitals mapnum, mapNpcNum
    End If

    If spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = GetTickCount
    End If
    
End Sub

Public Sub PlayerAttackNpcLHand(ByVal Attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n, s As Long
    Dim i As Long
    Dim str As Long
    Dim Def As Long
    Dim DropRate, r As Double
    Dim mapnum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer
    Dim EXPRATE As Long
    Dim NoneExp As Boolean
    Dim Punch As Boolean
    
    Punch = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' projectiles
    If npcNum < 1 Then Exit Sub
    Name = Trim$(NPC(npcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Shield) > 0 Then
        n = GetPlayerEquipment(Attacker, Shield)
    Else
        Punch = True
    End If

    ' สกิลติดตัวทำงานเมื่อถูกโจมตี? ใช้ได้แล้ว
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Spell(Player(Attacker).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Attacker).Spell(i)).PATK > 0 Then
                        If Spell(Player(Attacker).Spell(i)).PerSkill + (Spell(Player(Attacker).Spell(i)).S4 * Player(Attacker).skillLV(i)) >= rand(1, 100) Then
                            If Spell(Player(Attacker).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMoveLHand(Attacker, i, mapNpcNum, TARGET_TYPE_NPC)
                                'Call PlayerMsg(Attacker, "[Damage] : " & Player(Attacker).Spell(i), BrightGreen)
                            Else
                                Call CastSpellPassiveLHand(Attacker, i, mapNpcNum, TARGET_TYPE_NPC)
                                'Call PlayerMsg(Attacker, "[Heal] : " & Player(Attacker).Spell(i), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(Attacker), Trim$(Spell(Player(Attacker).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16
                            'Call PlayerMsg(Attacker, "[สกิลติดตัว] : " & Trim(Spell(Player(Attacker).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
        
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Sub
    
    If Not CanPlayerAttackNpcLHand(Attacker, mapNpcNum) Then Exit Sub
    
    If Damage >= MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) Then
    
        ' weapon say damage
        SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP), BrightCyan, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32) - 16, (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Shield)).Animation, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
            End If
        Else
            ' มีเสียงต่อยเมื่อไม่ใส่อาวุธ
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(mapnum, PUNCH_ANIM, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If
        
        ' Fixed animation on npc death when player spelled
        If spellnum > 0 Then
            Call SendAnimation(mapnum, Spell(spellnum).spellAnim, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
            SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        End If

        ' เช็คการคูณ Exp จากเซิฟเวอร์
        EXPRATE = frmServer.scrlExpRate.Value

        ' check spell level up ?
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Player(Attacker).skillLV(i) < MAX_SKILL_LEVEL Then
                    Player(Attacker).skillEXP(i) = Player(Attacker).skillEXP(i) + (NPC(npcNum).Level * (EXPRATE * 3))
                    Call CheckPlayerSkillUp(Attacker, i)
                    SendPlayerData Attacker
                Else
                    Player(Attacker).skillLV(i) = MAX_SKILL_LEVEL - 1
                    SendPlayerData Attacker
                End If
            End If
        Next
        
        SendActionMsg GetPlayerMap(Attacker), "EXP SKILL +" & NPC(npcNum).Level * (EXPRATE * 3), BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16
    
        ' เช็คการคูณ Exp จากเซิฟเวอร์
        EXPRATE = frmServer.scrlExpRate.Value
        ' Calculate exp to give attacker
        exp = (rand((NPC(npcNum).exp), (NPC(npcNum).EXP_max)) * EXPRATE)
        
        ' //////// สูตรคำนวน Exp แบบใหม่ !! //////////
        
        ' ถ้าไม่ใช่บอส
        If Not NPC(npcNum).BossNum > 0 Then
        
        ' ถ้าเวลต่างกับ npc เกิน 20 ให้ได้รับ Exp 10%
        
        If GetPlayerLevel(Attacker) > NPC(npcNum).Level + 20 Or GetPlayerLevel(Attacker) < NPC(npcNum).Level - 20 Then
            exp = exp * 0.1
            NoneExp = True
        Else
            NoneExp = False
        End If
        
        ' ถ้าเวลต่างกับ npc เกิน 10 ให้ได้รับ Exp ครึ่งเดียว
        If (GetPlayerLevel(Attacker) > NPC(npcNum).Level + 10 Or GetPlayerLevel(Attacker) < NPC(npcNum).Level - 10) And NoneExp = False Then
            exp = exp * 0.5
            ' Make sure we dont get less then 0
            If exp < 0 Then
                exp = 1
            End If
        End If
        
        End If
        
        ' /////////////////////////////////////////

        ' มีปาร์ตี้ ?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                'Call PlayerMsg(Attacker, "คุณได้รับ " & Exp & " Exp จากการสังหาร " & NPC(npcNum).Name, Yellow)
            Else
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                ' Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                Call SetPlayerExp(Attacker, 1)
                SendEXP Attacker
            End If
        Else
        ' ไม่มีปาร์ตี้
            If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                GivePlayerEXP Attacker, exp
                If exp > 0 Then
                    Call PlayerMsg(Attacker, "คุณได้รับ " & exp & " Exp จากการสังหาร " & NPC(npcNum).Name, Yellow)
                End If
            Else
                ' GivePlayerEXP Attacker, 0
                Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                Call SetPlayerExp(Attacker, 1)
                SendEXP Attacker
            End If
        End If
        
    ' npc ดรอปไอเทมเมื่อตาย
    For n = 1 To MAX_NPC_DROPS
        
        If NPC(npcNum).DropItem(n) = 0 Then Exit For

        r = Rnd
        DropRate = NPC(npcNum).DropChance(n) * frmServer.scrlDropRate.Value
        If DropRate > 1 Then DropRate = 1
        'Call PlayerMsg(Attacker, "ไอเทม : " & Trim(Item(NPC(npcNum).DropItem(n)).Name) & vbNewLine & "อัตราดรอป : " & Trim(DropRate * 100) & " ใน " & Trim(r * 100), Yellow)
        
        If (DropRate * 100) >= (r * 100) Then
            Call SpawnItem(NPC(npcNum).DropItem(n), NPC(npcNum).DropItemValue(n), mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If
        
    Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).NPC(mapNpcNum).num = 0
        MapNpc(mapnum).NPC(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' มีเสียง npc ตาย
        SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.seDie, 1
        
        'Checks if NPC was a pet
        If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
            Call PetDisband(MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner, mapnum) 'The pet was killed
        End If
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).NPC(mapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).NPC(mapNpcNum).HoT(i)
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
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = mapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
    
        ' ถ้า Npc ไม่ตายและได้รับดาเมจอยู่
        
        ' Kick System
        If n > 0 Then
            If Item(GetPlayerEquipment(Attacker, Shield)).Kick > 0 Then
                If Item(GetPlayerEquipment(Attacker, Shield)).Kick > rand(1, 100) Then
                    MapNpc(mapnum).NPC(mapNpcNum).StunDuration = 2
                    MapNpc(mapnum).NPC(mapNpcNum).StunTimer = GetTickCount
                    SendActionMsg mapnum, "ติดมึน !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) + 16
                    SendAnimation mapnum, Stun_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
                End If
            End If
        End If
        
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & Damage, BrightCyan, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32) - 16, (MapNpc(mapnum).NPC(mapNpcNum).y * 32) + 16
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
                
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Shield)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
            End If
        Else
            ' มีเสียงต่อยเมื่อไม่ใส่อาวุธ
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(mapnum, PUNCH_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
        End If

        ' Set the NPC target to the player
        If Damage > MapNpc(mapnum).NPC(mapNpcNum).GetDamage Or MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0 Or MapNpc(mapnum).NPC(mapNpcNum).Target = 0 Then
            MapNpc(mapnum).NPC(mapNpcNum).targetType = 1 ' player
            MapNpc(mapnum).NPC(mapNpcNum).Target = Attacker
            MapNpc(mapnum).NPC(mapNpcNum).GetDamage = Damage
        End If
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).NPC(i).num = MapNpc(mapnum).NPC(mapNpcNum).num Then
                     If Not MapNpc(mapnum).NPC(i).Target > 0 Then
                        MapNpc(mapnum).NPC(i).Target = Attacker
                        MapNpc(mapnum).NPC(i).targetType = 1 ' player
                    End If
                End If
            Next
        End If
        
        SendMapNpcVitals mapnum, mapNpcNum
    End If
    
End Sub

Public Sub PlayerPassiveNpc(ByVal Attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False, Optional ByVal Anim As Long, Optional ByVal Animated As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n, s As Long
    Dim i As Long
    Dim str As Long
    Dim Def As Long
    Dim DropRate, r As Double
    Dim mapnum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer
    Dim EXPRATE As Long
    Dim NoneExp As Boolean
    Dim Punch As Boolean
    
    Punch = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' projectiles
    If npcNum < 1 Then Exit Sub
    Name = Trim$(NPC(npcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    Else
        Punch = True
    End If
    
    ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
    ' TempPlayer(Attacker).stopRegen = True
    ' TempPlayer(Attacker).stopRegenTimer = GetTickCount
    
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Sub

    If Damage >= MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) Then
        
        If CanPlayerLHand(Attacker) Then
            SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP), White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32) + 16, (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        Else
            SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP), White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        End If
        
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        
        If Anim > 0 Then
            ' send animation
            If n > 0 Then
                If Not overTime Then
                    If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
                End If
            Else
                ' มีเสียงต่อยเมื่อไม่ใส่อาวุธ
                SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
                Call SendAnimation(mapnum, PUNCH_ANIM, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
            End If
        Else
            Call SendAnimation(mapnum, Anim, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If
        
        ' Fixed animation on npc death when player spelled
        If spellnum > 0 Then
            Call SendAnimation(mapnum, Spell(spellnum).spellAnim, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
            SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        End If
        
        ' เช็คการคูณ Exp จากเซิฟเวอร์
        EXPRATE = frmServer.scrlExpRate.Value

        ' check spell level up ?
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Player(Attacker).skillLV(i) < MAX_SKILL_LEVEL Then
                    Player(Attacker).skillEXP(i) = Player(Attacker).skillEXP(i) + (NPC(npcNum).Level * (EXPRATE * 3))
                    Call CheckPlayerSkillUp(Attacker, i)
                    SendPlayerData Attacker
                Else
                    Player(Attacker).skillLV(i) = MAX_SKILL_LEVEL - 1
                    SendPlayerData Attacker
                End If
            End If
        Next
        
        SendActionMsg GetPlayerMap(Attacker), "EXP SKILL +" & NPC(npcNum).Level * (EXPRATE * 3), BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16

        ' เช็คการคูณ Exp จากเซิฟเวอร์
        EXPRATE = frmServer.scrlExpRate.Value
        ' Calculate exp to give attacker
        exp = (rand((NPC(npcNum).exp), (NPC(npcNum).EXP_max)) * EXPRATE)
        
        ' //////// สูตรคำนวน Exp แบบใหม่ !! //////////
        
        ' ถ้าไม่ใช่บอส
        If Not NPC(npcNum).BossNum > 0 Then
        
        ' ถ้าเวลต่างกับ npc เกิน 20 ให้ได้รับ Exp 10%
        
        If GetPlayerLevel(Attacker) > NPC(npcNum).Level + 20 Or GetPlayerLevel(Attacker) < NPC(npcNum).Level - 20 Then
            exp = exp * 0.1
            NoneExp = True
        Else
            NoneExp = False
        End If
        
        ' ถ้าเวลต่างกับ npc เกิน 10 ให้ได้รับ Exp ครึ่งเดียว
        If (GetPlayerLevel(Attacker) > NPC(npcNum).Level + 10 Or GetPlayerLevel(Attacker) < NPC(npcNum).Level - 10) And NoneExp = False Then
            exp = exp * 0.5
            ' Make sure we dont get less then 0
            If exp < 0 Then
                exp = 1
            End If
        End If
        
        End If
        
        ' /////////////////////////////////////////

        ' มีปาร์ตี้ ?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                'Call PlayerMsg(Attacker, "คุณได้รับ " & Exp & " Exp จากการสังหาร " & NPC(npcNum).Name, Yellow)
            Else
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                ' Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                Call SetPlayerExp(Attacker, 1)
                SendEXP Attacker
            End If
        Else
        ' ไม่มีปาร์ตี้
            If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                GivePlayerEXP Attacker, exp
                If exp > 0 Then
                    Call PlayerMsg(Attacker, "คุณได้รับ " & exp & " Exp จากการสังหาร " & NPC(npcNum).Name, Yellow)
                End If
            Else
                ' GivePlayerEXP Attacker, 0
                Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                Call SetPlayerExp(Attacker, 1)
                SendEXP Attacker
            End If
        End If
    
    ' npc ดรอปไอเทมเมื่อตาย
    For n = 1 To MAX_NPC_DROPS
        
        If NPC(npcNum).DropItem(n) = 0 Then Exit For

        r = Rnd
        DropRate = NPC(npcNum).DropChance(n) * frmServer.scrlDropRate.Value
        If DropRate > 1 Then DropRate = 1
        'Call PlayerMsg(Attacker, "ไอเทม : " & Trim(Item(NPC(npcNum).DropItem(n)).Name) & vbNewLine & "อัตราดรอป : " & Trim(DropRate * 100) & " ใน " & Trim(r * 100), Yellow)
        
        If (DropRate * 100) >= (r * 100) Then
            Call SpawnItem(NPC(npcNum).DropItem(n), NPC(npcNum).DropItemValue(n), mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If
        
    Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).NPC(mapNpcNum).num = 0
        MapNpc(mapnum).NPC(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' มีเสียง npc ตาย
        SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.seDie, 1
        
        'Checks if NPC was a pet
        If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
            Call PetDisband(MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner, mapnum) 'The pet was killed
        End If
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).NPC(mapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).NPC(mapNpcNum).HoT(i)
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
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = mapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
    
        ' ถ้า Npc ไม่ตายและได้รับดาเมจอยู่
        
        ' Kick System
        If n > 0 Then
            If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > 0 Then
                If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > rand(1, 100) Then
                    MapNpc(mapnum).NPC(mapNpcNum).StunDuration = 2
                    MapNpc(mapnum).NPC(mapNpcNum).StunTimer = GetTickCount
                    SendActionMsg mapnum, "ติดมึน !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) + 16
                    SendAnimation mapnum, Stun_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
                End If
            End If
        End If
        
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        
        If CanPlayerLHand(Attacker) Then
            SendActionMsg mapnum, "-" & Damage, White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32) + 16, (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        Else
            SendActionMsg mapnum, "-" & Damage, White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        End If
        
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        
        If Animated = False Then
            ' send animation
            If n > 0 Then
                If Not overTime Then
                    If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
                End If
            Else
                If Animated = False Then
                    ' มีเสียงต่อยเมื่อไม่ใส่อาวุธ
                    SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
                    Call SendAnimation(mapnum, PUNCH_ANIM, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
                End If
            End If
        Else
            'SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.seAnimation, Anim
            Call SendAnimation(mapnum, Anim, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
        End If

        ' Set the NPC target to the player
        If Damage > MapNpc(mapnum).NPC(mapNpcNum).GetDamage Or MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0 Or MapNpc(mapnum).NPC(mapNpcNum).Target = 0 Then
            MapNpc(mapnum).NPC(mapNpcNum).targetType = 1 ' player
            MapNpc(mapnum).NPC(mapNpcNum).Target = Attacker
            MapNpc(mapnum).NPC(mapNpcNum).GetDamage = Damage
        End If

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).NPC(i).num = MapNpc(mapnum).NPC(mapNpcNum).num Then
                    MapNpc(mapnum).NPC(i).Target = Attacker
                    MapNpc(mapnum).NPC(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        ' MapNpc(mapnum).NPC(mapNpcNum).stopRegen = True
        ' MapNpc(mapnum).NPC(mapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunNPC mapNpcNum, mapnum, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Npc mapnum, mapNpcNum, spellnum, Attacker
            End If
        End If

        SendMapNpcVitals mapnum, mapNpcNum
    End If
    
End Sub

Public Sub PlayerPassiveNpcLHand(ByVal Attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False, Optional ByVal Anim As Long, Optional ByVal Animated As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n, s As Long
    Dim i As Long
    Dim str As Long
    Dim Def As Long
    Dim DropRate, r As Double
    Dim mapnum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer
    Dim EXPRATE As Long
    Dim NoneExp As Boolean
    Dim Punch As Boolean
    
    Punch = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' projectiles
    If npcNum < 1 Then Exit Sub
    Name = Trim$(NPC(npcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Shield) > 0 Then
        n = GetPlayerEquipment(Attacker, Shield)
    Else
        Punch = True
    End If
    
    ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
    ' TempPlayer(Attacker).stopRegen = True
    ' TempPlayer(Attacker).stopRegenTimer = GetTickCount
    
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Sub

    If Damage >= MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) Then
        
        If CanPlayerLHand(Attacker) Then
            SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP), White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32) + 16, (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        Else
            SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP), White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        End If
        
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        
        If Anim > 0 Then
            ' send animation
            If n > 0 Then
                If Not overTime Then
                    If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Shield)).Animation, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
                End If
            Else
                ' มีเสียงต่อยเมื่อไม่ใส่อาวุธ
                SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
                Call SendAnimation(mapnum, PUNCH_ANIM, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
            End If
        Else
            Call SendAnimation(mapnum, Anim, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If
        
        ' Fixed animation on npc death when player spelled
        If spellnum > 0 Then
            Call SendAnimation(mapnum, Spell(spellnum).spellAnim, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
            SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        End If
        
        ' เช็คการคูณ Exp จากเซิฟเวอร์
        EXPRATE = frmServer.scrlExpRate.Value

        ' check spell level up ?
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Player(Attacker).skillLV(i) < MAX_SKILL_LEVEL Then
                    Player(Attacker).skillEXP(i) = Player(Attacker).skillEXP(i) + (NPC(npcNum).Level * (EXPRATE * 3))
                    Call CheckPlayerSkillUp(Attacker, i)
                    SendPlayerData Attacker
                Else
                    Player(Attacker).skillLV(i) = MAX_SKILL_LEVEL - 1
                    SendPlayerData Attacker
                End If
            End If
        Next
        
        SendActionMsg GetPlayerMap(Attacker), "EXP SKILL +" & NPC(npcNum).Level * (EXPRATE * 3), BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16

        ' เช็คการคูณ Exp จากเซิฟเวอร์
        EXPRATE = frmServer.scrlExpRate.Value
        ' Calculate exp to give attacker
        exp = (rand((NPC(npcNum).exp), (NPC(npcNum).EXP_max)) * EXPRATE)
        
        ' //////// สูตรคำนวน Exp แบบใหม่ !! //////////
        
        ' ถ้าไม่ใช่บอส
        If Not NPC(npcNum).BossNum > 0 Then
        
        ' ถ้าเวลต่างกับ npc เกิน 20 ให้ได้รับ Exp 10%
        
        If GetPlayerLevel(Attacker) > NPC(npcNum).Level + 20 Or GetPlayerLevel(Attacker) < NPC(npcNum).Level - 20 Then
            exp = exp * 0.1
            NoneExp = True
        Else
            NoneExp = False
        End If
        
        ' ถ้าเวลต่างกับ npc เกิน 10 ให้ได้รับ Exp ครึ่งเดียว
        If (GetPlayerLevel(Attacker) > NPC(npcNum).Level + 10 Or GetPlayerLevel(Attacker) < NPC(npcNum).Level - 10) And NoneExp = False Then
            exp = exp * 0.5
            ' Make sure we dont get less then 0
            If exp < 0 Then
                exp = 1
            End If
        End If
        
        End If
        
        ' /////////////////////////////////////////

        ' มีปาร์ตี้ ?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                'Call PlayerMsg(Attacker, "คุณได้รับ " & Exp & " Exp จากการสังหาร " & NPC(npcNum).Name, Yellow)
            Else
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                ' Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                Call SetPlayerExp(Attacker, 1)
                SendEXP Attacker
            End If
        Else
        ' ไม่มีปาร์ตี้
            If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                GivePlayerEXP Attacker, exp
                If exp > 0 Then
                    Call PlayerMsg(Attacker, "คุณได้รับ " & exp & " Exp จากการสังหาร " & NPC(npcNum).Name, Yellow)
                End If
            Else
                ' GivePlayerEXP Attacker, 0
                Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                Call SetPlayerExp(Attacker, 1)
                SendEXP Attacker
            End If
        End If
    
    ' npc ดรอปไอเทมเมื่อตาย
    For n = 1 To MAX_NPC_DROPS
        
        If NPC(npcNum).DropItem(n) = 0 Then Exit For

        r = Rnd
        DropRate = NPC(npcNum).DropChance(n) * frmServer.scrlDropRate.Value
        If DropRate > 1 Then DropRate = 1
        'Call PlayerMsg(Attacker, "ไอเทม : " & Trim(Item(NPC(npcNum).DropItem(n)).Name) & vbNewLine & "อัตราดรอป : " & Trim(DropRate * 100) & " ใน " & Trim(r * 100), Yellow)
        
        If (DropRate * 100) >= (r * 100) Then
            Call SpawnItem(NPC(npcNum).DropItem(n), NPC(npcNum).DropItemValue(n), mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If
        
    Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).NPC(mapNpcNum).num = 0
        MapNpc(mapnum).NPC(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' มีเสียง npc ตาย
        SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.seDie, 1
        
        'Checks if NPC was a pet
        If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
            Call PetDisband(MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner, mapnum) 'The pet was killed
        End If
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).NPC(mapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).NPC(mapNpcNum).HoT(i)
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
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = mapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
    
        ' ถ้า Npc ไม่ตายและได้รับดาเมจอยู่
        
        ' Kick System
        If n > 0 Then
            If Item(GetPlayerEquipment(Attacker, Shield)).Kick > 0 Then
                If Item(GetPlayerEquipment(Attacker, Shield)).Kick > rand(1, 100) Then
                    MapNpc(mapnum).NPC(mapNpcNum).StunDuration = 2
                    MapNpc(mapnum).NPC(mapNpcNum).StunTimer = GetTickCount
                    SendActionMsg mapnum, "ติดมึน !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) + 16
                    SendAnimation mapnum, Stun_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
                End If
            End If
        End If
        
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a Shield and say damage
        
        If CanPlayerLHand(Attacker) Then
            SendActionMsg mapnum, "-" & Damage, White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32) + 16, (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        Else
            SendActionMsg mapnum, "-" & Damage, White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        End If
        
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        
        If Animated = False Then
            ' send animation
            If n > 0 Then
                If Not overTime Then
                    If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Shield)).Animation, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
                End If
            Else
                If Animated = False Then
                    ' มีเสียงต่อยเมื่อไม่ใส่อาวุธ
                    SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
                    Call SendAnimation(mapnum, PUNCH_ANIM, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
                End If
            End If
        Else
            'SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.seAnimation, Anim
            Call SendAnimation(mapnum, Anim, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
        End If

        ' Set the NPC target to the player
        If Damage > MapNpc(mapnum).NPC(mapNpcNum).GetDamage Or MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0 Or MapNpc(mapnum).NPC(mapNpcNum).Target = 0 Then
            MapNpc(mapnum).NPC(mapNpcNum).targetType = 1 ' player
            MapNpc(mapnum).NPC(mapNpcNum).Target = Attacker
            MapNpc(mapnum).NPC(mapNpcNum).GetDamage = Damage
        End If

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).NPC(i).num = MapNpc(mapnum).NPC(mapNpcNum).num Then
                    MapNpc(mapnum).NPC(i).Target = Attacker
                    MapNpc(mapnum).NPC(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        ' MapNpc(mapnum).NPC(mapNpcNum).stopRegen = True
        ' MapNpc(mapnum).NPC(mapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunNPC mapNpcNum, mapnum, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Npc mapnum, mapNpcNum, spellnum, Attacker
            End If
        End If

        SendMapNpcVitals mapnum, mapNpcNum
    End If
        
End Sub


Public Sub PlayerReflectNpc(ByVal Attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, ByVal LHand As Byte, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n, s As Long
    Dim i As Long
    Dim str As Long
    Dim Def As Long
    Dim DropRate, r As Double
    Dim mapnum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer
    Dim EXPRATE As Long
    Dim NoneExp As Boolean
    Dim Punch As Boolean
    
    Punch = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' projectiles
    If npcNum < 1 Then Exit Sub
    Name = Trim$(NPC(npcNum).Name)
    
    ' Check for weapon
    n = 0
    
    If GetPlayerEquipment(Attacker, Shield) > 0 Then
        n = GetPlayerEquipment(Attacker, Shield)
    ElseIf GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    Else
        Punch = True
    End If
    
    ' แก้ไขการแจ้งเตือนเมื่อสะท้อนไม่มีดาเมจ
    If Damage <= 0 Then
        SendActionMsg GetPlayerMap(Attacker), "อ่อนหัด !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 8
        SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
        Exit Sub
    End If
    
    ' สกิลติดตัวทำงานเมื่อถูกโจมตี? ใช้ได้แล้ว
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Spell(Player(Attacker).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Attacker).Spell(i)).PATK > 0 Then
                        If Spell(Player(Attacker).Spell(i)).PerSkill + (Spell(Player(Attacker).Spell(i)).S4 * Player(Attacker).skillLV(i)) >= rand(1, 100) Then
                            If Spell(Player(Attacker).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMove(Attacker, i, mapNpcNum, TARGET_TYPE_NPC)
                                'Call PlayerMsg(Attacker, "[Damage] : " & Player(Attacker).Spell(i), BrightGreen)
                            Else
                                Call CastSpell(Attacker, i, mapNpcNum, TARGET_TYPE_NPC)
                                'Call PlayerMsg(Attacker, "[Heal] : " & Player(Attacker).Spell(i), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(Attacker), Trim$(Spell(Player(Attacker).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16
                            'Call PlayerMsg(Attacker, "[สกิลติดตัว] : " & Trim(Spell(Player(Attacker).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
        
       ' Fixed reflect
        'If spellnum > 0 Then
        '    If Not CanPlayerAttackNpcLHand(Attacker, mapNpcNum, True) Then Exit Sub
        'Else
        '    If Not CanPlayerAttackNpcLHand(Attacker, mapNpcNum) Then Exit Sub
        'End If

    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Sub

    If Damage >= MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) Then
    
        ' weapon say damage
        SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP), Yellow, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
        
        ' send animation
        If Not spellnum > 0 Then
        
        If n > 0 Then
            If Not overTime Then
                If LHand = 1 Then
                    If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Shield)).Animation, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
                Else
                    If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
                End If
            End If
        Else
            ' มีเสียงต่อยเมื่อไม่ใส่อาวุธ
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(mapnum, PUNCH_ANIM, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If
        
        End If
        
        ' Fixed animation on npc death when player spelled
        If spellnum > 0 Then
            Call SendAnimation(mapnum, Spell(spellnum).spellAnim, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
            SendMapSound Attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
        End If
        
        ' เช็คการคูณ Exp จากเซิฟเวอร์
        EXPRATE = frmServer.scrlExpRate.Value

        ' check spell level up ?
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Player(Attacker).skillLV(i) < MAX_SKILL_LEVEL Then
                    Player(Attacker).skillEXP(i) = Player(Attacker).skillEXP(i) + (NPC(npcNum).Level * (EXPRATE * 3))
                    Call CheckPlayerSkillUp(Attacker, i)
                    SendPlayerData Attacker
                Else
                    Player(Attacker).skillLV(i) = MAX_SKILL_LEVEL - 1
                    SendPlayerData Attacker
                End If
            End If
        Next
        
        SendActionMsg GetPlayerMap(Attacker), "EXP SKILL +" & NPC(npcNum).Level * (EXPRATE * 3), BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16

        ' เช็คการคูณ Exp จากเซิฟเวอร์
        EXPRATE = frmServer.scrlExpRate.Value
        ' Calculate exp to give attacker
        exp = (rand((NPC(npcNum).exp), (NPC(npcNum).EXP_max)) * EXPRATE)
        
        ' //////// สูตรคำนวน Exp แบบใหม่ !! //////////
        
        ' ถ้าไม่ใช่บอส
        If Not NPC(npcNum).BossNum > 0 Then
        
        ' ถ้าเวลต่างกับ npc เกิน 20 ให้ได้รับ Exp 10%
        
        If GetPlayerLevel(Attacker) > NPC(npcNum).Level + 20 Or GetPlayerLevel(Attacker) < NPC(npcNum).Level - 20 Then
            exp = exp * 0.1
            NoneExp = True
        Else
            NoneExp = False
        End If
        
        ' ถ้าเวลต่างกับ npc เกิน 10 ให้ได้รับ Exp ครึ่งเดียว
        If (GetPlayerLevel(Attacker) > NPC(npcNum).Level + 10 Or GetPlayerLevel(Attacker) < NPC(npcNum).Level - 10) And NoneExp = False Then
            exp = exp * 0.5
            ' Make sure we dont get less then 0
            If exp < 0 Then
                exp = 1
            End If
        End If
        
        End If
        
        ' /////////////////////////////////////////

        ' มีปาร์ตี้ ?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
            Else
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                ' Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                Call SetPlayerExp(Attacker, 1)
                SendEXP Attacker
            End If
        Else
        ' ไม่มีปาร์ตี้
            If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                GivePlayerEXP Attacker, exp
                If exp > 0 Then
                    Call PlayerMsg(Attacker, "คุณได้รับ " & exp & " Exp จากการสังหาร " & NPC(npcNum).Name, Yellow)
                End If
            Else
                ' GivePlayerEXP Attacker, 0
                Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                Call SetPlayerExp(Attacker, 1)
                SendEXP Attacker
            End If
        End If
        
    ' npc ดรอปไอเทมเมื่อตาย
    For n = 1 To MAX_NPC_DROPS
        
        If NPC(npcNum).DropItem(n) = 0 Then Exit For

        r = Rnd
        DropRate = NPC(npcNum).DropChance(n) * frmServer.scrlDropRate.Value
        If DropRate > 1 Then DropRate = 1
        'Call PlayerMsg(Attacker, "ไอเทม : " & Trim(Item(NPC(npcNum).DropItem(n)).Name) & vbNewLine & "อัตราดรอป : " & Trim(DropRate * 100) & " ใน " & Trim(r * 100), Yellow)
        
        If (DropRate * 100) >= (r * 100) Then
            Call SpawnItem(NPC(npcNum).DropItem(n), NPC(npcNum).DropItemValue(n), mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If
        
    Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).NPC(mapNpcNum).num = 0
        MapNpc(mapnum).NPC(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' มีเสียง npc ตาย
        SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.seDie, 1
        
        'Checks if NPC was a pet
        If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
            Call PetDisband(MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner, mapnum) 'The pet was killed
        End If
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).NPC(mapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).NPC(mapNpcNum).HoT(i)
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
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = mapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
    
        ' ถ้า Npc ไม่ตายและได้รับดาเมจอยู่
        
        ' Kick System
        If n > 0 Then
            If LHand = 0 Then
                If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > 0 Then
                    If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > rand(1, 100) Then
                        MapNpc(mapnum).NPC(mapNpcNum).StunDuration = 2
                        MapNpc(mapnum).NPC(mapNpcNum).StunTimer = GetTickCount
                        SendActionMsg mapnum, "ติดมึน !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) + 16
                        SendAnimation mapnum, Stun_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
                    End If
                End If
            Else
                If Item(GetPlayerEquipment(Attacker, Shield)).Kick > 0 Then
                    If Item(GetPlayerEquipment(Attacker, Shield)).Kick > rand(1, 100) Then
                        MapNpc(mapnum).NPC(mapNpcNum).StunDuration = 2
                        MapNpc(mapnum).NPC(mapNpcNum).StunTimer = GetTickCount
                        SendActionMsg mapnum, "ติดมึน !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) + 16
                        SendAnimation mapnum, Stun_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
                    End If
                End If
            End If
        End If
        
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & Damage, Yellow, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
                
        If Not spellnum > 0 Then
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If LHand = 1 Then
                    If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Shield)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
                Else
                    If spellnum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
                End If
            End If
        Else
            ' มีเสียงต่อยเมื่อไม่ใส่อาวุธ
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(mapnum, PUNCH_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
        End If
        
        End If

        ' Set the NPC target to the player
        If Damage > MapNpc(mapnum).NPC(mapNpcNum).GetDamage Or MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0 Or MapNpc(mapnum).NPC(mapNpcNum).Target = 0 Then
            MapNpc(mapnum).NPC(mapNpcNum).targetType = 1 ' player
            MapNpc(mapnum).NPC(mapNpcNum).Target = Attacker
            MapNpc(mapnum).NPC(mapNpcNum).GetDamage = Damage
        End If

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).NPC(i).num = MapNpc(mapnum).NPC(mapNpcNum).num Then
                     If Not MapNpc(mapnum).NPC(i).Target > 0 Then
                        MapNpc(mapnum).NPC(i).Target = Attacker
                        MapNpc(mapnum).NPC(i).targetType = 1 ' player
                    End If
                End If
            Next
        End If
              
        SendMapNpcVitals mapnum, mapNpcNum
    End If
    
End Sub


' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal index As Long)
Dim mapnum As Long, npcNum As Long, BlockAmount As Long, Damage As Long
Dim Buffer As clsBuffer
Dim nMax As Long, n As Long

    nMax = 0

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(mapNpcNum, index) Then
        mapnum = GetPlayerMap(index)
        npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
        ' check if PLAYER can avoid the attack
        
        ' โจมตีไม่โดนเมื่อหายตัว
        If Player(index).BuffStatus(BUFF_INVISIBLE) = BUFF_INVISIBLE And NPC(mapNpcNum).BossNum < 1 Then
            SendActionMsg mapnum, "หาไม่เจอ !", White, 1, (Player(index).x * 32), (Player(index).y * 32) - 16
            ' Set NPC target to 0
            MapNpc(mapnum).NPC(mapNpcNum).Target = 0
            MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
            MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
            Exit Sub
        End If
        
        If CanPlayerDodge(index) And Not CanNpcCrit(npcNum) Then
            SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
            SendActionMsg mapnum, "พลาด !", White, 1, (Player(index).x * 32), (Player(index).y * 32) - 16
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcNum)

        ' * DMG if crit hit
        If CanNpcCrit(npcNum) Then
            Damage = Damage * (NPC(npcNum).CritChange / 10)
            SendActionMsg mapnum, "คริติคอล !", BrightCyan, 1, (Player(index).x * 32), (Player(index).y * 32) - 16
            SendAnimation mapnum, CRIT_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
            'Call PlayerMsg(index, "คุณโดนมอนส์เตอร์." & NPC(npcNum).Name & " โจมตีแบบ โป๊ะเชะ", Yellow)
        Else
            Damage = Damage - GetPlayerDef(index)
        End If
        
        ' randomise for up to 50% lower than max hit
        Damage = rand(Damage * 0.5, Damage)
        
        If GetPlayerEquipment(index, Shield) > 0 Then
            If Item(GetPlayerEquipment(index, Shield)).LHand > 0 Then
                nMax = 1
            Else
                nMax = 0
            End If
        End If
        
        n = rand(0, nMax)
        
        ' ระบบสะท้อนดาเมจ
        If CanPlayerBlock(index) Then
            If Not CanNpcDodge(mapnum, npcNum, mapNpcNum) Then
                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
                SendActionMsg mapnum, "สะท้อน !", Yellow, 1, (Player(index).x * 32), (Player(index).y * 32) - 16
                Call PlayerMsg(index, "คุณได้สะท้อนดาเมจไปหา " & Trim(NPC(npcNum).Name), BrightCyan)
                    If n = 0 Then
                        Call PlayerReflectNpc(index, mapNpcNum, Damage * (GetPlayerReflectDMG(index) / 100), 0)
                    Else
                        Call PlayerReflectNpc(index, mapNpcNum, Damage * (GetPlayerReflectDMG(index) / 100), 1)
                    End If
                Exit Sub
            Else
                ' ถ้า npc หลบการสะท้อน
                SendActionMsg mapnum, "หลบสะท้อน !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, mapNpcNum
            End If
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(mapNpcNum, index, Damage)
        Else
            SendActionMsg mapnum, "อ่อนหัด !", White, 1, (Player(index).x * 32), (Player(index).y * 32) - 16
            SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, index
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal index As Long) As Boolean
    Dim mapnum, npcNum, DistanceX, DistanceY As Long
    Dim Buffer As clsBuffer
    Dim petowner As Long

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).NPC(mapNpcNum).num <= 0 Then
        Exit Function
    End If
    
    'check if the NPC attacking us is actually our pet.
    'We don't want a rebellion on our hands now do we?
        
    mapnum = GetPlayerMap(index)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    'check if the NPC attacking us is actually our pet.
    'We don't want a rebellion on our hands now do we?
    
    If MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner = index Then Exit Function

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Npc aspd
    If NPC(MapNpc(mapnum).NPC(mapNpcNum).num).AttackSpeed > 0 Then
        If GetTickCount < MapNpc(mapnum).NPC(mapNpcNum).AttackTimer + NPC(MapNpc(mapnum).NPC(mapNpcNum).num).AttackSpeed Then
            Exit Function
        End If
    Else
        If GetTickCount < MapNpc(mapnum).NPC(mapNpcNum).AttackTimer + 100 Then
            Exit Function
        End If
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).StunDuration > 0 Then
        SendActionMsg mapnum, "มึน !", White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        Exit Function
    End If

    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SNpcAttack
    Buffer.WriteLong mapNpcNum
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing

    MapNpc(mapnum).NPC(mapNpcNum).AttackTimer = GetTickCount
    
    If Map(mapnum).Moral <> MAP_MORAL_PETARENA Then
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If Not Map(GetPlayerMap(petowner)).Moral = MAP_MORAL_NONE Then
            Call PlayerMsg(petowner, "ที่นี่คือเขตปลอดภัย ! คุณไม่สามารถโจมตีใครด้วยสัตว์เลี้ยงได้.", BrightRed)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerAccess(petowner) > ADMIN_MONITOR Then
            Call PlayerMsg(petowner, "GM ไม่สามารถโจมตีสัตว์เลี้ยงผู้เล่นได้.", BrightBlue)
            Exit Function
        End If
    End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerAccess(index) > ADMIN_MONITOR Then
            Call PlayerMsg(petowner, "คุณไม่สามารถโจมตี " & GetPlayerName(index) & " ด้วยสัตว์เลี้ยงได้ !", BrightRed)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerLevel(petowner) < 10 Then
        CanNpcAttackPlayer = False
            Call PlayerMsg(petowner, "คุณมีเลเวลต่ำกว่า 10, ไม่สามารถโจมตีผู้เล่นอื่นได้ !", BrightRed)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerLevel(index) < 10 Then
        CanNpcAttackPlayer = False
            Call PlayerMsg(petowner, GetPlayerName(petowner) & " มีเลเวลต่ำกว่า 10, คุณไม่สามารถโจมตีได้ !", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If npcNum > 0 Then
            DistanceX = MapNpc(mapnum).NPC(mapNpcNum).x - GetPlayerX(index)
            DistanceY = MapNpc(mapnum).NPC(mapNpcNum).y - GetPlayerY(index)
            If DistanceX < 0 Then DistanceX = DistanceX * -1
            If DistanceY < 0 Then DistanceY = DistanceY * -1
            ' Check if at same coordinates
            If DistanceX <= 1 And DistanceY <= 1 Then
                CanNpcAttackPlayer = True
            ElseIf DistanceX > 17 Or DistanceY > 14 Then
                MapNpc(mapnum).NPC(mapNpcNum).Target = 0
                MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
                MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
            End If
        End If
    End If
    
End Function

Function CanNpcReflectPlayer(ByVal mapNpcNum As Long, ByVal index As Long) As Boolean
    Dim mapnum, npcNum, DistanceX, DistanceY As Long
    Dim Buffer As clsBuffer
    Dim petowner As Long

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).NPC(mapNpcNum).num <= 0 Then
        Exit Function
    End If
    
    'check if the NPC attacking us is actually our pet.
    'We don't want a rebellion on our hands now do we?
        
    mapnum = GetPlayerMap(index)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    'check if the NPC attacking us is actually our pet.
    'We don't want a rebellion on our hands now do we?
    
    If MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner = index Then Exit Function

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).StunDuration > 0 Then
        SendActionMsg mapnum, "มึน !", White, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        Exit Function
    End If

    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SNpcAttack
    Buffer.WriteLong mapNpcNum
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing

    MapNpc(mapnum).NPC(mapNpcNum).AttackTimer = GetTickCount
    
    If Map(mapnum).Moral <> MAP_MORAL_PETARENA Then
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If Not Map(GetPlayerMap(petowner)).Moral = MAP_MORAL_NONE Then
            Call PlayerMsg(petowner, "ที่นี่คือเขตปลอดภัย ! คุณไม่สามารถโจมตีใครด้วยสัตว์เลี้ยงได้.", BrightRed)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerAccess(petowner) > ADMIN_MONITOR Then
            Call PlayerMsg(petowner, "GM ไม่สามารถโจมตีสัตว์เลี้ยงผู้เล่นได้.", BrightBlue)
            Exit Function
        End If
    End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerAccess(index) > ADMIN_MONITOR Then
            Call PlayerMsg(petowner, "คุณไม่สามารถโจมตี " & GetPlayerName(index) & " ด้วยสัตว์เลี้ยงได้ !", BrightRed)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerLevel(petowner) < 10 Then
            CanNpcReflectPlayer = False
            Call PlayerMsg(petowner, "คุณมีเลเวลต่ำกว่า 10, ไม่สามารถโจมตีผู้เล่นอื่นได้ !", BrightRed)
            Exit Function
        End If
    End If
    
    If MapNpc(mapnum).NPC(mapNpcNum).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(mapNpcNum).PetData.Owner
        If GetPlayerLevel(index) < 10 Then
        CanNpcReflectPlayer = False
            Call PlayerMsg(petowner, GetPlayerName(petowner) & " มีเลเวลต่ำกว่า 10, คุณไม่สามารถโจมตีได้ !", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If npcNum > 0 Then
            DistanceX = MapNpc(mapnum).NPC(mapNpcNum).x - GetPlayerX(index)
            DistanceY = MapNpc(mapnum).NPC(mapNpcNum).y - GetPlayerY(index)
            If DistanceX < 0 Then DistanceX = DistanceX * -1
            If DistanceY < 0 Then DistanceY = DistanceY * -1
            ' Check if at same coordinates
            If DistanceX <= 1 And DistanceY <= 1 Then
                CanNpcReflectPlayer = True
            ElseIf DistanceX > 17 Or DistanceY > 14 Then
                MapNpc(mapnum).NPC(mapNpcNum).Target = 0
                MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
                MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
            End If
        End If
    End If
    
End Function


Sub NpcAttackPlayer(ByVal mapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim exp As Long
    Dim mapnum As Long
    Dim i As Long, npcNum As Long
    Dim Buffer As clsBuffer
    Dim oldX As Long, oldY As Long
    
    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).NPC(mapNpcNum).num <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Victim)
    Name = Trim$(NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Name)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' Send this packet so they can see the npc attacking
    ' Set buffer = New clsBuffer
    ' buffer.WriteLong SNpcAttack
    ' buffer.WriteLong mapNpcNum
    ' SendDataToMap mapNum, buffer.ToArray()
    ' Set buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    If GetPlayerMap(Victim) <> mapnum Then
        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapNpcNum).Target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
        MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
        'reset the targetter for the player
        Exit Sub
    End If
    
    oldX = GetPlayerX(Victim)
    oldY = GetPlayerY(Victim)
    
    ' สกิลติดตัวทำงานเมื่อถูกโจมตี? ใช้ได้แล้ว
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Victim).Spell(i) > 0 Then
                If Spell(Player(Victim).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Victim).Spell(i)).PDEF > 0 Then
                        If Spell(Player(Victim).Spell(i)).PerSkill + (Spell(Player(Victim).Spell(i)).S4 * Player(Victim).skillLV(i)) >= rand(1, 100) Then
                            If Spell(Player(Victim).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMove(Victim, i, mapNpcNum, TARGET_TYPE_NPC)
                                'Call PlayerMsg(Victim, "[Damage] : " & Player(Victim).Spell(i - 1), BrightGreen)
                            Else
                                Call CastSpellPassive(Victim, i, mapNpcNum, TARGET_TYPE_NPC)
                                'Call PlayerMsg(Victim, "[Heal] : " & Player(Victim).Spell(i - 1), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(Victim), Trim$(Spell(Player(Victim).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
                            'Call PlayerMsg(Victim, "[สกิลติดตัว] : " & Trim(Spell(Player(Victim).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
    
    If oldX <> GetPlayerX(Victim) Or oldY <> GetPlayerY(Victim) Then Exit Sub
    
    If Not CanNpcReflectPlayer(mapNpcNum, Victim) Then Exit Sub
    
    ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
    ' MapNpc(mapnum).NPC(mapNpcNum).stopRegen = True
    ' MapNpc(mapnum).NPC(mapNpcNum).stopRegenTimer = GetTickCount

    ' เริ่มการตรวจสอบสถานะ
    
    ' โจมตีไม่โดนเมื่อหายตัว
    If Player(Victim).BuffStatus(BUFF_INVISIBLE) = BUFF_INVISIBLE Then
        SendActionMsg mapnum, "หาไม่เจอ !", White, 1, (Player(Victim).x * 32), (Player(Victim).y * 32) - 16
        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapNpcNum).Target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
        MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
        Exit Sub
    End If
    
    ' find slot (ตาบอด)
    If CanNpcNoEye(npcNum) = True Then
        Player(Victim).BuffStatus(BUFF_NOEYE) = BUFF_NOEYE
        Player(Victim).BuffTime(BUFF_NOEYE) = 4 ' 4 sec
        SendActionMsg GetPlayerMap(Victim), "ตาบอด !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
        Call PlayerMsg(Victim, "คุณได้ติดสถานะตาบอด.", BrightRed)
        ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
        'Player(Victim).BuffStatus(BUFF_TOXIN) = BUFF_TOXIN
        'Player(Victim).BuffTime(BUFF_TOXIN) = 10 ' 10 sec
        'TempPlayer(Victim).stopRegen = True
        'TempPlayer(Victim).stopRegenTimer = GetTickCount + 10000
        'Call PlayerMsg(Victim, "คุณได้ติดสถานะห้ามฟื้นฟู Hp.", BrightRed)
    End If

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapNpcNum).num
        
        ' kill player
        Call KillPlayer(Victim)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " ได้ถุกฆ่าโดย " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapNpcNum).Target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
        MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
        
        Call SendPlayerData(Victim)
    Else
        ' ผู้เล่นไม่ตาย จะส่งดาเมจไป
        
        ' Kick System Npc
        If NPC(npcNum).stat(Stats.willpower) > rand(1, 100) Then
            ' set the values on index
            TempPlayer(Victim).StunDuration = 2
            TempPlayer(Victim).StunTimer = GetTickCount
            ' send it to the index
            SendStunned Victim
            SendActionMsg GetPlayerMap(Victim), "ติดมึน !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
            SendAnimation GetPlayerMap(Victim), Stun_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
            
            ' ถ้าผู้เล่นกำลังร่ายสกิล และติดมึนจะยกเลิกการร่ายสกิลทันที
            If TempPlayer(Victim).spellBuffer.Spell > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
        
        End If
        
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        Call SendAnimation(mapnum, NPC(MapNpc(GetPlayerMap(Victim)).NPC(mapNpcNum).num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
        
        ' ถ้าผู้เล่นกำลังร่ายสกิล แล้วสกิลนั้นสามารถถูกยกเลิกได้ จะสั่งให้หยุดร่ายทันที
        If TempPlayer(Victim).spellBuffer.Spell > 0 Then
            If Spell(Player(Victim).Spell(TempPlayer(Victim).spellBuffer.Spell)).CanCancle > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
        End If
               
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapNpcNum).num
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
        ' TempPlayer(Victim).stopRegen = True
        ' TempPlayer(Victim).stopRegenTimer = GetTickCount
    End If

End Sub

Sub NpcPassivePlayer(ByVal mapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim exp As Long
    Dim mapnum As Long
    Dim i As Long, npcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).NPC(mapNpcNum).num <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Victim)
    Name = Trim$(NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Name)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' Send this packet so they can see the npc attacking
    ' Set buffer = New clsBuffer
    ' buffer.WriteLong SNpcAttack
    ' buffer.WriteLong mapNpcNum
    ' SendDataToMap mapNum, buffer.ToArray()
    ' Set buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    If GetPlayerMap(Victim) <> mapnum Then
        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapNpcNum).Target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
        MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
        'reset the targetter for the player
        Exit Sub
    End If
    
    ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
    ' MapNpc(mapnum).NPC(mapNpcNum).stopRegen = True
    ' MapNpc(mapnum).NPC(mapNpcNum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapNpcNum).num
        
        ' kill player
        Call KillPlayer(Victim)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " ได้ถุกฆ่าโดย " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapNpcNum).Target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
        MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
        
        Call SendPlayerData(Victim)
    Else
        ' ผู้เล่นไม่ตาย จะส่งดาเมจไป
        
        ' Kick System Npc
        If NPC(npcNum).stat(Stats.willpower) > rand(1, 100) Then
            ' set the values on index
            TempPlayer(Victim).StunDuration = 2
            TempPlayer(Victim).StunTimer = GetTickCount
            ' send it to the index
            SendStunned Victim
            SendActionMsg GetPlayerMap(Victim), "ติดมึน !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
            SendAnimation GetPlayerMap(Victim), Stun_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
            
            ' ถ้าผู้เล่นกำลังร่ายสกิล และติดมึนจะยกเลิกการร่ายสกิลทันที
            If TempPlayer(Victim).spellBuffer.Spell > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
        
        End If
        
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        Call SendAnimation(mapnum, NPC(MapNpc(GetPlayerMap(Victim)).NPC(mapNpcNum).num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
        
        ' ถ้าผู้เล่นกำลังร่ายสกิล แล้วสกิลนั้นสามารถถูกยกเลิกได้ จะสั่งให้หยุดร่ายทันที
        If TempPlayer(Victim).spellBuffer.Spell > 0 Then
            If Spell(Player(Victim).Spell(TempPlayer(Victim).spellBuffer.Spell)).CanCancle > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
        End If
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapNpcNum).num
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
        ' TempPlayer(Victim).stopRegen = True
        ' TempPlayer(Victim).stopRegenTimer = GetTickCount
    End If

End Sub


Sub NpcReflectPlayer(ByVal mapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long, ByVal LHand As Byte)
    Dim Name As String
    Dim exp As Long
    Dim mapnum As Long, npcNum As Long
    Dim i As Long, n As Long
    Dim Buffer As clsBuffer
    Dim oldX As Long, oldY As Long

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).NPC(mapNpcNum).num <= 0 Then
        Exit Sub
    End If
    
    ' Check for weapon
    n = 0
    
    If GetPlayerEquipment(Victim, Shield) > 0 Then
        n = GetPlayerEquipment(Victim, Shield)
    ElseIf GetPlayerEquipment(Victim, Weapon) > 0 Then
        n = GetPlayerEquipment(Victim, Weapon)
    End If

    mapnum = GetPlayerMap(Victim)
    Name = Trim$(NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Name)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' Send this packet so they can see the npc attacking
    ' Set buffer = New clsBuffer
    ' buffer.WriteLong SNpcAttack
    ' buffer.WriteLong mapNpcNum
    ' SendDataToMap mapNum, buffer.ToArray()
    ' Set buffer = Nothing
    
    ' แก้ไขการแจ้งเตือนเมื่อสะท้อนไม่มีดาเมจ
    If Damage <= 0 Then
        SendActionMsg GetPlayerMap(Victim), "อ่อนหัด !", White, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
        SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
        Exit Sub
    End If
    
    oldX = GetPlayerX(Victim)
    oldY = GetPlayerY(Victim)
    
    If GetPlayerMap(Victim) <> mapnum Then
        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapNpcNum).Target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
        MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
        'reset the targetter for the player
        Exit Sub
    End If
    
    ' สกิลติดตัวทำงานเมื่อถูกโจมตี? ใช้ได้แล้ว
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Victim).Spell(i) > 0 Then
                If Spell(Player(Victim).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Victim).Spell(i)).PDEF > 0 Then
                        If Spell(Player(Victim).Spell(i)).PerSkill + (Spell(Player(Victim).Spell(i)).S4 * Player(Victim).skillLV(i)) >= rand(1, 100) Then
                            If Spell(Player(Victim).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMove(Victim, i, mapNpcNum, TARGET_TYPE_NPC)
                                'Call PlayerMsg(Victim, "[Damage] : " & Player(Victim).Spell(i - 1), BrightGreen)
                            Else
                                Call CastSpellPassive(Victim, i, mapNpcNum, TARGET_TYPE_NPC)
                                'Call PlayerMsg(Victim, "[Heal] : " & Player(Victim).Spell(i - 1), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(Victim), Trim$(Spell(Player(Victim).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
                            'Call PlayerMsg(Victim, "[สกิลติดตัว] : " & Trim(Spell(Player(Victim).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
    
    If oldX <> GetPlayerX(Victim) Or oldY <> GetPlayerY(Victim) Then Exit Sub
    
    ' เริ่มการตรวจสอบสถานะ
    
    ' โจมตีไม่โดนเมื่อหายตัว
    If Player(Victim).BuffStatus(BUFF_INVISIBLE) = BUFF_INVISIBLE Then
        SendActionMsg mapnum, "หาไม่เจอ !", White, 1, (Player(Victim).x * 32), (Player(Victim).y * 32) - 16
        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapNpcNum).Target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
        MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
        Exit Sub
    End If
    
    ' find slot (ตาบอด)
    If CanNpcNoEye(npcNum) = True Then
        Player(Victim).BuffStatus(BUFF_NOEYE) = BUFF_NOEYE
        Player(Victim).BuffTime(BUFF_NOEYE) = 4 ' 4 sec
        SendActionMsg GetPlayerMap(Victim), "ตาบอด !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
        Call PlayerMsg(Victim, "คุณได้ติดสถานะตาบอด.", BrightRed)
        ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
        'Player(Victim).BuffStatus(BUFF_TOXIN) = BUFF_TOXIN
        'Player(Victim).BuffTime(BUFF_TOXIN) = 10 ' 10 sec
        'TempPlayer(Victim).stopRegen = True
        'TempPlayer(Victim).stopRegenTimer = GetTickCount + 10000
        'Call PlayerMsg(Victim, "คุณได้ติดสถานะห้ามฟื้นฟู Hp.", BrightRed)
    End If
    
    If Not CanNpcReflectPlayer(mapNpcNum, Victim) Then Exit Sub
    
    ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
    ' MapNpc(mapnum).NPC(mapNpcNum).stopRegen = True
    ' MapNpc(mapnum).NPC(mapNpcNum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapNpcNum).num
        
        ' kill player
        Call KillPlayer(Victim)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " ได้ถุกฆ่าโดย " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapNpcNum).Target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
        MapNpc(mapnum).NPC(mapNpcNum).GetDamage = 0
        
        Call SendPlayerData(Victim)
    Else
        ' ผู้เล่นไม่ตาย จะส่งดาเมจไป
        
        ' Kick System
        If n > 0 Then
            If LHand = 0 Then
                If Item(GetPlayerEquipment(Victim, Weapon)).Kick > 0 Then
                    If Item(GetPlayerEquipment(Victim, Weapon)).Kick > rand(1, 100) Then
                        ' set the values on index
                        TempPlayer(Victim).StunDuration = 2
                        TempPlayer(Victim).StunTimer = GetTickCount
                        ' send it to the index
                        SendStunned Victim
                        SendActionMsg GetPlayerMap(Victim), "ติดมึน !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
                        SendAnimation GetPlayerMap(Victim), Stun_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                        
            ' ถ้าผู้เล่นกำลังร่ายสกิล และติดมึนจะยกเลิกการร่ายสกิลทันที
            If TempPlayer(Victim).spellBuffer.Spell > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
            
                    End If
                End If
            Else
                If Item(GetPlayerEquipment(Victim, Shield)).Kick > 0 Then
                    If Item(GetPlayerEquipment(Victim, Shield)).Kick > rand(1, 100) Then
                        ' set the values on index
                        TempPlayer(Victim).StunDuration = 2
                        TempPlayer(Victim).StunTimer = GetTickCount
                        ' send it to the index
                        SendStunned Victim
                        SendActionMsg GetPlayerMap(Victim), "ติดมึน !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
                        SendAnimation GetPlayerMap(Victim), Stun_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                        
            ' ถ้าผู้เล่นกำลังร่ายสกิล และติดมึนจะยกเลิกการร่ายสกิลทันที
            If TempPlayer(Victim).spellBuffer.Spell > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
            
                    End If
                End If
            End If
        End If
              
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        Call SendAnimation(mapnum, NPC(MapNpc(GetPlayerMap(Victim)).NPC(mapNpcNum).num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
        
        ' ถ้าผู้เล่นกำลังร่ายสกิล แล้วสกิลนั้นสามารถถูกยกเลิกได้ จะสั่งให้หยุดร่ายทันที
        If TempPlayer(Victim).spellBuffer.Spell > 0 Then
            If Spell(Player(Victim).Spell(TempPlayer(Victim).spellBuffer.Spell)).CanCancle > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
        End If
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapNpcNum).num
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
    End If

End Sub


Sub NpcSpellPlayer(ByVal mapNpcNum As Long, ByVal Victim As Long, SpellSlotNum As Long)
    Dim mapnum As Long
    Dim i As Long
    Dim n As Long
    Dim spellnum As Long
    Dim Buffer As clsBuffer
    Dim InitDamage As Long
    Dim Damage As Long, Vital As Long
    Dim MaxHeals As Long
    Dim s(1 To 2) As Long
    Dim r(1 To 2) As Long
    
    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    If Not Victim > 0 Then Exit Sub

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).NPC(mapNpcNum).num <= 0 Then
        Exit Sub
    End If
    
    If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub
    If MapNpc(GetPlayerMap(Victim)).NPC(mapNpcNum).PetData.Owner = Victim Then Exit Sub

    ' The Variables
    mapnum = GetPlayerMap(Victim)
    spellnum = NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Spell(SpellSlotNum)
    
    ' CoolDown Time
    If MapNpc(mapnum).NPC(mapNpcNum).SpellTimer(SpellSlotNum) > GetTickCount Then Exit Sub
    If GetPlayerVital(Victim, HP) <= 0 Then Exit Sub
    
    ' Send this packet so they can see the person attacking [before cd]
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong mapNpcNum
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
    
        'Vital = Spell(spellnum).Vital
        
        'If Spell(spellnum).PhysicalDmg > 0 Then
        '    Vital = Vital + ((rand(NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Damage / 2, NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Damage) * Spell(spellnum).ATKPer) / 100)
        'End If
        
        'If Spell(spellnum).MagicDmg > 0 Then
        '    Vital = Vital + ((rand(NPC(MapNpc(mapnum).NPC(mapNpcNum).num).MATK / 2, NPC(MapNpc(mapnum).NPC(mapNpcNum).num).MATK) * Spell(spellnum).MagicPer) / 100)
        'End If
        
        ' New Vital
        Vital = Spell(spellnum).Vital + (Spell(spellnum).Vital * ((Spell(spellnum).S1 * (NPC(MapNpc(mapnum).NPC(mapNpcNum).num).stat(Stats.Strength)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetNpcDamage(MapNpc(mapnum).NPC(mapNpcNum).num) / 2, GetNpcDamage(MapNpc(mapnum).NPC(mapNpcNum).num))
            r(1) = (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (NPC(MapNpc(mapnum).NPC(mapNpcNum).num).stat(Stats.Strength) / 100)))
            'Vital = Vital + R(1)
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(NPC(MapNpc(mapnum).NPC(mapNpcNum).num).MATK / 2, NPC(MapNpc(mapnum).NPC(mapNpcNum).num).MATK)
            r(2) = (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (NPC(MapNpc(mapnum).NPC(mapNpcNum).num).stat(Stats.Strength) / 100)))
            'Vital = Vital + R(2)
        End If
    
    
    ' Spell Types
        Select Case Spell(spellnum).Type
            ' AOE Healing Spells
            Case SPELL_TYPE_HEALHP
            ' Make sure an npc waits for the spell to cooldown
            MaxHeals = 9999 + NPC(MapNpc(mapnum).NPC(mapNpcNum).num).stat(Stats.intelligence) * 10
            ' ถ้า npc ฮีลบ่อยเกินไป จะทำให้หยุดฮีล?
            'If MapNpc(mapnum).NPC(mapNpcNum).Heals >= MaxHeals Then Exit Sub
                If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= NPC(MapNpc(mapnum).NPC(mapNpcNum).num).HP * 0.5 Then
                    If Spell(spellnum).IsAoE Then
                        For i = 1 To MAX_MAP_NPCS
                            If MapNpc(mapnum).NPC(i).num > 0 Then
                                If MapNpc(mapnum).NPC(i).Vital(Vitals.HP) > 0 Then
                                    If isInRange(Spell(spellnum).AoE, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                        InitDamage = Vital + s(1) + r(1) + s(2) + r(2)
                    
                                        MapNpc(mapnum).NPC(i).Vital(Vitals.HP) = MapNpc(mapnum).NPC(i).Vital(Vitals.HP) + InitDamage
                                        SendActionMsg mapnum, "+" & InitDamage, BrightGreen, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32)
                                        Call SendAnimation(mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
                    
                                        If MapNpc(mapnum).NPC(i).Vital(Vitals.HP) >= NPC(MapNpc(mapnum).NPC(i).num).HP Then
                                            MapNpc(mapnum).NPC(i).Vital(Vitals.HP) = NPC(MapNpc(mapnum).NPC(i).num).HP
                                        End If
                                        
                                        ' แก้บัค ไม่อัพเดทเลือด npc ตอน Heal
                                        SendMapNpcVitals mapnum, mapNpcNum
                    
                                        MapNpc(mapnum).NPC(mapNpcNum).Heals = MapNpc(mapnum).NPC(mapNpcNum).Heals + 1
                    
                                        MapNpc(mapnum).NPC(mapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(spellnum).CDTime * 1000
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                        
                        ' msg spell
                        SendActionMsg mapnum, Trim(Spell(spellnum).Name), BrightGreen, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                    
                    Else
                    ' Non AOE Healing Spells
                        InitDamage = Vital + s(1) + r(1) + s(2) + r(2)
                    
                        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) + InitDamage
                        SendActionMsg mapnum, "+" & InitDamage, BrightGreen, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
                        Call SendAnimation(mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
                    
                        If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) >= NPC(MapNpc(mapnum).NPC(mapNpcNum).num).HP Then
                            MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = NPC(MapNpc(mapnum).NPC(mapNpcNum).num).HP
                        End If
                        
                        ' แก้บัค ไม่อัพเดทเลือด npc ตอน Heal
                        SendMapNpcVitals mapnum, mapNpcNum
                    
                        MapNpc(mapnum).NPC(mapNpcNum).Heals = MapNpc(mapnum).NPC(mapNpcNum).Heals + 1
                    
                        MapNpc(mapnum).NPC(mapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(spellnum).CDTime * 1000
                        
                        ' msg spell
                        SendActionMsg mapnum, Trim(Spell(spellnum).Name), BrightGreen, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                        Exit Sub
                    End If
                End If
                
            ' AOE Damaging Spells
            Case SPELL_TYPE_DAMAGEHP
            ' Make sure an npc waits for the spell to cooldown
                If Spell(spellnum).IsAoE Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = mapnum Then
                                If isInRange(Spell(spellnum).AoE, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, GetPlayerX(i), GetPlayerY(i)) Then
                                    InitDamage = Vital + s(1) + r(1) + s(2) + r(2)
                                    
                                    ' fixed damage
                                    If Spell(spellnum).CanMove > 0 Then
                                        Damage = InitDamage - GetPlayerDef(i)
                                    Else
                                        Damage = InitDamage - GetPlayerMDEF(i)
                                    End If
                                    
                                    If Not CanPlayerAbsorbMagic(i) Then
                                        If Damage <= 0 Then
                                            SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, mapNpcNum
                                            SendActionMsg GetPlayerMap(i), "แย่มาก !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                                            Exit Sub
                                        Else
                                            ' fixed damage
                                            If Spell(spellnum).CanMove > 0 Then
                                                NpcAttackPlayer mapNpcNum, i, Damage
                                            Else
                                                NpcPassivePlayer mapNpcNum, i, Damage
                                            End If
                                            
                                            SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, mapNpcNum
                                            MapNpc(mapnum).NPC(mapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(spellnum).CDTime * 1000
                                            Exit Sub
                                        End If
                                    Else
                                        ' Absorb
                                        MapNpc(mapnum).NPC(mapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(spellnum).CDTime * 1000
                                        SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                        SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    ' msg spell
                    SendActionMsg mapnum, Trim(Spell(spellnum).Name), BrightGreen, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                
                ' Non AoE Damaging Spells
                Else
                    If isInRange(Spell(spellnum).Range, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, GetPlayerX(Victim), GetPlayerY(Victim)) Then
                        InitDamage = Vital + s(1) + r(1) + s(2) + r(2)

                        ' fixed damage
                        If Spell(spellnum).CanMove > 0 Then
                            Damage = InitDamage - GetPlayerDef(Victim)
                        Else
                            Damage = InitDamage - GetPlayerMDEF(Victim)
                        End If
                        
                        ' msg spell
                        SendActionMsg mapnum, Trim(Spell(spellnum).Name), BrightGreen, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                        
                        If Not CanPlayerAbsorbMagic(Victim) Then
                            If Damage <= 0 Then
                                SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Victim
                                SendActionMsg GetPlayerMap(Victim), "แย่มาก !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
                                Exit Sub
                            Else
                                ' fixed damage
                                If Spell(spellnum).CanMove > 0 Then
                                    NpcAttackPlayer mapNpcNum, Victim, Damage
                                Else
                                    NpcPassivePlayer mapNpcNum, Victim, Damage
                                End If
                                
                                SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Victim
                                MapNpc(mapnum).NPC(mapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(spellnum).CDTime * 1000
                                Exit Sub
                            End If
                        Else
                            ' Absorb
                            MapNpc(mapnum).NPC(mapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(spellnum).CDTime * 1000
                            SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                            SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 16
                        End If
                    End If
                End If
                
                Case SPELL_TYPE_WARP
                
                'Call PlayerMsg(Victim, "1", BrightRed)
                    If IsPlaying(Victim) Then
                        If GetPlayerMap(Victim) = mapnum Then
                            If isInRange(Spell(spellnum).Range, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, GetPlayerX(Victim), GetPlayerY(Victim)) Then
                                ' Make sure an npc waits for the spell to cooldown
                
                                    InitDamage = Vital + s(1) + r(1) + s(2) + r(2)
                                    'Call PlayerMsg(Victim, "2", BrightRed)
                        
                                    Select Case Player(Victim).Dir
                        
                                    Case DIR_UP
                                        If Player(Victim).y + 1 < Map(mapnum).MaxY Then
                                            'Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                            Exit Sub
                                        End If
                            
                                        If Not Map(mapnum).Tile(Player(Victim).x, Player(Victim).y - 1).Type = TILE_TYPE_WALKABLE Then
                                            SendActionMsg mapnum, "ปิดบัง !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                                            Exit Sub
                                        Else
                                        
                                        If Not ((MapNpc(mapnum).NPC(mapNpcNum).x = Player(Victim).x) And (MapNpc(mapnum).NPC(mapNpcNum).y = Player(Victim).y - 1)) Then
                                            NpcWarp mapNpcNum, Player(Victim).x, Player(Victim).y - 1, DIR_UP, mapnum
                                        End If
                            
                                        End If
                                    Case DIR_DOWN
                                        If Player(Victim).y - 1 > 1 Then
                                            'Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                            Exit Sub
                                        End If
                            
                                        If Not Map(mapnum).Tile(Player(Victim).x, Player(Victim).y + 1).Type = TILE_TYPE_WALKABLE Then
                                            SendActionMsg mapnum, "ปิดบัง !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                                            Exit Sub
                                        Else
                                
                                        If Not ((MapNpc(mapnum).NPC(mapNpcNum).x = Player(Victim).x) And (MapNpc(mapnum).NPC(mapNpcNum).y = Player(Victim).y + 1)) Then
                                            NpcWarp mapNpcNum, Player(Victim).x, Player(Victim).y + 1, DIR_DOWN, mapnum
                                        End If
                                
                                        End If
                                    Case DIR_LEFT
                                        If Player(Victim).x + 1 > Map(mapnum).MaxX Then
                                            'Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                            Exit Sub
                                        End If
                        
                                        If Not Map(mapnum).Tile(Player(Victim).x + 1, Player(Victim).y).Type = TILE_TYPE_WALKABLE Then
                                            SendActionMsg mapnum, "ปิดบัง !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                                            Exit Sub
                                        Else
                                
                                        If Not ((MapNpc(mapnum).NPC(mapNpcNum).x = Player(Victim).x + 1) And (MapNpc(mapnum).NPC(mapNpcNum).y = Player(Victim).y)) Then
                                            NpcWarp mapNpcNum, Player(Victim).x + 1, Player(Victim).y, DIR_LEFT, mapnum
                                        End If
                                
                                        End If
                                    Case DIR_RIGHT
                                        If Player(Victim).x - 1 < 1 Then
                                            'Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                            Exit Sub
                                        End If
                            
                                        If Not Map(mapnum).Tile(Player(Victim).x - 1, Player(Victim).y).Type = TILE_TYPE_WALKABLE Then
                                            SendActionMsg mapnum, "ปิดบัง !", BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                                            Exit Sub
                                        Else
                                
                                        If Not ((MapNpc(mapnum).NPC(mapNpcNum).x = Player(Victim).x - 1) And (MapNpc(mapnum).NPC(mapNpcNum).y = Player(Victim).y)) Then
                                            NpcWarp mapNpcNum, Player(Victim).x - 1, Player(Victim).y, DIR_RIGHT, mapnum
                                        End If
                                
                                        End If
                                    End Select
                        
                                    'Call PlayerMsg(Victim, "3", BrightRed)
                        
                                    ' fixed damage
                                    If Spell(spellnum).CanMove > 0 Then
                                        Damage = InitDamage - GetPlayerDef(Victim)
                                        'Call PlayerMsg(Victim, "Damage : " & InitDamage & " - " & GetPlayerDef(Victim) & " = " & InitDamage - GetPlayerDef(Victim), Yellow)
                                        'Call PlayerMsg(Victim, "ATK : " & S(1), Yellow)
                                        'Call PlayerMsg(Victim, "MATK : " & S(2), Yellow)
                                        'Call PlayerMsg(Victim, "VATK : " & (S(1) * Spell(spellnum).ATKPer / 100) + (S(1) * (Spell(spellnum).S2 * (NPC(MapNpc(mapnum).NPC(mapNpcNum).num).stat(Stats.Strength) / 100))), Yellow)
                                        'Call PlayerMsg(Victim, "CAL : " & InitDamage & " - " & GetPlayerDef(Victim) & " = " & Damage, Yellow)
                                    Else
                                        Damage = InitDamage - GetPlayerMDEF(Victim)
                                        'Call PlayerMsg(Victim, "Vital : " & InitDamage & " - " & GetPlayerMDEF(Victim) & " = " & InitDamage - GetPlayerMDEF(Victim), Yellow)
                                        'Call PlayerMsg(Victim, "ATK : " & S(1), Yellow)
                                        'Call PlayerMsg(Victim, "MATK : " & S(2), Yellow)
                                        'Call PlayerMsg(Victim, "VATK : " & (S(1) * Spell(spellnum).ATKPer / 100) + (S(1) * (Spell(spellnum).S2 * (NPC(MapNpc(mapnum).NPC(mapNpcNum).num).stat(Stats.Strength) / 100))), Yellow)
                                        'Call PlayerMsg(Victim, "CAL : " & InitDamage & " - " & GetPlayerMDEF(Victim) & " = " & Damage, Yellow)
                                    End If
                                           
                                    ' msg spell
                                    SendActionMsg mapnum, Trim(Spell(spellnum).Name), BrightGreen, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32) - 16
                                    
                                    If Not CanPlayerAbsorbMagic(Victim) Then
                                        If Damage <= 0 Then
                                            SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Victim
                                            SendActionMsg GetPlayerMap(Victim), "แย่มาก !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
                                            Exit Sub
                                        Else
                                            ' fixed damage
                                            If Spell(spellnum).CanMove > 0 Then
                                                NpcAttackPlayer mapNpcNum, Victim, Damage
                                            Else
                                                NpcPassivePlayer mapNpcNum, Victim, Damage
                                            End If
                                
                                            SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Victim
                                            MapNpc(mapnum).NPC(mapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(spellnum).CDTime * 1000
                                            Exit Sub
                                        End If
                                    Else
                                        ' Absorb
                                        MapNpc(mapnum).NPC(mapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(spellnum).CDTime * 1000
                                        SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                                        SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 16
                                    End If

                            End If
                        End If
                    End If
                        
            End Select
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long)
Dim BlockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim Damage As Long
Dim DEFP As Long, NDEF As Boolean, NDEFLHAND As Boolean
    
    Damage = 0
    NDEF = False
    
    mapnum = GetPlayerMap(Attacker)
           
    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, Victim) Then
           
        ' check if NPC can avoid the attack
        If CanPlayerDodge(Victim) And Not CanPlayerCrit(Attacker) Then
            SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
            SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 16
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        DEFP = GetPlayerDef(Victim)
        
        ' ระบบเจาะเกราะ
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If Item(GetPlayerEquipment(Attacker, Weapon)).NDEF > 0 Then
                NDEF = True
            End If
        End If
        
        ' x1.2 Critical ! +ระบบเพิ่มความแรงคริติคอล
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * GetPlayerCritDamage(Attacker, False)
            SendActionMsg mapnum, "คริติคอล !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 16
            SendAnimation mapnum, CRIT_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
        Else
            ' ระบบเจาะเกราะ
            If NDEF = True Then
                Damage = Damage - (DEFP - ((DEFP * Item(GetPlayerEquipment(Attacker, Weapon)).NDEF) / 100))
            Else
                Damage = Damage - DEFP
            End If
        End If
        
        ' ระบบสะท้อน
        If CanPlayerBlock(Victim) Then
             If Not CanPlayerDodge(Attacker) Then
                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                SendActionMsg mapnum, "สะท้อน !", BrightCyan, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 16
                Call PlayerMsg(Attacker, "ผู้เล่น " & Trim(Player(Victim).Name) & " ได้สะท้อนดาเมจกลับ.", BrightCyan)
                Call PlayerReflectPlayer(Victim, Attacker, Damage * (GetPlayerReflectDMG(Attacker) / 100), 0)
                Exit Sub
            Else
                SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, Attacker
                SendActionMsg mapnum, "หลบสะท้อน !", BrightRed, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) - 16
            End If
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, Victim, Damage)
            
            ' ระบบ Vampire
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If Item(GetPlayerEquipment(Attacker, Weapon)).Vampire > 0 Then
            
                ' แก้ไขบัคดูดเลือดเกิน !!
                If GetPlayerMaxVital(Attacker, HP) > Player(Attacker).Vital(Vitals.HP) + Int((Damage * (Item(GetPlayerEquipment(Attacker, Weapon)).Vampire / 100))) Then
                    Player(Attacker).Vital(Vitals.HP) = Player(Attacker).Vital(Vitals.HP) + Int((Damage * (Item(GetPlayerEquipment(Attacker, Weapon)).Vampire / 100)))
                Else
                    Player(Attacker).Vital(Vitals.HP) = GetPlayerMaxVital(Attacker, HP)
                End If
                
                ' send vitals to party if in one
                If TempPlayer(Attacker).inParty > 0 Then SendPartyVitals TempPlayer(Attacker).inParty, Attacker
                SendActionMsg GetPlayerMap(Attacker), "+" & Int((Damage * (Item(GetPlayerEquipment(Attacker, Weapon)).Vampire / 100))), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Attacker) * 32, GetPlayerY(Attacker) * 32
                SendAnimation mapnum, Vampire_ANIM, 0, 0, TARGET_TYPE_PLAYER, Attacker
                SendVital Attacker, HP
            End If
        End If
        
        Else
            SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
            SendActionMsg mapnum, "อ่อนหัด !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 16
            ' Call PlayerMsg(Attacker, "การโจมตีของคุณไร้ผล.", BrightRed)
        End If
    
    End If
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayerLHand(ByVal Attacker As Long, ByVal Victim As Long)
Dim BlockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim Damage As Long
Dim DEFP As Long, NDEF As Boolean, NDEFLHAND As Boolean
    
    Damage = 0
    NDEF = False

    ' Can we attack the npc?
    If CanPlayerAttackPlayerLHand(Attacker, Victim) Then
    
        mapnum = GetPlayerMap(Attacker)
        
        ' check if NPC can avoid the attack
        If CanPlayerDodge(Victim) And Not CanPlayerCrit(Attacker) Then
            SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
            SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 16
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamageLHand(Attacker)
        DEFP = GetPlayerDef(Victim)
        
        ' ระบบเจาะเกราะ
        If GetPlayerEquipment(Attacker, Shield) > 0 Then
            If Item(GetPlayerEquipment(Attacker, Shield)).NDEF > 0 Then
                NDEF = True
            End If
        End If
        
        ' x1.2 Critical ! +ระบบเพิ่มความแรงคริติคอล
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * GetPlayerCritDamage(Attacker, True)
            SendActionMsg mapnum, "คริติคอล !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 16
            SendAnimation mapnum, CRIT_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
        Else
            ' ระบบเจาะเกราะ
            If NDEF = True Then
                Damage = Damage - (DEFP - ((DEFP * Item(GetPlayerEquipment(Attacker, Shield)).NDEF) / 100))
            Else
                Damage = Damage - DEFP
            End If
        End If
        
        ' ระบบสะท้อน
        If CanPlayerBlock(Victim) Then
             If Not CanPlayerDodge(Attacker) Then
                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                SendActionMsg mapnum, "สะท้อน !", BrightCyan, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 16
                Call PlayerMsg(Attacker, "ผู้เล่น " & Trim(Player(Victim).Name) & " ได้สะท้อนดาเมจกลับ.", BrightCyan)
                Call PlayerReflectPlayer(Victim, Attacker, Damage * (GetPlayerReflectDMG(Attacker) / 100), 0)
                Exit Sub
            Else
                SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, Attacker
                SendActionMsg mapnum, "หลบสะท้อน !", BrightRed, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) - 16
            End If
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayerLHand(Attacker, Victim, Damage)
            
            ' ระบบ Vampire
        If GetPlayerEquipment(Attacker, Shield) > 0 Then
            If Item(GetPlayerEquipment(Attacker, Shield)).Vampire > 0 Then
            
                ' แก้ไขบัคดูดเลือดเกิน !!
                If GetPlayerMaxVital(Attacker, HP) > Player(Attacker).Vital(Vitals.HP) + Int((Damage * (Item(GetPlayerEquipment(Attacker, Shield)).Vampire / 100))) Then
                    Player(Attacker).Vital(Vitals.HP) = Player(Attacker).Vital(Vitals.HP) + Int((Damage * (Item(GetPlayerEquipment(Attacker, Shield)).Vampire / 100)))
                Else
                    Player(Attacker).Vital(Vitals.HP) = GetPlayerMaxVital(Attacker, HP)
                End If
                
                ' send vitals to party if in one
                If TempPlayer(Attacker).inParty > 0 Then SendPartyVitals TempPlayer(Attacker).inParty, Attacker
                SendActionMsg GetPlayerMap(Attacker), "+" & Int((Damage * (Item(GetPlayerEquipment(Attacker, Shield)).Vampire / 100))), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Attacker) * 32, GetPlayerY(Attacker) * 32
                SendAnimation mapnum, Vampire_ANIM, 0, 0, TARGET_TYPE_PLAYER, Attacker
                SendVital Attacker, HP
            End If
        End If
        
        Else
            SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
            SendActionMsg mapnum, "อ่อนหัด !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 16
            ' Call PlayerMsg(Attacker, "การโจมตีของคุณไร้ผล.", BrightRed)
        End If
    
    End If
End Sub

' projectiles fixed
Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean

    If Not IsSpell And Not IsProjectile Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(Attacker).AttackTimer + (2000 + ((GetPlayerStat(Attacker, Stats.Agility) * 5))) Then Exit Function
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

    ' อยู่ในปาร์ตี้เดียวกันไม่สามารถโจมตีกันเองได้
    If TempPlayer(Attacker).inParty > 0 Then
        If TempPlayer(Attacker).inParty = TempPlayer(Victim).inParty Then
            'Call PlayerMsg(Attacker, "ไม่สามารถโจมตีผู้เล่นที่อยู่ในปาร์ตี้เดียวกันได้ !", BrightRed)
            Exit Function
        End If
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "ที่นี่คือเขตปลอดภัย !", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function
    
    ' ตรวจสอบการมึน
    If TempPlayer(Attacker).StunDuration > 0 Then
        'Call PlayerMsg(Attacker, "คุณกำลังมึนงง.", BrightRed)
        'SendActionMsg GetPlayerMap(Attacker), "มึน !", BrightRed, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        Exit Function
    End If

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "GM ไม่สามารถโจมตีผู้เล่นอื่นได้.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "คุณไม่สามารถโจมตีผู้เล่น " & GetPlayerName(Victim) & " !", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "คุณมีเลเวลต่ำกว่า 10, ไม่สามารถโจมตีผู้เล่นอื่นได้ !", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " มีเลเวลต่ำกว่า 10, คุณไม่สามารถโจมตีได้ !", BrightRed)
        Exit Function
    End If

    TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(Attacker).Target = Victim
    SendTarget Attacker
    CanPlayerAttackPlayer = True

End Function

' projectiles fixed
Function CanPlayerAttackPlayerLHand(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function
    
    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function
    
    If IsProjectile = True Then Exit Function
    
    If Not IsSpell Then
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
                'Exit Function
        End Select
    End If
    
    ' อยู่ในปาร์ตี้เดียวกันไม่สามารถโจมตีกันเองได้
    If TempPlayer(Attacker).inParty > 0 Then
        If TempPlayer(Attacker).inParty = TempPlayer(Victim).inParty Then
            'Call PlayerMsg(Attacker, "ไม่สามารถโจมตีผู้เล่นที่อยู่ในปาร์ตี้เดียวกันได้ !", BrightRed)
            Exit Function
        End If
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "ที่นี่คือเขตปลอดภัย !", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function

    ' ตรวจสอบการมึน
    If TempPlayer(Attacker).StunDuration > 0 Then
        'Call PlayerMsg(Attacker, "คุณกำลังมึนงง.", BrightRed)
        'SendActionMsg GetPlayerMap(Attacker), "มึน !", BrightRed, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        Exit Function
    End If

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "GM ไม่สามารถโจมตีผู้เล่นอื่นได้.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "คุณไม่สามารถโจมตีผู้เล่น " & GetPlayerName(Victim) & " !", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "คุณมีเลเวลต่ำกว่า 10, ไม่สามารถโจมตีผู้เล่นอื่นได้ !", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " มีเลเวลต่ำกว่า 10, คุณไม่สามารถโจมตีได้ !", BrightRed)
        Exit Function
    End If

    TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(Attacker).Target = Victim
    SendTarget Attacker
    CanPlayerAttackPlayerLHand = True

End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim oldX As Long, oldY As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    oldX = GetPlayerX(Victim)
    oldY = GetPlayerY(Victim)
    
    ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
    ' TempPlayer(Attacker).stopRegen = True
    ' TempPlayer(Attacker).stopRegenTimer = GetTickCount

    ' สกิลติดตัวทำงานเมื่อโจมตี? ใช้ได้แล้ว
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Spell(Player(Attacker).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Attacker).Spell(i)).PATK > 0 Then
                        If Spell(Player(Attacker).Spell(i)).PerSkill >= rand(1, 100) Then
                            If Spell(Player(Attacker).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMove(Attacker, i, Victim, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(Attacker, "[Damage] : " & Player(Attacker).Spell(i - 1), BrightGreen)
                            Else
                                Call CastSpellPassive(Attacker, i, Victim, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(Attacker, "[Heal] : " & Player(Attacker).Spell(i - 1), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(Attacker), Trim$(Spell(Player(Attacker).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16
                            'Call PlayerMsg(Attacker, "[สกิลติดตัว] : " & Trim(Spell(Player(Attacker).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
        
        ' สกิลติดตัวทำงานเมื่อถูกโจมตี? ใช้ได้แล้ว
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Victim).Spell(i) > 0 Then
                If Spell(Player(Victim).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Victim).Spell(i)).PDEF > 0 Then
                        If Spell(Player(Victim).Spell(i)).PerSkill + (Spell(Player(Victim).Spell(i)).S4 * Player(Victim).skillLV(i)) >= rand(1, 100) Then
                            If Spell(Player(Victim).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMove(Victim, i, Attacker, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(victim, "[Damage] : " & Player(victim).Spell(i - 1), BrightGreen)
                            Else
                                Call CastSpellPassive(Victim, i, Attacker, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(victim, "[Heal] : " & Player(victim).Spell(i - 1), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(victim), Trim$(Spell(Player(victim).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32) + 16
                            'Call PlayerMsg(victim, "[สกิลติดตัว] : " & Trim(Spell(Player(victim).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
        
        If oldX <> GetPlayerX(Victim) Or oldY <> GetPlayerY(Victim) Then Exit Sub
        
        If Player(Victim).Vital(Vitals.HP) <= 0 Then Exit Sub

    ' ถ้าดาเมจที่โจมตีมากกว่าเลือดศัตรูที่เหลือ จะทำให้ตาย
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
    
        If CanPlayerLHand(Attacker) Then
            SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32) + 8, (GetPlayerY(Victim) * 32)
        Else
            SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        End If
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
        
        ' Killer
        Select Case Player(Attacker).Killer
        Case 0: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 200
        Call SendAnimation(GetPlayerMap(Attacker), 200, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้เริ่มฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 1: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 201
        Call SendAnimation(GetPlayerMap(Attacker), 201, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่าฝึกหัด.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 2: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 202
        Call SendAnimation(GetPlayerMap(Attacker), 202, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 3: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 203
        Call SendAnimation(GetPlayerMap(Attacker), 203, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ชำนาญอาวุธ.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 4: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 204
        Call SendAnimation(GetPlayerMap(Attacker), 204, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่าไร้เงา.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 5: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 205
        Call SendAnimation(GetPlayerMap(Attacker), 205, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ปีศาจ.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 6: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 206
        Call SendAnimation(GetPlayerMap(Attacker), 206, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ปีศาจนักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 7: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 207
        Call SendAnimation(GetPlayerMap(Attacker), 207, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้พิฆาต.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 8: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 208
        Call SendAnimation(GetPlayerMap(Attacker), 208, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ตำนานนักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 9 To 13: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 209
        Call SendAnimation(GetPlayerMap(Attacker), 209, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้กระหายเลือด.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case Else: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 210
        Call SendAnimation(GetPlayerMap(Attacker), 210, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ราชาผู้อมตะ.", BrightRed)
        Call GlobalMsg("ใครก็ได้ฆ่า " & GetPlayerName(Attacker) & " ที.", Yellow)
        Player(Attacker).Killer = 15
        
        End Select
        
        ' แก้ไขบัค ไม่ตรวจสอบอาวุธเมื่อโจมตีผู้เล่น
        If n > 0 Then
            ' send anim
            Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Weapon)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
        Else
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(GetPlayerMap(Victim), PUNCH_ANIM, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " ได้ถูกฆ่าโดย " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        
        exp = (GetPlayerNextLevel(Victim) * 0.05)

        ' Make sure we dont get less then 0
        If GetPlayerExp(Victim) < exp Then
            exp = 0
        End If

        ' ตรวจสอบเงื่อนไขถ้า exp = 0 จะไม่เสียค่าประสบการณ์
        If exp = 0 Then
            Call PlayerMsg(Victim, "คุณไม่เหลือค่าประสบการณ์จากการตาย.", BrightRed)
            Call PlayerMsg(Attacker, "คุณไม่ได้รับ Exp จากการสังหาร (เพราะศัตรูไม่เหลือค่า Exp).", BrightBlue)
            Call SetPlayerExp(Victim, exp)
            SendEXP Victim
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - exp)
            Call PlayerMsg(Victim, "คุณได้สูญเสีย exp " & exp & " จากการตาย.", BrightRed)
            SendEXP Victim
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                    Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                Else
                    Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                    Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                    Call SetPlayerExp(Attacker, 1)
                    SendEXP Attacker
                End If
            Else
                ' no party - keep exp for self
                If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                    GivePlayerEXP Attacker, exp
                    Call PlayerMsg(Attacker, "คุณได้รับ " & exp & " Exp จากการสังหารผู้เล่น " & Player(Victim).Name, Yellow)
                Else
                    ' GivePlayerEXP Attacker, 0
                    Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                    Call SetPlayerExp(Attacker, 1)
                    SendEXP Attacker
                End If
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " ถูกตั้งค่าหัวว่าสังหารผู้เล่นอื่น !!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " ได้โดนปิดตำนานการสังหารผู้เล่น !!", BrightRed)
        End If
        
        Call CheckTasks(Attacker, QUEST_TYPE_GOKILL, Victim)
        Call OnDeath(Victim)
    Else
        ' ผู้เล่นไม่ตายและยังได้รับดาเมจ
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
        
        ' ถ้าผู้เล่นกำลังร่ายสกิล แล้วสกิลนั้นสามารถถูกยกเลิกได้ จะสั่งให้หยุดร่ายทันที
        If TempPlayer(Victim).spellBuffer.Spell > 0 Then
            If Spell(Player(Victim).Spell(TempPlayer(Victim).spellBuffer.Spell)).CanCancle > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
        End If
        
        ' แก้ไขบัค ไม่ตรวจสอบอาวุธเมื่อโจมตีผู้เล่น
        If n > 0 Then
            ' send anim
            Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Weapon)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
        Else
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(GetPlayerMap(Victim), PUNCH_ANIM, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        If CanPlayerLHand(Attacker) Then
            SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32) + 8, (GetPlayerY(Victim) * 32)
        Else
            SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        End If
        
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' Kick System
        If n > 0 Then
            If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > 0 Then
                If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > rand(1, 100) Then
                    ' set the values on index
                    TempPlayer(Victim).StunDuration = 2
                    TempPlayer(Victim).StunTimer = GetTickCount
                    ' send it to the index
                    SendStunned Victim
                    SendActionMsg GetPlayerMap(Victim), "ติดมึน !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
                    SendAnimation GetPlayerMap(Victim), Stun_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                    
            ' ถ้าผู้เล่นกำลังร่ายสกิล และติดมึนจะยกเลิกการร่ายสกิลทันที
            If TempPlayer(Victim).spellBuffer.Spell > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
            
                End If
            End If
        End If
        
        ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
        ' TempPlayer(Victim).stopRegen = True
        ' TempPlayer(Victim).stopRegenTimer = GetTickCount
        
        
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

Sub PlayerPassivePlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0, Optional ByVal spellAnim As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
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
    
    ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
    ' TempPlayer(Attacker).stopRegen = True
    ' TempPlayer(Attacker).stopRegenTimer = GetTickCount
    
    If Player(Victim).Vital(Vitals.HP) <= 0 Then Exit Sub

    ' ถ้าดาเมจที่โจมตีมากกว่าเลือดศัตรูที่เหลือ จะทำให้ตาย
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
    
        If CanPlayerLHand(Attacker) Then
            SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32) + 8, (GetPlayerY(Victim) * 32)
        Else
            SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        End If
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
        
        ' Killer
        Select Case Player(Attacker).Killer
        Case 0: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 200
        Call SendAnimation(GetPlayerMap(Attacker), 200, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้เริ่มฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 1: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 201
        Call SendAnimation(GetPlayerMap(Attacker), 201, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่าฝึกหัด.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 2: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 202
        Call SendAnimation(GetPlayerMap(Attacker), 202, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 3: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 203
        Call SendAnimation(GetPlayerMap(Attacker), 203, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ชำนาญอาวุธ.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 4: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 204
        Call SendAnimation(GetPlayerMap(Attacker), 204, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่าไร้เงา.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 5: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 205
        Call SendAnimation(GetPlayerMap(Attacker), 205, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ปีศาจ.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 6: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 206
        Call SendAnimation(GetPlayerMap(Attacker), 206, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ปีศาจนักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 7: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 207
        Call SendAnimation(GetPlayerMap(Attacker), 207, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้พิฆาต.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 8: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 208
        Call SendAnimation(GetPlayerMap(Attacker), 208, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ตำนานนักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 9 To 13: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 209
        Call SendAnimation(GetPlayerMap(Attacker), 209, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้กระหายเลือด.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case Else: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 210
        Call SendAnimation(GetPlayerMap(Attacker), 210, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ราชาผู้อมตะ.", BrightRed)
        Call GlobalMsg("ใครก็ได้ฆ่า " & GetPlayerName(Attacker) & " ที.", Yellow)
        Player(Attacker).Killer = 15
        
        End Select
        
        If spellAnim < 1 Then
            ' แก้ไขบัค ไม่ตรวจสอบอาวุธเมื่อโจมตีผู้เล่น
            If n > 0 Then
                ' send anim
                Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Weapon)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
            Else
                SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
                Call SendAnimation(GetPlayerMap(Victim), PUNCH_ANIM, GetPlayerX(Victim), GetPlayerY(Victim))
            End If
        Else
            Call SendAnimation(GetPlayerMap(Victim), spellAnim, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " ได้ถูกฆ่าโดย " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        
        exp = (GetPlayerNextLevel(Victim) * 0.05)

        ' Make sure we dont get less then 0
        If GetPlayerExp(Victim) < exp Then
            exp = 0
        End If

        ' ตรวจสอบเงื่อนไขถ้า exp = 0 จะไม่เสียค่าประสบการณ์
        If exp = 0 Then
            Call PlayerMsg(Victim, "คุณไม่เหลือค่าประสบการณ์จากการตาย.", BrightRed)
            Call PlayerMsg(Attacker, "คุณไม่ได้รับ Exp จากการสังหาร (เพราะศัตรูไม่เหลือค่า Exp).", BrightBlue)
            Call SetPlayerExp(Victim, exp)
            SendEXP Victim
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - exp)
            Call PlayerMsg(Victim, "คุณได้สูญเสีย exp " & exp & " จากการตาย.", BrightRed)
            SendEXP Victim
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                    Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                Else
                    Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                    Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                    Call SetPlayerExp(Attacker, 1)
                    SendEXP Attacker
                End If
            Else
                ' no party - keep exp for self
                If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                    GivePlayerEXP Attacker, exp
                    Call PlayerMsg(Attacker, "คุณได้รับ " & exp & " Exp จากการสังหารผู้เล่น " & Player(Victim).Name, Yellow)
                Else
                    ' GivePlayerEXP Attacker, 0
                    Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                    Call SetPlayerExp(Attacker, 1)
                    SendEXP Attacker
                End If
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " ถูกตั้งค่าหัวว่าสังหารผู้เล่นอื่น !!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " ได้โดนปิดตำนานการสังหารผู้เล่น !!", BrightRed)
        End If
        
        Call CheckTasks(Attacker, QUEST_TYPE_GOKILL, Victim)
        Call OnDeath(Victim)
    Else
        ' ผู้เล่นไม่ตายและยังได้รับดาเมจ
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
        
        ' ถ้าผู้เล่นกำลังร่ายสกิล แล้วสกิลนั้นสามารถถูกยกเลิกได้ จะสั่งให้หยุดร่ายทันที
        If TempPlayer(Victim).spellBuffer.Spell > 0 Then
            If Spell(Player(Victim).Spell(TempPlayer(Victim).spellBuffer.Spell)).CanCancle > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
        End If
        
        If spellAnim < 1 Then
            ' แก้ไขบัค ไม่ตรวจสอบอาวุธเมื่อโจมตีผู้เล่น
            If n > 0 Then
                ' send anim
                Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Weapon)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
            Else
                SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
                Call SendAnimation(GetPlayerMap(Victim), PUNCH_ANIM, GetPlayerX(Victim), GetPlayerY(Victim))
            End If
        Else
            Call SendAnimation(GetPlayerMap(Victim), spellAnim, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        If CanPlayerLHand(Attacker) Then
            SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32) + 8, (GetPlayerY(Victim) * 32)
        Else
            SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        End If
        
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' Kick System
        If n > 0 Then
            If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > 0 Then
                If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > rand(1, 100) Then
                    ' set the values on index
                    TempPlayer(Victim).StunDuration = 2
                    TempPlayer(Victim).StunTimer = GetTickCount
                    ' send it to the index
                    SendStunned Victim
                    SendActionMsg GetPlayerMap(Victim), "ติดมึน !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
                    SendAnimation GetPlayerMap(Victim), Stun_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                    
            ' ถ้าผู้เล่นกำลังร่ายสกิล และติดมึนจะยกเลิกการร่ายสกิลทันที
            If TempPlayer(Victim).spellBuffer.Spell > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
            
                End If
            End If
        End If
        
        ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
        ' TempPlayer(Victim).stopRegen = True
        ' TempPlayer(Victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunPlayer Victim, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Player Victim, spellnum, Attacker
            End If
        End If
    End If

End Sub

Sub PlayerPassivePlayerLHand(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0, Optional ByVal spellAnim As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Shield) > 0 Then
        n = GetPlayerEquipment(Attacker, Shield)
    End If
    
    ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
    ' TempPlayer(Attacker).stopRegen = True
    ' TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Player(Victim).Vital(Vitals.HP) <= 0 Then Exit Sub

    ' ถ้าดาเมจที่โจมตีมากกว่าเลือดศัตรูที่เหลือ จะทำให้ตาย
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
    
        If CanPlayerLHand(Attacker) Then
            SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32) + 8, (GetPlayerY(Victim) * 32)
        Else
            SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        End If
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
        
        ' Killer
        Select Case Player(Attacker).Killer
        Case 0: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 200
        Call SendAnimation(GetPlayerMap(Attacker), 200, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้เริ่มฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 1: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 201
        Call SendAnimation(GetPlayerMap(Attacker), 201, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่าฝึกหัด.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 2: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 202
        Call SendAnimation(GetPlayerMap(Attacker), 202, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 3: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 203
        Call SendAnimation(GetPlayerMap(Attacker), 203, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ชำนาญอาวุธ.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 4: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 204
        Call SendAnimation(GetPlayerMap(Attacker), 204, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่าไร้เงา.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 5: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 205
        Call SendAnimation(GetPlayerMap(Attacker), 205, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ปีศาจ.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 6: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 206
        Call SendAnimation(GetPlayerMap(Attacker), 206, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ปีศาจนักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 7: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 207
        Call SendAnimation(GetPlayerMap(Attacker), 207, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้พิฆาต.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 8: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 208
        Call SendAnimation(GetPlayerMap(Attacker), 208, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ตำนานนักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 9 To 13: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 209
        Call SendAnimation(GetPlayerMap(Attacker), 209, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้กระหายเลือด.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case Else: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 210
        Call SendAnimation(GetPlayerMap(Attacker), 210, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ราชาผู้อมตะ.", BrightRed)
        Call GlobalMsg("ใครก็ได้ฆ่า " & GetPlayerName(Attacker) & " ที.", Yellow)
        Player(Attacker).Killer = 15
        
        End Select
        
        If spellAnim < 1 Then
            ' แก้ไขบัค ไม่ตรวจสอบอาวุธเมื่อโจมตีผู้เล่น
            If n > 0 Then
                ' send anim
                Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Shield)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
            Else
                SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
                Call SendAnimation(GetPlayerMap(Victim), PUNCH_ANIM, GetPlayerX(Victim), GetPlayerY(Victim))
            End If
        Else
            Call SendAnimation(GetPlayerMap(Victim), spellAnim, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " ได้ถูกฆ่าโดย " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        
        exp = (GetPlayerNextLevel(Victim) * 0.05)

        ' Make sure we dont get less then 0
        If GetPlayerExp(Victim) < exp Then
            exp = 0
        End If

        ' ตรวจสอบเงื่อนไขถ้า exp = 0 จะไม่เสียค่าประสบการณ์
        If exp = 0 Then
            Call PlayerMsg(Victim, "คุณไม่เหลือค่าประสบการณ์จากการตาย.", BrightRed)
            Call PlayerMsg(Attacker, "คุณไม่ได้รับ Exp จากการสังหาร (เพราะศัตรูไม่เหลือค่า Exp).", BrightBlue)
            Call SetPlayerExp(Victim, exp)
            SendEXP Victim
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - exp)
            Call PlayerMsg(Victim, "คุณได้สูญเสีย exp " & exp & " จากการตาย.", BrightRed)
            SendEXP Victim
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                    Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                Else
                    Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                    Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                    Call SetPlayerExp(Attacker, 1)
                    SendEXP Attacker
                End If
            Else
                ' no party - keep exp for self
                If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                    GivePlayerEXP Attacker, exp
                    Call PlayerMsg(Attacker, "คุณได้รับ " & exp & " Exp จากการสังหารผู้เล่น " & Player(Victim).Name, Yellow)
                Else
                    ' GivePlayerEXP Attacker, 0
                    Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                    Call SetPlayerExp(Attacker, 1)
                    SendEXP Attacker
                End If
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " ถูกตั้งค่าหัวว่าสังหารผู้เล่นอื่น !!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " ได้โดนปิดตำนานการสังหารผู้เล่น !!", BrightRed)
        End If
        
        Call CheckTasks(Attacker, QUEST_TYPE_GOKILL, Victim)
        Call OnDeath(Victim)
    Else
        ' ผู้เล่นไม่ตายและยังได้รับดาเมจ
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
        
        ' ถ้าผู้เล่นกำลังร่ายสกิล แล้วสกิลนั้นสามารถถูกยกเลิกได้ จะสั่งให้หยุดร่ายทันที
        If TempPlayer(Victim).spellBuffer.Spell > 0 Then
            If Spell(Player(Victim).Spell(TempPlayer(Victim).spellBuffer.Spell)).CanCancle > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
        End If
        
        If spellAnim < 1 Then
            ' แก้ไขบัค ไม่ตรวจสอบอาวุธเมื่อโจมตีผู้เล่น
            If n > 0 Then
                ' send anim
                Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Shield)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
            Else
                SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
                Call SendAnimation(GetPlayerMap(Victim), PUNCH_ANIM, GetPlayerX(Victim), GetPlayerY(Victim))
            End If
        Else
            Call SendAnimation(GetPlayerMap(Victim), spellAnim, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        If CanPlayerLHand(Attacker) Then
            SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32) + 8, (GetPlayerY(Victim) * 32)
        Else
            SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        End If
        
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' Kick System
        If n > 0 Then
            If Item(GetPlayerEquipment(Attacker, Shield)).Kick > 0 Then
                If Item(GetPlayerEquipment(Attacker, Shield)).Kick > rand(1, 100) Then
                    ' set the values on index
                    TempPlayer(Victim).StunDuration = 2
                    TempPlayer(Victim).StunTimer = GetTickCount
                    ' send it to the index
                    SendStunned Victim
                    SendActionMsg GetPlayerMap(Victim), "ติดมึน !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
                    SendAnimation GetPlayerMap(Victim), Stun_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                    
            ' ถ้าผู้เล่นกำลังร่ายสกิล และติดมึนจะยกเลิกการร่ายสกิลทันที
            If TempPlayer(Victim).spellBuffer.Spell > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
            
                End If
            End If
        End If
        
        ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
        ' TempPlayer(Victim).stopRegen = True
        ' TempPlayer(Victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunPlayer Victim, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Player Victim, spellnum, Attacker
            End If
        End If
    End If

End Sub


Sub PlayerAttackPlayerLHand(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim oldX As Long, oldY As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Shield) > 0 Then
        n = GetPlayerEquipment(Attacker, Shield)
    End If
    
    oldX = GetPlayerX(Victim)
    oldY = GetPlayerY(Victim)
    
    ' สกิลติดตัวทำงานเมื่อโจมตี? ใช้ได้แล้ว
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Spell(Player(Attacker).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Attacker).Spell(i)).PATK > 0 Then
                        If Spell(Player(Attacker).Spell(i)).PerSkill >= rand(1, 100) Then
                            If Spell(Player(Attacker).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMoveLHand(Attacker, i, Victim, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(Attacker, "[Damage] : " & Player(Attacker).Spell(i - 1), BrightGreen)
                            Else
                                Call CastSpellPassiveLHand(Attacker, i, Victim, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(Attacker, "[Heal] : " & Player(Attacker).Spell(i - 1), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(Attacker), Trim$(Spell(Player(Attacker).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16
                            'Call PlayerMsg(Attacker, "[สกิลติดตัว] : " & Trim(Spell(Player(Attacker).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
        
        ' สกิลติดตัวทำงานเมื่อถูกโจมตี? ใช้ได้แล้ว
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Victim).Spell(i) > 0 Then
                If Spell(Player(Victim).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Victim).Spell(i)).PDEF > 0 Then
                        If Spell(Player(Victim).Spell(i)).PerSkill + (Spell(Player(Victim).Spell(i)).S4 * Player(Victim).skillLV(i)) >= rand(1, 100) Then
                            If Spell(Player(Victim).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMove(Victim, i, Attacker, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(victim, "[Damage] : " & Player(victim).Spell(i - 1), BrightGreen)
                            Else
                                Call CastSpellPassive(Victim, i, Attacker, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(victim, "[Heal] : " & Player(victim).Spell(i - 1), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(victim), Trim$(Spell(Player(victim).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32) + 16
                            'Call PlayerMsg(victim, "[สกิลติดตัว] : " & Trim(Spell(Player(victim).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
        
        If oldX <> GetPlayerX(Victim) Or oldY <> GetPlayerY(Victim) Then Exit Sub
        If Player(Victim).Vital(Vitals.HP) <= 0 Then Exit Sub
        
        ' fixed bug
        'If Not CanPlayerAttackPlayerLHand(Attacker, Victim) Then Exit Sub

    ' ถ้าดาเมจที่โจมตีมากกว่าเลือดศัตรูที่เหลือ จะทำให้ตาย
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32) - 8, (GetPlayerY(Victim) * 32)
                
        ' Killer
        Select Case Player(Attacker).Killer
        Case 0: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 200
        Call SendAnimation(GetPlayerMap(Attacker), 200, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้เริ่มฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 1: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 201
        Call SendAnimation(GetPlayerMap(Attacker), 201, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่าฝึกหัด.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 2: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 202
        Call SendAnimation(GetPlayerMap(Attacker), 202, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 3: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 203
        Call SendAnimation(GetPlayerMap(Attacker), 203, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ชำนาญอาวุธ.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 4: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 204
        Call SendAnimation(GetPlayerMap(Attacker), 204, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่าไร้เงา.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 5: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 205
        Call SendAnimation(GetPlayerMap(Attacker), 205, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ปีศาจ.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 6: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 206
        Call SendAnimation(GetPlayerMap(Attacker), 206, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ปีศาจนักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 7: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 207
        Call SendAnimation(GetPlayerMap(Attacker), 207, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้พิฆาต.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 8: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 208
        Call SendAnimation(GetPlayerMap(Attacker), 208, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ตำนานนักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 9 To 13: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 209
        Call SendAnimation(GetPlayerMap(Attacker), 209, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้กระหายเลือด.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case Else: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 210
        Call SendAnimation(GetPlayerMap(Attacker), 210, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ราชาผู้อมตะ.", BrightRed)
        Call GlobalMsg("ใครก็ได้ฆ่า " & GetPlayerName(Attacker) & " ที.", Yellow)
        Player(Attacker).Killer = 15
        
        End Select
                
        ' แก้ไขบัค ไม่ตรวจสอบอาวุธเมื่อโจมตีผู้เล่น
        If n > 0 Then
            ' send anim
            Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Shield)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
        Else
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(GetPlayerMap(Victim), PUNCH_ANIM, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " ได้ถูกฆ่าโดย " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        
        exp = (GetPlayerNextLevel(Victim) * 0.05)

        ' Make sure we dont get less then 0
        If GetPlayerExp(Victim) < exp Then
            exp = 0
        End If

        ' ตรวจสอบเงื่อนไขถ้า exp = 0 จะไม่เสียค่าประสบการณ์
        If exp = 0 Then
            Call PlayerMsg(Victim, "คุณไม่เหลือค่าประสบการณ์จากการตาย.", BrightRed)
            Call PlayerMsg(Attacker, "คุณไม่ได้รับ Exp จากการสังหาร (เพราะศัตรูไม่เหลือค่า Exp).", BrightBlue)
            Call SetPlayerExp(Victim, exp)
            SendEXP Victim
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - exp)
            Call PlayerMsg(Victim, "คุณได้สูญเสีย exp " & exp & " จากการตาย.", BrightRed)
            SendEXP Victim
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                    Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                Else
                    Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                    Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                    Call SetPlayerExp(Attacker, 1)
                    SendEXP Attacker
                End If
            Else
                ' no party - keep exp for self
                If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                    GivePlayerEXP Attacker, exp
                    Call PlayerMsg(Attacker, "คุณได้รับ " & exp & " Exp จากการสังหารผู้เล่น " & Player(Victim).Name, Yellow)
                Else
                    ' GivePlayerEXP Attacker, 0
                    Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                    Call SetPlayerExp(Attacker, 1)
                    SendEXP Attacker
                End If
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " ถูกตั้งค่าหัวว่าสังหารผู้เล่นอื่น !!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " ได้โดนปิดตำนานการสังหารผู้เล่น !!", BrightRed)
        End If
        
        Call CheckTasks(Attacker, QUEST_TYPE_GOKILL, Victim)
        Call OnDeath(Victim)
    Else
        ' ผู้เล่นไม่ตายและยังได้รับดาเมจ
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' ถ้าผู้เล่นกำลังร่ายสกิล แล้วสกิลนั้นสามารถถูกยกเลิกได้ จะสั่งให้หยุดร่ายทันที
        If TempPlayer(Victim).spellBuffer.Spell > 0 Then
            If Spell(Player(Victim).Spell(TempPlayer(Victim).spellBuffer.Spell)).CanCancle > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
        End If
        
        ' แก้ไขบัค ไม่ตรวจสอบอาวุธเมื่อโจมตีผู้เล่น
        If n > 0 Then
            ' send anim
            Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Shield)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
        Else
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(GetPlayerMap(Victim), PUNCH_ANIM, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32) - 8, (GetPlayerY(Victim) * 32) - 8
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' Kick System
        If n > 0 Then
            If Item(GetPlayerEquipment(Attacker, Shield)).Kick > 0 Then
                If Item(GetPlayerEquipment(Attacker, Shield)).Kick > rand(1, 100) Then
                    ' set the values on index
                    TempPlayer(Victim).StunDuration = 2
                    TempPlayer(Victim).StunTimer = GetTickCount
                    ' send it to the index
                    SendStunned Victim
                    SendActionMsg GetPlayerMap(Victim), "ติดมึน !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
                    SendAnimation GetPlayerMap(Victim), Stun_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                    
            ' ถ้าผู้เล่นกำลังร่ายสกิล และติดมึนจะยกเลิกการร่ายสกิลทันที
            If TempPlayer(Victim).spellBuffer.Spell > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
            
                End If
            End If
        End If
        
    End If
    
End Sub

Sub PlayerReflectPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, ByVal LHand As Byte, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim oldX As Long, oldY As Long
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Then
        Exit Sub
    End If
        
    ' Check for weapon
    n = 0
    
    If LHand = 0 Then
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            n = GetPlayerEquipment(Attacker, Weapon)
        End If
    Else
        If GetPlayerEquipment(Attacker, Shield) > 0 Then
            n = GetPlayerEquipment(Attacker, Shield)
        End If
    End If
    
    oldX = GetPlayerX(Victim)
    oldY = GetPlayerY(Victim)
    
    ' แก้ไขการแจ้งเตือนเมื่อสะท้อนไม่มีดาเมจ
    If Damage <= 0 Then
        SendAnimation GetPlayerMap(Victim), PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
        SendActionMsg GetPlayerMap(Victim), "อ่อนหัด !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
        Exit Sub
    End If
    
    ' สกิลติดตัวทำงานเมื่อโจมตี? ใช้ได้แล้ว
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Attacker).Spell(i) > 0 Then
                If Spell(Player(Attacker).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Attacker).Spell(i)).PATK > 0 Then
                        If Spell(Player(Attacker).Spell(i)).PerSkill >= rand(1, 100) Then
                            If Spell(Player(Attacker).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMove(Attacker, i, Victim, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(Attacker, "[Damage] : " & Player(Attacker).Spell(i - 1), BrightGreen)
                            Else
                                Call CastSpellPassive(Attacker, i, Victim, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(Attacker, "[Heal] : " & Player(Attacker).Spell(i - 1), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(Attacker), Trim$(Spell(Player(Attacker).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32) + 16
                            'Call PlayerMsg(Attacker, "[สกิลติดตัว] : " & Trim(Spell(Player(Attacker).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
        
        ' สกิลติดตัวทำงานเมื่อถูกโจมตี? ใช้ได้แล้ว
        For i = 1 To MAX_PLAYER_SPELLS
            If Player(Victim).Spell(i) > 0 Then
                If Spell(Player(Victim).Spell(i)).Name <> vbNullString Then
                    If Spell(Player(Victim).Spell(i)).PDEF > 0 Then
                        If Spell(Player(Victim).Spell(i)).PerSkill + (Spell(Player(Victim).Spell(i)).S4 * Player(Victim).skillLV(i)) >= rand(1, 100) Then
                            If Spell(Player(Victim).Spell(i)).CanMove > 0 Then
                                Call CastSpellCanMove(Victim, i, Attacker, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(victim, "[Damage] : " & Player(victim).Spell(i - 1), BrightGreen)
                            Else
                                Call CastSpellPassive(Victim, i, Attacker, TARGET_TYPE_PLAYER)
                                'Call PlayerMsg(victim, "[Heal] : " & Player(victim).Spell(i - 1), BrightGreen)
                            End If
                            'SendActionMsg GetPlayerMap(victim), Trim$(Spell(Player(victim).Spell(i)).Name) & " !", BrightGreen, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32) + 16
                            'Call PlayerMsg(victim, "[สกิลติดตัว] : " & Trim(Spell(Player(victim).Spell(i)).Name), BrightGreen)
                        End If
                    End If
                End If
            End If
        Next
        
        If oldX <> GetPlayerX(Victim) Or oldY <> GetPlayerY(Victim) Then Exit Sub
        If Player(Victim).Vital(Vitals.HP) <= 0 Then Exit Sub
        
        ' fixed bug
        'If spellnum > 0 Then
        '    If Not CanPlayerAttackPlayerLHand(Attacker, Victim, True) Then Exit Sub
        'Else
        '    If Not CanPlayerAttackPlayerLHand(Attacker, Victim) Then Exit Sub
        'End If
   
    ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
    ' TempPlayer(Attacker).stopRegen = True
    ' TempPlayer(Attacker).stopRegenTimer = GetTickCount

    ' ถ้าดาเมจที่โจมตีมากกว่าเลือดศัตรูที่เหลือ จะทำให้ตาย
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
    
        ' Killer
        Select Case Player(Attacker).Killer
        Case 0: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 200
        Call SendAnimation(GetPlayerMap(Attacker), 200, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้เริ่มฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 1: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 201
        Call SendAnimation(GetPlayerMap(Attacker), 201, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่าฝึกหัด.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 2: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 202
        Call SendAnimation(GetPlayerMap(Attacker), 202, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 3: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 203
        Call SendAnimation(GetPlayerMap(Attacker), 203, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ชำนาญอาวุธ.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 4: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 204
        Call SendAnimation(GetPlayerMap(Attacker), 204, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > นักฆ่าไร้เงา.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 5: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 205
        Call SendAnimation(GetPlayerMap(Attacker), 205, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ปีศาจ.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 6: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 206
        Call SendAnimation(GetPlayerMap(Attacker), 206, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ปีศาจนักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 7: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 207
        Call SendAnimation(GetPlayerMap(Attacker), 207, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้พิฆาต.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 8: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 208
        Call SendAnimation(GetPlayerMap(Attacker), 208, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ตำนานนักฆ่า.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case 9 To 13: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 209
        Call SendAnimation(GetPlayerMap(Attacker), 209, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ผู้กระหายเลือด.", BrightRed)
        Player(Attacker).Killer = Player(Attacker).Killer + 1
        
        Case Else: SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 210
        Call SendAnimation(GetPlayerMap(Attacker), 210, 0, 0, TARGET_TYPE_PLAYER, Attacker)
        Call GlobalMsg("ผู้เล่น " & GetPlayerName(Attacker) & " ได้รับฉายา > ราชาผู้อมตะ.", BrightRed)
        Call GlobalMsg("ใครก็ได้ฆ่า " & GetPlayerName(Attacker) & " ที.", Yellow)
        Player(Attacker).Killer = 15
        
        End Select
    
        If LHand = 1 Then
            SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), Yellow, 1, (GetPlayerX(Victim) * 32) + 8, (GetPlayerY(Victim) * 32)
        Else
            SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        End If
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
        
        If Not spellnum > 0 Then
        
        ' แก้ไขบัค ไม่ตรวจสอบอาวุธเมื่อโจมตีผู้เล่น
        If n > 0 Then
            If LHand = 0 Then
                ' send anim
                Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Weapon)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
            Else
                ' send anim
                Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Shield)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
            End If
        Else
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(GetPlayerMap(Victim), PUNCH_ANIM, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " ได้ถูกฆ่าโดย " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        
        exp = (GetPlayerNextLevel(Victim) * 0.05)

        ' Make sure we dont get less then 0
        If GetPlayerExp(Victim) < exp Then
            exp = 0
        End If

        ' ตรวจสอบเงื่อนไขถ้า exp = 0 จะไม่เสียค่าประสบการณ์
        If exp = 0 Then
            Call PlayerMsg(Victim, "คุณไม่เหลือค่าประสบการณ์จากการตาย.", BrightRed)
            Call PlayerMsg(Attacker, "คุณไม่ได้รับ Exp จากการสังหาร (เพราะศัตรูไม่เหลือค่า Exp).", BrightBlue)
            Call SetPlayerExp(Victim, exp)
            SendEXP Victim
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - exp)
            Call PlayerMsg(Victim, "คุณได้สูญเสีย exp " & exp & " จากการตาย.", BrightRed)
            SendEXP Victim
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                    Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                Else
                    Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker
                    Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                    Call SetPlayerExp(Attacker, 1)
                    SendEXP Attacker
                End If
            Else
                ' no party - keep exp for self
                If GetPlayerLevel(Attacker) <> MAX_LEVELS Then
                    GivePlayerEXP Attacker, exp
                    Call PlayerMsg(Attacker, "คุณได้รับ " & exp & " Exp จากการสังหารผู้เล่น " & Player(Victim).Name, Yellow)
                Else
                    ' GivePlayerEXP Attacker, 0
                    Call PlayerMsg(Attacker, "คุณมีเลเวลสูงสุดแล้ว.", BrightRed)
                    Call SetPlayerExp(Attacker, 1)
                    SendEXP Attacker
                End If
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " ถูกตั้งค่าหัวว่าสังหารผู้เล่นอื่น !!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " ได้โดนปิดตำนานการสังหารผู้เล่น !!", BrightRed)
        End If
        
        Call CheckTasks(Attacker, QUEST_TYPE_GOKILL, Victim)
        Call OnDeath(Victim)
    Else
        ' ผู้เล่นไม่ตายและยังได้รับดาเมจ
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        If spellnum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
        
        ' ถ้าผู้เล่นกำลังร่ายสกิล แล้วสกิลนั้นสามารถถูกยกเลิกได้ จะสั่งให้หยุดร่ายทันที
        If TempPlayer(Victim).spellBuffer.Spell > 0 Then
            If Spell(Player(Victim).Spell(TempPlayer(Victim).spellBuffer.Spell)).CanCancle > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
        End If
        
        If Not spellnum > 0 Then
        
        ' แก้ไขบัค ไม่ตรวจสอบอาวุธเมื่อโจมตีผู้เล่น
        If n > 0 Then
            If LHand = 0 Then
                ' send anim
                Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Weapon)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
            Else
                ' send anim
                Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Shield)).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
            End If
        Else
            SendPlayerSound Attacker, GetPlayerX(Attacker), GetPlayerY(Attacker), SoundEntity.sePunch, 1
            Call SendAnimation(GetPlayerMap(Victim), PUNCH_ANIM, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        End If
        
        If LHand = 1 Then
            SendActionMsg GetPlayerMap(Victim), "-" & Damage, Yellow, 1, (GetPlayerX(Victim) * 32) + 8, (GetPlayerY(Victim) * 32)
        Else
            SendActionMsg GetPlayerMap(Victim), "-" & Damage, Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        End If
        
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' Kick System
        If n > 0 Then
            If LHand = 0 Then
                If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > 0 Then
                    If Item(GetPlayerEquipment(Attacker, Weapon)).Kick > rand(1, 100) Then
                        ' set the values on index
                        TempPlayer(Victim).StunDuration = 2
                        TempPlayer(Victim).StunTimer = GetTickCount
                        ' send it to the index
                        SendStunned Victim
                        SendActionMsg GetPlayerMap(Victim), "ติดมึน !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
                        SendAnimation GetPlayerMap(Victim), Stun_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                        
            ' ถ้าผู้เล่นกำลังร่ายสกิล และติดมึนจะยกเลิกการร่ายสกิลทันที
            If TempPlayer(Victim).spellBuffer.Spell > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
            
                    End If
                End If
            Else
                If Item(GetPlayerEquipment(Attacker, Shield)).Kick > 0 Then
                    If Item(GetPlayerEquipment(Attacker, Shield)).Kick > rand(1, 100) Then
                        ' set the values on index
                        TempPlayer(Victim).StunDuration = 2
                        TempPlayer(Victim).StunTimer = GetTickCount
                        ' send it to the index
                        SendStunned Victim
                        SendActionMsg GetPlayerMap(Victim), "ติดมึน !", Yellow, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) + 16
                        SendAnimation GetPlayerMap(Victim), Stun_ANIM, 0, 0, TARGET_TYPE_PLAYER, Victim
                        
            ' ถ้าผู้เล่นกำลังร่ายสกิล และติดมึนจะยกเลิกการร่ายสกิลทันที
            If TempPlayer(Victim).spellBuffer.Spell > 0 Then
                ' Clear spell casting
                TempPlayer(Victim).spellBuffer.Spell = 0
                TempPlayer(Victim).spellBuffer.Timer = 0
                TempPlayer(Victim).spellBuffer.Target = 0
                TempPlayer(Victim).spellBuffer.tType = 0
                SendPlayerData Victim
                Call SendClearSpellBuffer(Victim)
                SendActionMsg GetPlayerMap(Victim), "ขัดจังหวะ !", BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32) - 8
            End If
            
                    End If
                End If
            End If
        End If
        
        ' หยุดการฟื้นฟู hp & mp ถ้าโดนโจมตี
        ' TempPlayer(Victim).stopRegen = True
        ' TempPlayer(Victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunPlayer Victim, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Player Victim, spellnum, Attacker
            End If
        End If
    End If

End Sub


' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal index As Long, ByVal spellslot As Long)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    Dim HPCost As Long
    
    Dim targetType As Byte
    Dim Target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)
    HasBuffered = False
    
    If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg index, "สกิลยังอยู่ในสถานะดีเลย์ โปรดรออีก ! " & TempPlayer(index).SpellCD(spellslot) / 1000 & " วินาที.", BrightRed
        SendClearSpellBuffer index
        Exit Sub
    End If

    MPCost = Spell(spellnum).MPCost
    HPCost = Spell(spellnum).HPCost

    If GetPlayerVital(index, Vitals.HP) < HPCost Then
        Call PlayerMsg(index, "ต้องการ Hp " & HPCost & " ในการใช้สกิลนี้.", BrightRed)
        SendClearSpellBuffer index
        Exit Sub
    End If

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "ต้องการ Mp " & MPCost & " เพื่อใช้สกิลนี้.", BrightRed)
        SendClearSpellBuffer index
        Exit Sub
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "ต้องการเลเวล " & LevelReq & " ในการใช้สกิลนี้.", BrightRed)
        SendClearSpellBuffer index
        Exit Sub
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "ต้องการ GM ในการใช้สกิลนี้.", BrightRed)
        SendClearSpellBuffer index
        Exit Sub
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้สกิลนี้.", BrightRed)
            SendClearSpellBuffer index
            Exit Sub
        End If
    End If
    
    ' fixed ! bug of toxin
    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
        If Player(index).BuffStatus(BUFF_TOXIN) = BUFF_TOXIN Then
            Call PlayerMsg(index, "คุณอยู่ในสถานะห้ามฟื้นฟู Hp.", BrightRed)
            SendClearSpellBuffer index
            Exit Sub
        End If
    End If
    
    ' ไม่สามารถกดใช้สกิลติดตัวได้?
    If Spell(spellnum).Passive > 0 Then
        Call PlayerMsg(index, "ไม่สามารถกดใช้สกิลติดตัวได้.", BrightRed)
        SendClearSpellBuffer index
        Exit Sub
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 2 ' Targetted
        Else
            SpellCastType = 3 ' Targetted AOE
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 0 ' Self-cast
        Else
            SpellCastType = 1 ' Self-cast AOE
        End If
    End If
    
    targetType = TempPlayer(index).targetType
    Target = TempPlayer(index).Target
    Range = Spell(spellnum).Range
        
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                PlayerMsg index, "คุณไม่มีเป้าหมายที่ต้องการใช้สกิล.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg index, "เป้าหมายอยู่ไกลเกินระยะสกิล.", BrightRed
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y) Then
                    PlayerMsg index, "เป้าหมายอยู่ไกลเกินระยะสกิล.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapnum, Spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index
        
        ' สกิลติดตัว ต้องไม่ขึ้นแจ้งร่ายเวทย์
        If Spell(spellnum).Passive <= 0 Then
            SendActionMsg mapnum, "กำลังร่าย " & Trim$(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] " & " !", BrightCyan, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        End If
        
        TempPlayer(index).spellBuffer.Spell = spellslot
        TempPlayer(index).spellBuffer.Timer = GetTickCount
        TempPlayer(index).spellBuffer.Target = TempPlayer(index).Target
        TempPlayer(index).spellBuffer.tType = TempPlayer(index).targetType

        SendPlayerData index
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If
    
End Sub

' ระบบสกิล V1.0 แบบกายภาพ
Public Sub CastSpell(ByVal index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal targetType As Byte)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    Dim HPCost As Long
    Dim xt As Long
    Dim yt As Long
    Dim curProjecTile As Long, CurEquipment As Long
    Dim s(1 To 2) As Long
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    Dim Dur As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub

    MPCost = Spell(spellnum).MPCost
    HPCost = Spell(spellnum).HPCost

    ' spell fixed
    If Spell(spellnum).CanMove > 0 Then
        Call CastSpellSCanMove(index, spellslot, Target, targetType)
        Exit Sub
    End If
    
    If TempPlayer(index).StunDuration > 0 Then
        Call PlayerMsg(index, "ไม่สามารถใช้สกิลนี้ขณะติดมึนได้", BrightRed)
        Exit Sub
    End If
    
    If GetPlayerVital(index, Vitals.HP) < HPCost Then
        Call PlayerMsg(index, "ต้องการ Hp " & HPCost & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "ต้องการ Mp " & MPCost & " ในการใช้สกิลนี้", BrightRed)
        Exit Sub
    End If
   
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "ต้องการเลเวล " & LevelReq & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If
   
    AccessReq = Spell(spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "ต้องการ GM ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If
   
    ClassReq = Spell(spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้สกิลนี้.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' fixed ! bug of toxin
    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
        If Player(index).BuffStatus(BUFF_TOXIN) = BUFF_TOXIN Then
            Call PlayerMsg(index, "คุณอยู่ในสถานะห้ามฟื้นฟู Hp.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' ไม่สามารถกดใช้สกิลติดตัวได้?
    'If Spell(spellnum).Passive > 0 Then
    '    Call PlayerMsg(index, "ไม่สามารถกดใช้สกิลติดตัวได้.", BrightRed)
    '    Exit Sub
   ' End If
   
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
   
   '-------- Spell Damage V 2.0 ---------
   
   ' set damage for spell not pet spell
   If Spell(spellnum).Type <> SPELL_TYPE_PET Then
   
        Vital = Spell(spellnum).Vital + (Spell(spellnum).Vital * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
   End If
   
    ' ระบบสัตว์เลี้ยง 1.0 Vital = Pet number with spell
    If Spell(spellnum).Type = SPELL_TYPE_PET Then

        SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
        Call SpawnPet(index, GetPlayerMap(index), Spell(spellnum).Vital)
        PetFollowOwner index
        DidCast = True

    End If
    
    ' ตั้งค่า Vital for projectile
    If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        Vital = Spell(spellnum).Projectile.Damage + (Spell(spellnum).Projectile.Damage * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
    End If
    
    ' add script mode
    If Spell(spellnum).Type = SPELL_TYPE_SCRIPT Then
        Vital = Spell(spellnum).Vital
    End If
    
    ' -------- End Damage --------
    
    AoE = Spell(spellnum).AoE
    Range = Spell(spellnum).Range
   
    Select Case SpellCastType
        Case 0 ' เลือกเป้าหมายเป็นตัวเอง
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_PET  ' ระบบสัตว์เลี้ยง
                
                Case SPELL_TYPE_SCRIPT
                    ' Script mode
                    Call UseScript(index, Vital, TempPlayer(index).Target, Spell(spellnum).Duration)
                    Call SendAnimation(GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index)
                    DidCast = True ' fixed
                    
                Case SPELL_TYPE_HEALHP
                    'Call PlayerMsg(index, "9", BrightRed)
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_PROJECTILE
                    DidCast = True ' <<< Fixed
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    
                    ' ต้องเอาตั้งค่า Dir ออกเพราะบัค ทำให้ไม่วาร์ปตามพิกัดที่คำนวน
                    ' SetPlayerDir index, Spell(spellnum).Dir
                    
                    PlayerWarp index, Spell(spellnum).Map, Spell(spellnum).x, Spell(spellnum).y
                    
        ' สกิลวาร์ปตามแกน เมื่อสกิลเป็นแบบวาร์ปและ แผนที่ = 0,  x คือ พุ่งไปข้างหน้า และ y คือพุ่งถอยหลัง.
        
         If Spell(spellnum).Map = 0 Then
         
         If Player(index).Dir = 0 Then ' Dir Up
         xt = Player(index).x
         yt = Player(index).y - (Spell(spellnum).x) + (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then ' ถ้าไปเจอสิ่งกีดขวาง ทำให้ส่งผลอยู่ที่เดิม
             SetPlayerX index, xt
             SetPlayerY index, yt
             SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 1 Then ' Dir Down
         xt = Player(index).x
         yt = Player(index).y + (Spell(spellnum).x) - (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 2 Then ' Dir Left
         xt = Player(index).x - (Spell(spellnum).x) + (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 3 Then ' Dir right
         xt = Player(index).x + (Spell(spellnum).x) - (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
     End If
                    
            SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
            DidCast = True
            
            End Select
            
        Case 1, 3 ' สกิลแบบ AOE และ AOE แบบรอบเป้าหมาย
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
            
                If targetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                Else
                    x = MapNpc(mapnum).NPC(Target).x
                    y = MapNpc(mapnum).NPC(Target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            If Not CanPlayerAbsorbMagic(i) Then
                                                If Vital > GetPlayerMDEF(i) Then
                                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                                    PlayerPassivePlayer index, i, Vital - GetPlayerMDEF(i), spellnum
                                                Else
                                                    SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                    SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                                End If
                                            Else
                                                ' Absorb
                                                SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        If Not CanNpcAbsorbMagic(MapNpc(mapnum).NPC(i).num) Then
                                            If Vital - rand(NPC(i).stat(intelligence), NPC(i).stat(intelligence) * 2) > 0 Then
                                                SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, i
                                                PlayerPassiveNpc index, i, Vital - rand(NPC(i).stat(intelligence), NPC(i).stat(intelligence) * 2), spellnum
                                            Else
                                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                            End If
                                        Else
                                            ' Absorb
                                            SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                            SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = False
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    
                If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                                        If TempPlayer(i).inParty = TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    Else
                                        SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then Exit Sub
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                Else
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If TempPlayer(i).inParty <> TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                End If
            End Select
            
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            Else
                x = MapNpc(mapnum).NPC(Target).x
                y = MapNpc(mapnum).NPC(Target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
           
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            If Vital > GetPlayerMDEF(Target) Then
                                If Not CanPlayerAbsorbMagic(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerPassivePlayer index, Target, Vital - GetPlayerMDEF(Target), spellnum
                                    DidCast = True
                                Else
                                    'Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2) > 0 Then
                                If Not CanNpcAbsorbMagic(MapNpc(mapnum).NPC(Target).num) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerPassiveNpc index, Target, Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2), spellnum
                                    DidCast = True
                                Else
                                    ' Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True  ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = True
                        DidCast = True
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                        End If
                    End If
                    
                    ' Fixed spell type warp attack
                    Case SPELL_TYPE_WARP
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        
                        Select Case GetPlayerDir(Target)
                        
                        Case DIR_UP
                            If Player(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y + 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_DOWN
                            If Player(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y - 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_LEFT
                            If Player(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(Player(Target).x + 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, Player(Target).x + 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                        Case DIR_RIGHT
                            If Player(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x - 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, Player(Target).x - 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                            
                        End Select
                        
                        If CanPlayerAttackPlayerLHand(index, Target, True) Then
                            If Vital > GetPlayerMDEF(Target) Then
                                If Not CanPlayerAbsorbMagic(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerPassivePlayer index, Target, Vital - GetPlayerMDEF(Target), spellnum
                                    DidCast = True
                                Else
                                    'Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                        
                    Else
                    
                        Select Case MapNpc(mapnum).NPC(Target).Dir
                        
                        Case DIR_UP
                            If MapNpc(mapnum).NPC(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y + 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "UP : " & DIR_UP, BrightRed)
                            End If
                        Case DIR_DOWN
                            If MapNpc(mapnum).NPC(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y - 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "DOWN : " & DIR_DOWN, BrightRed)
                            End If
                        Case DIR_LEFT
                            If MapNpc(mapnum).NPC(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x + 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x + 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "LEFT : " & DIR_LEFT, BrightRed)
                            End If
                        Case DIR_RIGHT
                            If MapNpc(mapnum).NPC(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x - 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x - 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "RIGHT : " & DIR_RIGHT, BrightRed)
                            End If
                            
                        End Select
                    
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2) > 0 Then
                                If Not CanNpcAbsorbMagic(MapNpc(mapnum).NPC(Target).num) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerPassiveNpc index, Target, Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2), spellnum
                                    DidCast = True
                                Else
                                    ' Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                            End If
                        End If
                    End If
                    
            End Select
    End Select
    
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) - HPCost)
        Call SendVital(index, Vitals.HP)
        ' send vitals to party if in one
        
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
       
        TempPlayer(index).SpellCD(spellslot) = GetTickCount + (Spell(spellnum).CDTime * 1000)
        ' ส่งค่า ดีเลย์สกิล
        Call SendCooldown(index, spellslot)
        
        If Not Spell(spellnum).Map > 0 Then
            SendActionMsg mapnum, Trim$(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] " & " !", BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32 + 8
            If Spell(spellnum).Passive > 0 Then
                Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] ", BrightGreen)
            End If
        End If
        
        ' แก้ไขบัคสกิลแบบธนู
        If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        ' Spell New type fixed
        If Spell(spellnum).Projectile.Pic > 0 Then
        
        ' Call ProjecTileSpell(index, spellnum)
            
        ' prevent subscript
        If index > MAX_PLAYERS Or index < 1 Then Exit Sub
        
        ' get the players current equipment
        CurEquipment = GetPlayerSpell(index, spellslot)

        ' check if they've got equipment
        If CurEquipment < 1 Or CurEquipment > MAX_SPELLS Then Exit Sub

        ' set the curprojectile
        For i = 1 To MAX_PLAYER_PROJECTILES
            If TempPlayer(index).Projectile(i).Pic = 0 Then
            ' just incase there is left over data
            ClearProjectile index, i
            ' set the curprojtile
            curProjecTile = i
            Exit For
            End If
        Next

        ' check for subscript
        If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

        ' populate the data in the player rec
        With TempPlayer(index).Projectile(curProjecTile)
        .Damage = Vital
        .Direction = GetPlayerDir(index)
        .Pic = Spell(CurEquipment).Projectile.Pic
        .Range = Spell(CurEquipment).Projectile.Range
        .Speed = Spell(CurEquipment).Projectile.Speed
        .x = GetPlayerX(index)
        .y = GetPlayerY(index)
        End With

            ' trololol, they have no more projectile space left
            If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

            ' update the projectile on the map
            SendProjectileToMap index, curProjecTile
        
            End If
        
        ' Send the update
        Call SendStats(index)
        SendPlayerData index
    End If
    
    End If
    
End Sub

' ระบบสกิล V1.0 แบบกายภาพ
Public Sub CastSpellSCanMove(ByVal index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal targetType As Byte)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    Dim HPCost As Long
    Dim xt As Long
    Dim yt As Long
    Dim curProjecTile As Long, CurEquipment As Long
    Dim s(1 To 2) As Long
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    Dim Dur As Long
    Dim num As Long
       
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        'PlayerMsg index, "สกิลยังอยู่ในสถานะดีเลย์ โปรดรออีก ! " & TempPlayer(index).SpellCD(spellslot) / 1000 & " วินาที.", BrightRed
        'Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name) & " กำลังดีเลย์.", BrightRed)
        Exit Sub
    End If
    
    MPCost = Spell(spellnum).MPCost
    HPCost = Spell(spellnum).HPCost
    
    If TempPlayer(index).StunDuration > 0 Then
        Call PlayerMsg(index, "ไม่สามารถใช้สกิลนี้ขณะติดมึนได้", BrightRed)
        Exit Sub
    End If

    If GetPlayerVital(index, Vitals.HP) < HPCost Then
        Call PlayerMsg(index, "ต้องการ Hp " & HPCost & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "ต้องการ Mp " & MPCost & " ในการใช้สกิลนี้", BrightRed)
        Exit Sub
    End If
   'Call PlayerMsg(index, "4", BrightRed)
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "ต้องการเลเวล " & LevelReq & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    AccessReq = Spell(spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "ต้องการ GM ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ClassReq = Spell(spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้สกิลนี้.", BrightRed)
            Exit Sub
        End If
    End If

    ' fixed ! bug of toxin
    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
        If Player(index).BuffStatus(BUFF_TOXIN) = BUFF_TOXIN Then
            Call PlayerMsg(index, "คุณอยู่ในสถานะห้ามฟื้นฟู Hp.", BrightRed)
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
   
   '-------- Spell Damage V 2.0 ---------
   
   ' set damage for spell not pet spell
   If Spell(spellnum).Type <> SPELL_TYPE_PET Then
   
        Vital = Spell(spellnum).Vital + (Spell(spellnum).Vital * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
   End If
   
    ' ระบบสัตว์เลี้ยง 1.0 Vital = Pet number with spell
    If Spell(spellnum).Type = SPELL_TYPE_PET Then

        SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
        Call SpawnPet(index, GetPlayerMap(index), Spell(spellnum).Vital)
        PetFollowOwner index
        DidCast = True

    End If
    
    ' ตั้งค่า Vital for projectile
    If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        Vital = Spell(spellnum).Projectile.Damage + (Spell(spellnum).Projectile.Damage * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
    End If
    
    ' add script mode
    If Spell(spellnum).Type = SPELL_TYPE_SCRIPT Then
        Vital = Spell(spellnum).Vital
    End If
    
    ' -------- End Damage --------
    
    AoE = Spell(spellnum).AoE
    Range = Spell(spellnum).Range
       
    Select Case SpellCastType
        Case 0 ' เลือกเป้าหมายเป็นตัวเอง
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_PET  ' ระบบสัตว์เลี้ยง
                
                Case SPELL_TYPE_SCRIPT
                    ' Script mode
                    Call UseScript(index, Vital, TempPlayer(index).Target, Spell(spellnum).Duration)
                    Call SendAnimation(GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index)
                    DidCast = True ' fixed
                
                Case SPELL_TYPE_HEALHP
                    'Call PlayerMsg(index, "9", BrightRed)
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_PROJECTILE
                    DidCast = True ' <<< Fixed
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    
                    ' ต้องเอาตั้งค่า Dir ออกเพราะบัค ทำให้ไม่วาร์ปตามพิกัดที่คำนวน
                    ' SetPlayerDir index, Spell(spellnum).Dir
                    
                    PlayerWarp index, Spell(spellnum).Map, Spell(spellnum).x, Spell(spellnum).y
                    
        ' สกิลวาร์ปตามแกน เมื่อสกิลเป็นแบบวาร์ปและ แผนที่ = 0,  x คือ พุ่งไปข้างหน้า และ y คือพุ่งถอยหลัง.
        
         If Spell(spellnum).Map = 0 Then
         
         If Player(index).Dir = 0 Then ' Dir Up
         xt = Player(index).x
         yt = Player(index).y - (Spell(spellnum).x) + (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then ' ถ้าไปเจอสิ่งกีดขวาง ทำให้ส่งผลอยู่ที่เดิม
             SetPlayerX index, xt
             SetPlayerY index, yt
             SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 1 Then ' Dir Down
         xt = Player(index).x
         yt = Player(index).y + (Spell(spellnum).x) - (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 2 Then ' Dir Left
         xt = Player(index).x - (Spell(spellnum).x) + (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 3 Then ' Dir right
         xt = Player(index).x + (Spell(spellnum).x) - (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
     End If
                    
            SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
            DidCast = True
            
            End Select
            
        Case 1, 3 ' สกิลแบบ AOE และ AOE แบบรอบเป้าหมาย
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
            
                If targetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                Else
                    x = MapNpc(mapnum).NPC(Target).x
                    y = MapNpc(mapnum).NPC(Target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            If Vital > GetPlayerDef(i) Then
                                                If Not CanPlayerDodge(i) Then
                                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                                    PlayerReflectPlayer index, i, Vital - GetPlayerDef(i), 0, spellnum
                                                Else
                                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                                End If
                                            Else
                                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        If Vital > GetNpcDEF(MapNpc(mapnum).NPC(i).num) Then
                                            If Not CanNpcDodge(mapnum, MapNpc(mapnum).NPC(i).num, i) Then
                                                SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, i
                                                PlayerReflectNpc index, i, Vital - GetNpcDEF(MapNpc(mapnum).NPC(i).num), 0, spellnum
                                            Else
                                                SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                                SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                            End If
                                        Else
                                            SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                            SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = False
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    
                If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                                        If TempPlayer(i).inParty = TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    Else
                                        SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then Exit Sub
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                Else
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If TempPlayer(i).inParty <> TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                End If
            End Select
            
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            Else
                x = MapNpc(mapnum).NPC(Target).x
                y = MapNpc(mapnum).NPC(Target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
           
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            If Vital > GetPlayerDef(Target) Then
                                If Not CanPlayerDodge(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerReflectPlayer index, Target, Vital - GetPlayerDef(Target), 0, spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital > GetNpcDEF(MapNpc(mapnum).NPC(Target).num) Then
                                If Not CanNpcDodge(mapnum, MapNpc(mapnum).NPC(Target).num, Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerReflectNpc index, Target, Vital - GetNpcDEF(MapNpc(mapnum).NPC(Target).num), 0, spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True  ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = True
                        DidCast = True
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                        End If
                    End If
                    
                    ' Fixed spell type warp attack
                    Case SPELL_TYPE_WARP
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        
                        Select Case GetPlayerDir(Target)
                        
                        Case DIR_UP
                            If Player(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y + 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_DOWN
                            If Player(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y - 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_LEFT
                            If Player(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(Player(Target).x + 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, Player(Target).x + 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                        Case DIR_RIGHT
                            If Player(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x - 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, Player(Target).x - 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                            
                        End Select
                        
                        If CanPlayerAttackPlayerLHand(index, Target, True) Then
                            If Vital > GetPlayerDef(Target) Then
                                If Not CanPlayerDodge(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerReflectPlayer index, Target, Vital - GetPlayerDef(Target), 0, spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                    Else
                 
                        Select Case MapNpc(mapnum).NPC(Target).Dir
                        
                        Case DIR_UP
                            If MapNpc(mapnum).NPC(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y + 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "UP : " & DIR_UP, BrightRed)
                            End If
                        Case DIR_DOWN
                            If MapNpc(mapnum).NPC(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y - 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "DOWN : " & DIR_DOWN, BrightRed)
                            End If
                        Case DIR_LEFT
                            If MapNpc(mapnum).NPC(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x + 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x + 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "LEFT : " & DIR_LEFT, BrightRed)
                            End If
                        Case DIR_RIGHT
                            If MapNpc(mapnum).NPC(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x - 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x - 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "RIGHT : " & DIR_RIGHT, BrightRed)
                            End If
                            
                        End Select
                    
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital > GetNpcDEF(MapNpc(mapnum).NPC(Target).num) Then
                                If Not CanNpcDodge(mapnum, MapNpc(mapnum).NPC(Target).num, Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerReflectNpc index, Target, Vital - GetNpcDEF(MapNpc(mapnum).NPC(Target).num), 0, spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                        
                    End If
                    
            End Select
    End Select
    
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) - HPCost)
        Call SendVital(index, Vitals.HP)
        ' send vitals to party if in one
        
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
       
        TempPlayer(index).SpellCD(spellslot) = GetTickCount + (Spell(spellnum).CDTime * 1000)
        ' ส่งค่า ดีเลย์สกิล
        Call SendCooldown(index, spellslot)
        
        If Not Spell(spellnum).Map = 1 Then
            SendActionMsg mapnum, Trim$(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] " & " !", BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32 + 8
            If Spell(spellnum).Passive > 0 Then
                Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name), BrightGreen)
            End If
        End If
        
        ' แก้ไขบัคสกิลแบบธนู
        If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        ' Spell New type fixed
        If Spell(spellnum).Projectile.Pic > 0 Then
        
        ' Call ProjecTileSpell(index, spellnum)
            
        ' prevent subscript
        If index > MAX_PLAYERS Or index < 1 Then Exit Sub
        
        ' get the players current equipment
        CurEquipment = GetPlayerSpell(index, spellslot)

        ' check if they've got equipment
        If CurEquipment < 1 Or CurEquipment > MAX_SPELLS Then Exit Sub

        ' set the curprojectile
        For i = 1 To MAX_PLAYER_PROJECTILES
            If TempPlayer(index).Projectile(i).Pic = 0 Then
            ' just incase there is left over data
            ClearProjectile index, i
            ' set the curprojtile
            curProjecTile = i
            Exit For
            End If
        Next

        ' check for subscript
        If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

        ' populate the data in the player rec
        With TempPlayer(index).Projectile(curProjecTile)
        .Damage = Vital
        .Direction = GetPlayerDir(index)
        .Pic = Spell(CurEquipment).Projectile.Pic
        .Range = Spell(CurEquipment).Projectile.Range
        .Speed = Spell(CurEquipment).Projectile.Speed
        .x = GetPlayerX(index)
        .y = GetPlayerY(index)
        End With

            ' trololol, they have no more projectile space left
            If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

            ' update the projectile on the map
            SendProjectileToMap index, curProjecTile
        
            End If
        
        ' Send the update
        Call SendStats(index)
        SendPlayerData index
    End If
    
    End If
    
End Sub

' ระบบสกิล V1.0 แบบเวทย์มนต์
Public Sub CastSpellCanMove(ByVal index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal targetType As Byte)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    Dim HPCost As Long
    Dim xt As Long
    Dim yt As Long
    Dim curProjecTile As Long, CurEquipment As Long
    Dim s(1 To 2) As Long
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    Dim Dur As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        'PlayerMsg index, "สกิลยังอยู่ในสถานะดีเลย์ โปรดรออีก ! " & TempPlayer(index).SpellCD(spellslot) / 1000 & " วินาที.", BrightRed
        'Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name) & " กำลังดีเลย์.", BrightRed)
        Exit Sub
    End If
    
    MPCost = Spell(spellnum).MPCost
    HPCost = Spell(spellnum).HPCost
    
    If TempPlayer(index).StunDuration > 0 Then
        Call PlayerMsg(index, "ไม่สามารถใช้สกิลนี้ขณะติดมึนได้", BrightRed)
        Exit Sub
    End If

    If GetPlayerVital(index, Vitals.HP) < HPCost Then
        Call PlayerMsg(index, "ต้องการ Hp " & HPCost & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "ต้องการ Mp " & MPCost & " ในการใช้สกิลนี้", BrightRed)
        Exit Sub
    End If
   'Call PlayerMsg(index, "4", BrightRed)
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "ต้องการเลเวล " & LevelReq & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    AccessReq = Spell(spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "ต้องการ GM ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ClassReq = Spell(spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้สกิลนี้.", BrightRed)
            Exit Sub
        End If
    End If

    ' fixed ! bug of toxin
    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
        If Player(index).BuffStatus(BUFF_TOXIN) = BUFF_TOXIN Then
            Call PlayerMsg(index, "คุณอยู่ในสถานะห้ามฟื้นฟู Hp.", BrightRed)
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
   
   '-------- Spell Damage V 2.0 ---------
   
   ' set damage for spell not pet spell
   If Spell(spellnum).Type <> SPELL_TYPE_PET Then
   
        Vital = Spell(spellnum).Vital + (Spell(spellnum).Vital * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
   End If
   
    ' ระบบสัตว์เลี้ยง 1.0 Vital = Pet number with spell
    If Spell(spellnum).Type = SPELL_TYPE_PET Then

        SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
        Call SpawnPet(index, GetPlayerMap(index), Spell(spellnum).Vital)
        PetFollowOwner index
        DidCast = True

    End If
    
    ' ตั้งค่า Vital for projectile
    If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        Vital = Spell(spellnum).Projectile.Damage + (Spell(spellnum).Projectile.Damage * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
    End If
    
    ' add script mode
    If Spell(spellnum).Type = SPELL_TYPE_SCRIPT Then
        Vital = Spell(spellnum).Vital
    End If
    
    ' -------- End Damage --------
    
    AoE = Spell(spellnum).AoE
    Range = Spell(spellnum).Range
       
    Select Case SpellCastType
        Case 0 ' เลือกเป้าหมายเป็นตัวเอง
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_PET  ' ระบบสัตว์เลี้ยง
                
                Case SPELL_TYPE_SCRIPT
                    ' Script mode
                    Call UseScript(index, Vital, TempPlayer(index).Target, Spell(spellnum).Duration)
                    Call SendAnimation(GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index)
                    DidCast = True ' fixed
                
                Case SPELL_TYPE_HEALHP
                    'Call PlayerMsg(index, "9", BrightRed)
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_PROJECTILE
                    DidCast = True ' <<< Fixed
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    
                    ' ต้องเอาตั้งค่า Dir ออกเพราะบัค ทำให้ไม่วาร์ปตามพิกัดที่คำนวน
                    ' SetPlayerDir index, Spell(spellnum).Dir
                    
                    PlayerWarp index, Spell(spellnum).Map, Spell(spellnum).x, Spell(spellnum).y
                    
        ' สกิลวาร์ปตามแกน เมื่อสกิลเป็นแบบวาร์ปและ แผนที่ = 0,  x คือ พุ่งไปข้างหน้า และ y คือพุ่งถอยหลัง.
        
         If Spell(spellnum).Map = 0 Then
         
         If Player(index).Dir = 0 Then ' Dir Up
         xt = Player(index).x
         yt = Player(index).y - (Spell(spellnum).x) + (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then ' ถ้าไปเจอสิ่งกีดขวาง ทำให้ส่งผลอยู่ที่เดิม
             SetPlayerX index, xt
             SetPlayerY index, yt
             SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 1 Then ' Dir Down
         xt = Player(index).x
         yt = Player(index).y + (Spell(spellnum).x) - (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 2 Then ' Dir Left
         xt = Player(index).x - (Spell(spellnum).x) + (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 3 Then ' Dir right
         xt = Player(index).x + (Spell(spellnum).x) - (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
     End If
                    
            SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
            DidCast = True
            
            End Select
            
        Case 1, 3 ' สกิลแบบ AOE และ AOE แบบรอบเป้าหมาย
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
            
                If targetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                Else
                    x = MapNpc(mapnum).NPC(Target).x
                    y = MapNpc(mapnum).NPC(Target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            If Vital > GetPlayerDef(i) Then
                                                If Not CanPlayerDodge(i) Then
                                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                                    PlayerPassivePlayer index, i, Vital - GetPlayerDef(i), spellnum
                                                Else
                                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                                End If
                                            Else
                                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        If Vital > GetNpcDEF(MapNpc(mapnum).NPC(i).num) Then
                                            If Not CanNpcDodge(mapnum, MapNpc(mapnum).NPC(i).num, i) Then
                                                SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, i
                                                PlayerPassiveNpc index, i, Vital - GetNpcDEF(MapNpc(mapnum).NPC(i).num), spellnum
                                            Else
                                                SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                                SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                            End If
                                        Else
                                            SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                            SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = False
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    
                If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                                        If TempPlayer(i).inParty = TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    Else
                                        SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then Exit Sub
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                Else
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If TempPlayer(i).inParty <> TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                End If
            End Select
            
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            Else
                x = MapNpc(mapnum).NPC(Target).x
                y = MapNpc(mapnum).NPC(Target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
           
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            If Vital > GetPlayerDef(Target) Then
                                If Not CanPlayerDodge(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerPassivePlayer index, Target, Vital - GetPlayerDef(Target), spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital > GetNpcDEF(MapNpc(mapnum).NPC(Target).num) Then
                                If Not CanNpcDodge(mapnum, MapNpc(mapnum).NPC(Target).num, Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerPassiveNpc index, Target, Vital - GetNpcDEF(MapNpc(mapnum).NPC(Target).num), spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True  ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = True
                        DidCast = True
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                        End If
                    End If
                    
                    ' Fixed spell type warp attack
                    Case SPELL_TYPE_WARP
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        
                        Select Case GetPlayerDir(Target)
                        
                        Case DIR_UP
                            If Player(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y + 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_DOWN
                            If Player(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y - 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_LEFT
                            If Player(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(Player(Target).x + 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, Player(Target).x + 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                        Case DIR_RIGHT
                            If Player(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x - 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, Player(Target).x - 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                            
                        End Select
                        
                        If CanPlayerAttackPlayerLHand(index, Target, True) Then
                            If Vital > GetPlayerDef(Target) Then
                                If Not CanPlayerDodge(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerPassivePlayer index, Target, Vital - GetPlayerDef(Target), spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                    
                    Else
                 
                        Select Case MapNpc(mapnum).NPC(Target).Dir
                        
                        Case DIR_UP
                            If MapNpc(mapnum).NPC(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y + 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "UP : " & DIR_UP, BrightRed)
                            End If
                        Case DIR_DOWN
                            If MapNpc(mapnum).NPC(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y - 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "DOWN : " & DIR_DOWN, BrightRed)
                            End If
                        Case DIR_LEFT
                            If MapNpc(mapnum).NPC(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x + 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x + 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "LEFT : " & DIR_LEFT, BrightRed)
                            End If
                        Case DIR_RIGHT
                            If MapNpc(mapnum).NPC(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x - 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x - 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "RIGHT : " & DIR_RIGHT, BrightRed)
                            End If
                            
                        End Select
                    
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital > GetNpcDEF(MapNpc(mapnum).NPC(Target).num) Then
                                If Not CanNpcDodge(mapnum, MapNpc(mapnum).NPC(Target).num, Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerPassiveNpc index, Target, Vital - GetNpcDEF(MapNpc(mapnum).NPC(Target).num), spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                        
                        
                    End If
                    
            End Select
    End Select
    
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) - HPCost)
        Call SendVital(index, Vitals.HP)
        ' send vitals to party if in one
        
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
       
        TempPlayer(index).SpellCD(spellslot) = GetTickCount + (Spell(spellnum).CDTime * 1000)
        ' ส่งค่า ดีเลย์สกิล
        Call SendCooldown(index, spellslot)
        
        If Not Spell(spellnum).Map > 0 Then
            SendActionMsg mapnum, Trim$(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] " & " !", BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32 + 8
            If Spell(spellnum).Passive > 0 Then
                Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] ", BrightGreen)
            End If
        End If
        
        ' แก้ไขบัคสกิลแบบธนู
        If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        ' Spell New type fixed
        If Spell(spellnum).Projectile.Pic > 0 Then
        
        ' Call ProjecTileSpell(index, spellnum)
            
        ' prevent subscript
        If index > MAX_PLAYERS Or index < 1 Then Exit Sub
        
        ' get the players current equipment
        CurEquipment = GetPlayerSpell(index, spellslot)

        ' check if they've got equipment
        If CurEquipment < 1 Or CurEquipment > MAX_SPELLS Then Exit Sub

        ' set the curprojectile
        For i = 1 To MAX_PLAYER_PROJECTILES
            If TempPlayer(index).Projectile(i).Pic = 0 Then
            ' just incase there is left over data
            ClearProjectile index, i
            ' set the curprojtile
            curProjecTile = i
            Exit For
            End If
        Next

        ' check for subscript
        If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

        ' populate the data in the player rec
        With TempPlayer(index).Projectile(curProjecTile)
        .Damage = Vital
        .Direction = GetPlayerDir(index)
        .Pic = Spell(CurEquipment).Projectile.Pic
        .Range = Spell(CurEquipment).Projectile.Range
        .Speed = Spell(CurEquipment).Projectile.Speed
        .x = GetPlayerX(index)
        .y = GetPlayerY(index)
        End With

            ' trololol, they have no more projectile space left
            If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

            ' update the projectile on the map
            SendProjectileToMap index, curProjecTile
        
            End If
        
        ' Send the update
        Call SendStats(index)
        SendPlayerData index
    End If
    
    End If
    
End Sub


' ระบบสกิล V1.0 แบบเวทย์มนต์
Public Sub CastSpellCanMoveLHand(ByVal index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal targetType As Byte)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    Dim HPCost As Long
    Dim xt As Long
    Dim yt As Long
    Dim curProjecTile As Long, CurEquipment As Long
    Dim s(1 To 2) As Long
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    Dim Dur As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        'PlayerMsg index, "สกิลยังอยู่ในสถานะดีเลย์ โปรดรออีก ! " & TempPlayer(index).SpellCD(spellslot) / 1000 & " วินาที.", BrightRed
        'Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name) & " กำลังดีเลย์.", BrightRed)
        Exit Sub
    End If
    
    MPCost = Spell(spellnum).MPCost
    HPCost = Spell(spellnum).HPCost
    
    If TempPlayer(index).StunDuration > 0 Then
        Call PlayerMsg(index, "ไม่สามารถใช้สกิลนี้ขณะติดมึนได้", BrightRed)
        Exit Sub
    End If

    If GetPlayerVital(index, Vitals.HP) < HPCost Then
        Call PlayerMsg(index, "ต้องการ Hp " & HPCost & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "ต้องการ Mp " & MPCost & " ในการใช้สกิลนี้", BrightRed)
        Exit Sub
    End If
   'Call PlayerMsg(index, "4", BrightRed)
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "ต้องการเลเวล " & LevelReq & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    AccessReq = Spell(spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "ต้องการ GM ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ClassReq = Spell(spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้สกิลนี้.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' fixed ! bug of toxin
    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
        If Player(index).BuffStatus(BUFF_TOXIN) = BUFF_TOXIN Then
            Call PlayerMsg(index, "คุณอยู่ในสถานะห้ามฟื้นฟู Hp.", BrightRed)
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
   
   '-------- Spell Damage V 2.0 ---------
   
   ' set damage for spell not pet spell
   If Spell(spellnum).Type <> SPELL_TYPE_PET Then
   
        Vital = Spell(spellnum).Vital + (Spell(spellnum).Vital * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
   End If
   
    ' ระบบสัตว์เลี้ยง 1.0 Vital = Pet number with spell
    If Spell(spellnum).Type = SPELL_TYPE_PET Then

        SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
        Call SpawnPet(index, GetPlayerMap(index), Spell(spellnum).Vital)
        PetFollowOwner index
        DidCast = True

    End If
    
    ' ตั้งค่า Vital for projectile
    If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        Vital = Spell(spellnum).Projectile.Damage + (Spell(spellnum).Projectile.Damage * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
    End If
    
    ' add script mode
    If Spell(spellnum).Type = SPELL_TYPE_SCRIPT Then
        Vital = Spell(spellnum).Vital
    End If
    
    ' -------- End Damage --------
    
    AoE = Spell(spellnum).AoE
    Range = Spell(spellnum).Range
       
    Select Case SpellCastType
        Case 0 ' เลือกเป้าหมายเป็นตัวเอง
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_PET  ' ระบบสัตว์เลี้ยง
                
                Case SPELL_TYPE_SCRIPT
                    ' Script mode
                    Call UseScript(index, Vital, TempPlayer(index).Target, Spell(spellnum).Duration)
                    Call SendAnimation(GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index)
                    DidCast = True ' fixed
                
                Case SPELL_TYPE_HEALHP
                    'Call PlayerMsg(index, "9", BrightRed)
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_PROJECTILE
                    DidCast = True ' <<< Fixed
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    
                    ' ต้องเอาตั้งค่า Dir ออกเพราะบัค ทำให้ไม่วาร์ปตามพิกัดที่คำนวน
                    ' SetPlayerDir index, Spell(spellnum).Dir
                    
                    PlayerWarp index, Spell(spellnum).Map, Spell(spellnum).x, Spell(spellnum).y
                    
        ' สกิลวาร์ปตามแกน เมื่อสกิลเป็นแบบวาร์ปและ แผนที่ = 0,  x คือ พุ่งไปข้างหน้า และ y คือพุ่งถอยหลัง.
        
         If Spell(spellnum).Map = 0 Then
         
         If Player(index).Dir = 0 Then ' Dir Up
         xt = Player(index).x
         yt = Player(index).y - (Spell(spellnum).x) + (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then ' ถ้าไปเจอสิ่งกีดขวาง ทำให้ส่งผลอยู่ที่เดิม
             SetPlayerX index, xt
             SetPlayerY index, yt
             SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 1 Then ' Dir Down
         xt = Player(index).x
         yt = Player(index).y + (Spell(spellnum).x) - (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 2 Then ' Dir Left
         xt = Player(index).x - (Spell(spellnum).x) + (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 3 Then ' Dir right
         xt = Player(index).x + (Spell(spellnum).x) - (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
     End If
                    
            SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
            DidCast = True
            
            End Select
            
        Case 1, 3 ' สกิลแบบ AOE และ AOE แบบรอบเป้าหมาย
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
            
                If targetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                Else
                    x = MapNpc(mapnum).NPC(Target).x
                    y = MapNpc(mapnum).NPC(Target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayerLHand(index, i, True) Then
                                            If Vital > GetPlayerDef(i) Then
                                                If Not CanPlayerDodge(i) Then
                                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                                    PlayerPassivePlayerLHand index, i, Vital - GetPlayerDef(i), spellnum
                                                Else
                                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                                End If
                                            Else
                                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        If Vital > GetNpcDEF(MapNpc(mapnum).NPC(i).num) Then
                                            If Not CanNpcDodge(mapnum, MapNpc(mapnum).NPC(i).num, i) Then
                                                SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, i
                                                PlayerPassiveNpcLHand index, i, Vital - GetNpcDEF(MapNpc(mapnum).NPC(i).num), spellnum
                                            Else
                                                SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                                SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                            End If
                                        Else
                                            SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                            SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = False
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    
                If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                                        If TempPlayer(i).inParty = TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    Else
                                        SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then Exit Sub
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                Else
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If TempPlayer(i).inParty <> TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                End If
            End Select
            
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            Else
                x = MapNpc(mapnum).NPC(Target).x
                y = MapNpc(mapnum).NPC(Target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
           
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            If Vital > GetPlayerDef(Target) Then
                                If Not CanPlayerDodge(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerPassivePlayerLHand index, Target, Vital - GetPlayerDef(Target), spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital > GetNpcDEF(MapNpc(mapnum).NPC(Target).num) Then
                                If Not CanNpcDodge(mapnum, MapNpc(mapnum).NPC(Target).num, Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerPassiveNpcLHand index, Target, Vital - GetNpcDEF(MapNpc(mapnum).NPC(Target).num), spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True  ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = True
                        DidCast = True
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                        End If
                    End If
                    
                    ' Fixed spell type warp attack
                    Case SPELL_TYPE_WARP
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        
                        Select Case GetPlayerDir(Target)
                        
                        Case DIR_UP
                            If Player(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y + 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_DOWN
                            If Player(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y - 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_LEFT
                            If Player(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(Player(Target).x + 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, Player(Target).x + 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                        Case DIR_RIGHT
                            If Player(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x - 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, Player(Target).x - 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                            
                        End Select
                        
                        If CanPlayerAttackPlayerLHand(index, Target, True) Then
                            If Vital > GetPlayerDef(Target) Then
                                If Not CanPlayerDodge(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerPassivePlayerLHand index, Target, Vital - GetPlayerDef(Target), spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                    
                    Else
                 
                        Select Case MapNpc(mapnum).NPC(Target).Dir
                        
                        Case DIR_UP
                            If MapNpc(mapnum).NPC(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y + 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "UP : " & DIR_UP, BrightRed)
                            End If
                        Case DIR_DOWN
                            If MapNpc(mapnum).NPC(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y - 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "DOWN : " & DIR_DOWN, BrightRed)
                            End If
                        Case DIR_LEFT
                            If MapNpc(mapnum).NPC(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x + 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x + 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "LEFT : " & DIR_LEFT, BrightRed)
                            End If
                        Case DIR_RIGHT
                            If MapNpc(mapnum).NPC(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x - 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x - 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "RIGHT : " & DIR_RIGHT, BrightRed)
                            End If
                            
                        End Select
                    
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital > GetNpcDEF(MapNpc(mapnum).NPC(Target).num) Then
                                If Not CanNpcDodge(mapnum, MapNpc(mapnum).NPC(Target).num, Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerPassiveNpcLHand index, Target, Vital - GetNpcDEF(MapNpc(mapnum).NPC(Target).num), spellnum
                                    DidCast = True
                                Else
                                    SendAnimation mapnum, DODGE_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "พลาด !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                    
                    End If
                    
            End Select
    End Select
    
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) - HPCost)
        Call SendVital(index, Vitals.HP)
        ' send vitals to party if in one
        
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
       
        TempPlayer(index).SpellCD(spellslot) = GetTickCount + (Spell(spellnum).CDTime * 1000)
        ' ส่งค่า ดีเลย์สกิล
        Call SendCooldown(index, spellslot)
        
        If Not Spell(spellnum).Map > 0 Then
            SendActionMsg mapnum, Trim$(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] " & " !", BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32 + 8
            If Spell(spellnum).Passive > 0 Then
                Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] ", BrightGreen)
            End If
        End If
        
        ' แก้ไขบัคสกิลแบบธนู
        If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        ' Spell New type fixed
        If Spell(spellnum).Projectile.Pic > 0 Then
        
        ' Call ProjecTileSpell(index, spellnum)
            
        ' prevent subscript
        If index > MAX_PLAYERS Or index < 1 Then Exit Sub
        
        ' get the players current equipment
        CurEquipment = GetPlayerSpell(index, spellslot)

        ' check if they've got equipment
        If CurEquipment < 1 Or CurEquipment > MAX_SPELLS Then Exit Sub

        ' set the curprojectile
        For i = 1 To MAX_PLAYER_PROJECTILES
            If TempPlayer(index).Projectile(i).Pic = 0 Then
            ' just incase there is left over data
            ClearProjectile index, i
            ' set the curprojtile
            curProjecTile = i
            Exit For
            End If
        Next

        ' check for subscript
        If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

        ' populate the data in the player rec
        With TempPlayer(index).Projectile(curProjecTile)
        .Damage = Vital
        .Direction = GetPlayerDir(index)
        .Pic = Spell(CurEquipment).Projectile.Pic
        .Range = Spell(CurEquipment).Projectile.Range
        .Speed = Spell(CurEquipment).Projectile.Speed
        .x = GetPlayerX(index)
        .y = GetPlayerY(index)
        End With

            ' trololol, they have no more projectile space left
            If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

            ' update the projectile on the map
            SendProjectileToMap index, curProjecTile
        
            End If
        
        ' Send the update
        Call SendStats(index)
        SendPlayerData index
    End If
    
    End If
    
End Sub


' ############
' ## Spells ##
' ############

Public Sub BufferPSpell(ByVal index As Long, ByVal spellslot As Long)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    Dim HPCost As Long
    
    Dim targetType As Byte
    Dim Target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)
    
    If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg index, "สกิลยังอยู่ในสถานะดีเลย์ โปรดรออีก ! " & TempPlayer(index).SpellCD(spellslot) / 1000 & " วินาที.", BrightRed
        Exit Sub
    End If

    MPCost = Spell(spellnum).MPCost
    HPCost = Spell(spellnum).HPCost

    If GetPlayerVital(index, Vitals.HP) < HPCost Then
        Call PlayerMsg(index, "ต้องการ Hp " & HPCost & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "ต้องการ Mp " & MPCost & " เพื่อใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "ต้องการเลเวล " & LevelReq & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "ต้องการ GM ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้สกิลนี้.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' fixed ! bug of toxin
    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
        If Player(index).BuffStatus(BUFF_TOXIN) = BUFF_TOXIN Then
            Call PlayerMsg(index, "คุณอยู่ในสถานะห้ามฟื้นฟู Hp.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 2 ' Targetted
        Else
            SpellCastType = 3 ' Targetted AOE
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 0 ' Self-cast
        Else
            SpellCastType = 1 ' Self-cast AOE
        End If
    End If
    
    targetType = TempPlayer(index).targetType
    Target = TempPlayer(index).Target
    Range = Spell(spellnum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                PlayerMsg index, "คุณไม่มีเป้าหมายที่ต้องการใช้สกิล.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg index, "เป้าหมายอยู่ไกลเกินระยะสกิล.", BrightRed
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y) Then
                    PlayerMsg index, "เป้าหมายอยู่ไกลเกินระยะสกิล.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapnum, Spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index

        SendActionMsg mapnum, "กำลังร่าย " & Trim$(Spell(spellnum).Name) & " !", BrightCyan, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        TempPlayer(index).spellBuffer.Spell = spellslot
        TempPlayer(index).spellBuffer.Timer = GetTickCount
        TempPlayer(index).spellBuffer.Target = TempPlayer(index).Target
        TempPlayer(index).spellBuffer.tType = TempPlayer(index).targetType
        
        ' Send the update
        'Call SendStats(Index)
        SendPlayerData index
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If
End Sub

' ระบบสกิลติดตัว V1.0

Public Sub CastSpellPassive(ByVal index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal targetType As Byte)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    Dim HPCost As Long
    Dim xt As Long
    Dim yt As Long
    Dim curProjecTile As Long, CurEquipment As Long
    Dim s(1 To 2) As Long
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    Dim Dur As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        'PlayerMsg index, "สกิลยังอยู่ในสถานะดีเลย์ โปรดรออีก ! " & TempPlayer(index).SpellCD(spellslot) / 1000 & " วินาที.", BrightRed
        'Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name) & " กำลังดีเลย์.", BrightRed)
        Exit Sub
    End If
    
    MPCost = Spell(spellnum).MPCost
    HPCost = Spell(spellnum).HPCost
    
    If TempPlayer(index).StunDuration > 0 Then
        Call PlayerMsg(index, "ไม่สามารถใช้สกิลนี้ขณะติดมึนได้", BrightRed)
        Exit Sub
    End If

    If GetPlayerVital(index, Vitals.HP) < HPCost Then
        Call PlayerMsg(index, "ต้องการ Hp " & HPCost & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "ต้องการ Mp " & MPCost & " ในการใช้สกิลนี้", BrightRed)
        Exit Sub
    End If
   'Call PlayerMsg(index, "4", BrightRed)
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "ต้องการเลเวล " & LevelReq & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    AccessReq = Spell(spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "ต้องการ GM ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ClassReq = Spell(spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้สกิลนี้.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' fixed ! bug of toxin
    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
        If Player(index).BuffStatus(BUFF_TOXIN) = BUFF_TOXIN Then
            Call PlayerMsg(index, "คุณอยู่ในสถานะห้ามฟื้นฟู Hp.", BrightRed)
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
   
   '-------- Spell Damage V 2.0 ---------
   
   ' set damage for spell not pet spell
   If Spell(spellnum).Type <> SPELL_TYPE_PET Then
   
        Vital = Spell(spellnum).Vital + (Spell(spellnum).Vital * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
   End If
   
    ' ระบบสัตว์เลี้ยง 1.0 Vital = Pet number with spell
    If Spell(spellnum).Type = SPELL_TYPE_PET Then

        SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
        Call SpawnPet(index, GetPlayerMap(index), Spell(spellnum).Vital)
        PetFollowOwner index
        DidCast = True

    End If
    
    ' ตั้งค่า Vital for projectile
    If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        Vital = Spell(spellnum).Projectile.Damage + (Spell(spellnum).Projectile.Damage * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
    End If
    
    ' add script mode
    If Spell(spellnum).Type = SPELL_TYPE_SCRIPT Then
        Vital = Spell(spellnum).Vital
    End If
    
    ' -------- End Damage --------
    
    AoE = Spell(spellnum).AoE
    Range = Spell(spellnum).Range
       
    Select Case SpellCastType
        Case 0 ' เลือกเป้าหมายเป็นตัวเอง
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_PET  ' ระบบสัตว์เลี้ยง
                
                Case SPELL_TYPE_SCRIPT
                    ' Script mode
                    Call UseScript(index, Vital, TempPlayer(index).Target, Spell(spellnum).Duration)
                    Call SendAnimation(GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index)
                    DidCast = True ' fixed
                
                Case SPELL_TYPE_HEALHP
                    'Call PlayerMsg(index, "9", BrightRed)
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, spellnum
                    'Call PlayerMsg(index, "Vital = " & Vital & " + " & s(1) & " + " & s(2), Yellow)
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_PROJECTILE
                    DidCast = True ' <<< Fixed
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    
                    ' ต้องเอาตั้งค่า Dir ออกเพราะบัค ทำให้ไม่วาร์ปตามพิกัดที่คำนวน
                    ' SetPlayerDir index, Spell(spellnum).Dir
                    
                    PlayerWarp index, Spell(spellnum).Map, Spell(spellnum).x, Spell(spellnum).y
                    
        ' สกิลวาร์ปตามแกน เมื่อสกิลเป็นแบบวาร์ปและ แผนที่ = 0,  x คือ พุ่งไปข้างหน้า และ y คือพุ่งถอยหลัง.
        
         If Spell(spellnum).Map = 0 Then
         
         If Player(index).Dir = 0 Then ' Dir Up
         xt = Player(index).x
         yt = Player(index).y - (Spell(spellnum).x) + (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then ' ถ้าไปเจอสิ่งกีดขวาง ทำให้ส่งผลอยู่ที่เดิม
             SetPlayerX index, xt
             SetPlayerY index, yt
             SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 1 Then ' Dir Down
         xt = Player(index).x
         yt = Player(index).y + (Spell(spellnum).x) - (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 2 Then ' Dir Left
         xt = Player(index).x - (Spell(spellnum).x) + (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 3 Then ' Dir right
         xt = Player(index).x + (Spell(spellnum).x) - (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
     End If
                    
            SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
            DidCast = True
            
            End Select
            
        Case 1, 3 ' สกิลแบบ AOE และ AOE แบบรอบเป้าหมาย
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
            
                If targetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                Else
                    x = MapNpc(mapnum).NPC(Target).x
                    y = MapNpc(mapnum).NPC(Target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            If Not CanPlayerAbsorbMagic(i) Then
                                                If Vital > GetPlayerMDEF(i) Then
                                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                                    PlayerPassivePlayer index, i, Vital - GetPlayerMDEF(i), spellnum
                                                Else
                                                    SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                    SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                                End If
                                            Else
                                                ' Absorb
                                                SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        If Not CanNpcAbsorbMagic(MapNpc(mapnum).NPC(i).num) Then
                                            If Vital - rand(NPC(i).stat(intelligence), NPC(i).stat(intelligence) * 2) > 0 Then
                                                SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, i
                                                PlayerPassiveNpc index, i, Vital - rand(NPC(i).stat(intelligence), NPC(i).stat(intelligence) * 2), spellnum
                                            Else
                                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                            End If
                                        Else
                                            ' Absorb
                                            SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                            SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = False
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    
                If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                                        If TempPlayer(i).inParty = TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    Else
                                        SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then Exit Sub
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                Else
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If TempPlayer(i).inParty <> TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                End If
            End Select
            
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            Else
                x = MapNpc(mapnum).NPC(Target).x
                y = MapNpc(mapnum).NPC(Target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
           
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            If Vital > GetPlayerMDEF(Target) Then
                                If Not CanPlayerAbsorbMagic(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerPassivePlayer index, Target, Vital - GetPlayerMDEF(Target), spellnum
                                    DidCast = True
                                Else
                                    'Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2) > 0 Then
                                If Not CanNpcAbsorbMagic(MapNpc(mapnum).NPC(Target).num) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerPassiveNpc index, Target, Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2), spellnum
                                    DidCast = True
                                Else
                                    ' Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True  ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = True
                        DidCast = True
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                        End If
                    End If
                    
                    ' Fixed spell type warp attack
                    Case SPELL_TYPE_WARP
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        
                        Select Case GetPlayerDir(Target)
                        
                        Case DIR_UP
                            If Player(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y + 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_DOWN
                            If Player(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y - 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_LEFT
                            If Player(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(Player(Target).x + 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, Player(Target).x + 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                        Case DIR_RIGHT
                            If Player(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x - 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, Player(Target).x - 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                            
                        End Select
                        
                        If CanPlayerAttackPlayerLHand(index, Target, True) Then
                            If Vital > GetPlayerMDEF(Target) Then
                                If Not CanPlayerAbsorbMagic(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerPassivePlayer index, Target, Vital - GetPlayerMDEF(Target), spellnum
                                    DidCast = True
                                Else
                                    'Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                        
                    Else
                    
                        Select Case MapNpc(mapnum).NPC(Target).Dir
                        
                        Case DIR_UP
                            If MapNpc(mapnum).NPC(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y + 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "UP : " & DIR_UP, BrightRed)
                            End If
                        Case DIR_DOWN
                            If MapNpc(mapnum).NPC(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y - 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "DOWN : " & DIR_DOWN, BrightRed)
                            End If
                        Case DIR_LEFT
                            If MapNpc(mapnum).NPC(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x + 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x + 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "LEFT : " & DIR_LEFT, BrightRed)
                            End If
                        Case DIR_RIGHT
                            If MapNpc(mapnum).NPC(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x - 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x - 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "RIGHT : " & DIR_RIGHT, BrightRed)
                            End If
                            
                        End Select
                    
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2) > 0 Then
                                If Not CanNpcAbsorbMagic(MapNpc(mapnum).NPC(Target).num) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerPassiveNpc index, Target, Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2), spellnum
                                    DidCast = True
                                Else
                                    ' Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                    End If
                    
            End Select
    End Select
    
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) - HPCost)
        Call SendVital(index, Vitals.HP)
        ' send vitals to party if in one
        
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
       
        TempPlayer(index).SpellCD(spellslot) = GetTickCount + (Spell(spellnum).CDTime * 1000)
        ' ส่งค่า ดีเลย์สกิล
        Call SendCooldown(index, spellslot)
        
        If Not Spell(spellnum).Map > 0 Then
            SendActionMsg mapnum, Trim$(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] " & " !", BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32 + 8
            If Spell(spellnum).Passive > 0 Then
                Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] ", BrightGreen)
            End If
        End If
        
        ' แก้ไขบัคสกิลแบบธนู
        If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        ' Spell New type fixed
        If Spell(spellnum).Projectile.Pic > 0 Then
        
        ' Call ProjecTileSpell(index, spellnum)
            
        ' prevent subscript
        If index > MAX_PLAYERS Or index < 1 Then Exit Sub
        
        ' get the players current equipment
        CurEquipment = GetPlayerSpell(index, spellslot)

        ' check if they've got equipment
        If CurEquipment < 1 Or CurEquipment > MAX_SPELLS Then Exit Sub

        ' set the curprojectile
        For i = 1 To MAX_PLAYER_PROJECTILES
            If TempPlayer(index).Projectile(i).Pic = 0 Then
            ' just incase there is left over data
            ClearProjectile index, i
            ' set the curprojtile
            curProjecTile = i
            Exit For
            End If
        Next

        ' check for subscript
        If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

        ' populate the data in the player rec
        With TempPlayer(index).Projectile(curProjecTile)
        .Damage = Vital
        .Direction = GetPlayerDir(index)
        .Pic = Spell(CurEquipment).Projectile.Pic
        .Range = Spell(CurEquipment).Projectile.Range
        .Speed = Spell(CurEquipment).Projectile.Speed
        .x = GetPlayerX(index)
        .y = GetPlayerY(index)
        End With

            ' trololol, they have no more projectile space left
            If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

            ' update the projectile on the map
            SendProjectileToMap index, curProjecTile
        
            End If
        
        ' Send the update
        Call SendStats(index)
        SendPlayerData index
    End If
    
    End If
    
End Sub

Public Sub CastSpellPassiveLHand(ByVal index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal targetType As Byte)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    Dim HPCost As Long
    Dim xt As Long
    Dim yt As Long
    Dim curProjecTile As Long, CurEquipment As Long
    Dim s(1 To 2) As Long
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    Dim Dur As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        'PlayerMsg index, "สกิลยังอยู่ในสถานะดีเลย์ โปรดรออีก ! " & TempPlayer(index).SpellCD(spellslot) / 1000 & " วินาที.", BrightRed
        'Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name) & " กำลังดีเลย์.", BrightRed)
        Exit Sub
    End If
    
    MPCost = Spell(spellnum).MPCost
    HPCost = Spell(spellnum).HPCost
    
    If TempPlayer(index).StunDuration > 0 Then
        Call PlayerMsg(index, "ไม่สามารถใช้สกิลนี้ขณะติดมึนได้", BrightRed)
        Exit Sub
    End If

    If GetPlayerVital(index, Vitals.HP) < HPCost Then
        Call PlayerMsg(index, "ต้องการ Hp " & HPCost & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "ต้องการ Mp " & MPCost & " ในการใช้สกิลนี้", BrightRed)
        Exit Sub
    End If
   'Call PlayerMsg(index, "4", BrightRed)
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "ต้องการเลเวล " & LevelReq & " ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    AccessReq = Spell(spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "ต้องการ GM ในการใช้สกิลนี้.", BrightRed)
        Exit Sub
    End If

    ClassReq = Spell(spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "ต้องการอาชีพ " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " ในการใช้สกิลนี้.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' fixed ! bug of toxin
    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
        If Player(index).BuffStatus(BUFF_TOXIN) = BUFF_TOXIN Then
            Call PlayerMsg(index, "คุณอยู่ในสถานะห้ามฟื้นฟู Hp.", BrightRed)
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
   
   '-------- Spell Damage V 2.0 ---------
   
   ' set damage for spell not pet spell
   If Spell(spellnum).Type <> SPELL_TYPE_PET Then
   
        Vital = Spell(spellnum).Vital + (Spell(spellnum).Vital * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
   End If
   
    ' ระบบสัตว์เลี้ยง 1.0 Vital = Pet number with spell
    If Spell(spellnum).Type = SPELL_TYPE_PET Then

        SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
        Call SpawnPet(index, GetPlayerMap(index), Spell(spellnum).Vital)
        PetFollowOwner index
        DidCast = True

    End If
    
    ' ตั้งค่า Vital for projectile
    If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        Vital = Spell(spellnum).Projectile.Damage + (Spell(spellnum).Projectile.Damage * ((Spell(spellnum).S1 * (Player(index).skillLV(spellslot)) / 100)))
        
        If Spell(spellnum).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(spellnum).ATKPer / 100) + (s(1) * (Spell(spellnum).S2 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
        If Spell(spellnum).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(spellnum).MagicPer / 100) + (s(2) * (Spell(spellnum).S3 * (Player(index).skillLV(spellslot) / 100)))
        End If
        
    End If
    
    ' add script mode
    If Spell(spellnum).Type = SPELL_TYPE_SCRIPT Then
        Vital = Spell(spellnum).Vital
    End If
    
    ' -------- End Damage --------
    
    AoE = Spell(spellnum).AoE
    Range = Spell(spellnum).Range
       
    Select Case SpellCastType
        Case 0 ' เลือกเป้าหมายเป็นตัวเอง
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_PET  ' ระบบสัตว์เลี้ยง
                
                Case SPELL_TYPE_SCRIPT
                    ' Script mode
                    Call UseScript(index, Vital, TempPlayer(index).Target, Spell(spellnum).Duration)
                    Call SendAnimation(GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index)
                    DidCast = True ' fixed
                
                Case SPELL_TYPE_HEALHP
                    'Call PlayerMsg(index, "9", BrightRed)
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_PROJECTILE
                    DidCast = True ' <<< Fixed
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    
                    ' ต้องเอาตั้งค่า Dir ออกเพราะบัค ทำให้ไม่วาร์ปตามพิกัดที่คำนวน
                    ' SetPlayerDir index, Spell(spellnum).Dir
                    
                    PlayerWarp index, Spell(spellnum).Map, Spell(spellnum).x, Spell(spellnum).y
                    
        ' สกิลวาร์ปตามแกน เมื่อสกิลเป็นแบบวาร์ปและ แผนที่ = 0,  x คือ พุ่งไปข้างหน้า และ y คือพุ่งถอยหลัง.
        
         If Spell(spellnum).Map = 0 Then
         
         If Player(index).Dir = 0 Then ' Dir Up
         xt = Player(index).x
         yt = Player(index).y - (Spell(spellnum).x) + (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then ' ถ้าไปเจอสิ่งกีดขวาง ทำให้ส่งผลอยู่ที่เดิม
             SetPlayerX index, xt
             SetPlayerY index, yt
             SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 1 Then ' Dir Down
         xt = Player(index).x
         yt = Player(index).y + (Spell(spellnum).x) - (Spell(spellnum).y)
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If yt < Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt + 1
                 Loop
                 If yt < Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If yt > Player(index).y Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 yt = yt - 1
                 Loop
                 If yt > Player(index).y Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 2 Then ' Dir Left
         xt = Player(index).x - (Spell(spellnum).x) + (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
         
         If Player(index).Dir = 3 Then ' Dir right
         xt = Player(index).x + (Spell(spellnum).x) - (Spell(spellnum).y)
         yt = Player(index).y
         If xt > Map(mapnum).MaxX Then xt = Map(mapnum).MaxX
         If yt > Map(mapnum).MaxY Then yt = Map(mapnum).MaxY
         If xt < 1 Then xt = 1
         If yt < 1 Then yt = 1
         If Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE Then
         SetPlayerX index, xt
         SetPlayerY index, yt
         SendPlayerXYToMap index
         Else
             If xt < Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt + 1
                 Loop
                 If xt < Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             Else
             If xt > Player(index).x Then
                 Do Until Map(mapnum).Tile(xt, yt).Type = TILE_TYPE_WALKABLE
                 xt = xt - 1
                 Loop
                 If xt > Player(index).x Then
                 SetPlayerX index, xt
                 SetPlayerY index, yt
                 SendPlayerXYToMap index
                 End If
             End If
             End If
         End If
         End If
     End If
                    
            SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
            DidCast = True
            
            End Select
            
        Case 1, 3 ' สกิลแบบ AOE และ AOE แบบรอบเป้าหมาย
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
            
                If targetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                Else
                    x = MapNpc(mapnum).NPC(Target).x
                    y = MapNpc(mapnum).NPC(Target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayerLHand(index, i, True) Then
                                            If Not CanPlayerAbsorbMagic(i) Then
                                                If Vital > GetPlayerMDEF(i) Then
                                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                                    PlayerPassivePlayerLHand index, i, Vital - GetPlayerMDEF(i), spellnum
                                                Else
                                                    SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                    SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                                End If
                                            Else
                                                ' Absorb
                                                SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, i
                                                SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32) - 16
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If CanPlayerAttackNpcLHand(index, i, True) Then
                                        If Not CanNpcAbsorbMagic(MapNpc(mapnum).NPC(i).num) Then
                                            If Vital - rand(NPC(i).stat(intelligence), NPC(i).stat(intelligence) * 2) > 0 Then
                                                SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, i
                                                PlayerPassiveNpcLHand index, i, Vital - rand(NPC(i).stat(intelligence), NPC(i).stat(intelligence) * 2), spellnum
                                            Else
                                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                            End If
                                        Else
                                            ' Absorb
                                            SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_NPC, i
                                            SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (MapNpc(mapnum).NPC(i).x * 32), (MapNpc(mapnum).NPC(i).y * 32) - 16
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = False
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    
                If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                                        If TempPlayer(i).inParty = TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    Else
                                        SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Or Spell(spellnum).Type = SPELL_TYPE_HEALMP Then Exit Sub
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                Else
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If TempPlayer(i).inParty <> TempPlayer(index).inParty Then
                                            SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                End If
            End Select
            
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            Else
                x = MapNpc(mapnum).NPC(Target).x
                y = MapNpc(mapnum).NPC(Target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "เป้าหมายไม่อยู่ในระยะ.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
           
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            If Vital > GetPlayerMDEF(Target) Then
                                If Not CanPlayerAbsorbMagic(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerPassivePlayerLHand index, Target, Vital - GetPlayerMDEF(Target), spellnum
                                    DidCast = True
                                Else
                                    'Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpcLHand(index, Target, True) Then
                            If Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2) > 0 Then
                                If Not CanNpcAbsorbMagic(MapNpc(mapnum).NPC(Target).num) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerPassiveNpcLHand index, Target, Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2), spellnum
                                    DidCast = True
                                Else
                                    ' Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP, SPELL_TYPE_PROJECTILE
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True  ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
                        'increment = True
                        DidCast = True
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                        End If
                    End If
                    
                    ' Fixed spell type warp attack
                    Case SPELL_TYPE_WARP
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        
                        Select Case GetPlayerDir(Target)
                        
                        Case DIR_UP
                            If Player(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y + 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_DOWN
                            If Player(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x, Player(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, Player(Target).x
                                SetPlayerY index, Player(Target).y - 1
                                SendPlayerXYToMap index
                            End If
                        Case DIR_LEFT
                            If Player(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(Player(Target).x + 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, Player(Target).x + 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                        Case DIR_RIGHT
                            If Player(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(Player(Target).x - 1, Player(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, Player(Target).x - 1
                                SetPlayerY index, Player(Target).y
                                SendPlayerXYToMap index
                            End If
                            
                        End Select
                        
                        If CanPlayerAttackPlayerLHand(index, Target, True) Then
                            If Vital > GetPlayerMDEF(Target) Then
                                If Not CanPlayerAbsorbMagic(Target) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                    PlayerPassivePlayer index, Target, Vital - GetPlayerMDEF(Target), spellnum
                                    DidCast = True
                                Else
                                    'Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_PLAYER, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32) - 16
                                DidCast = True
                            End If
                        End If
                        
                    Else
                    
                        Select Case MapNpc(mapnum).NPC(Target).Dir
                        
                        Case DIR_UP
                            If MapNpc(mapnum).NPC(Target).y + 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y + 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_UP
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y + 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "UP : " & DIR_UP, BrightRed)
                            End If
                        Case DIR_DOWN
                            If MapNpc(mapnum).NPC(Target).y - 1 > Map(mapnum).MaxY Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y - 1).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_DOWN
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y - 1
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "DOWN : " & DIR_DOWN, BrightRed)
                            End If
                        Case DIR_LEFT
                            If MapNpc(mapnum).NPC(Target).x + 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                            
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x + 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_LEFT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x + 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "LEFT : " & DIR_LEFT, BrightRed)
                            End If
                        Case DIR_RIGHT
                            If MapNpc(mapnum).NPC(Target).x - 1 > Map(mapnum).MaxX Then
                                Call PlayerMsg(index, "ตำแหน่งเกินกว่าแผนที่", BrightRed)
                                Exit Sub
                            End If
                        
                            If Not Map(mapnum).Tile(MapNpc(mapnum).NPC(Target).x - 1, MapNpc(mapnum).NPC(Target).y).Type = TILE_TYPE_WALKABLE Then
                                Call PlayerMsg(index, "ไม่สามารถไปยังตำแหน่งที่ต้องการได้", BrightRed)
                                Exit Sub
                            Else
                                Player(index).Dir = DIR_RIGHT
                                SetPlayerX index, MapNpc(mapnum).NPC(Target).x - 1
                                SetPlayerY index, MapNpc(mapnum).NPC(Target).y
                                SendPlayerXYToMap index
                                'Call PlayerMsg(index, "RIGHT : " & DIR_RIGHT, BrightRed)
                            End If
                            
                        End Select
                    
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2) > 0 Then
                                If Not CanNpcAbsorbMagic(MapNpc(mapnum).NPC(Target).num) Then
                                    SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                    PlayerPassiveNpc index, Target, Vital - rand(NPC(Target).stat(intelligence), NPC(Target).stat(intelligence) * 2), spellnum
                                    DidCast = True
                                Else
                                    ' Absorb
                                    SendAnimation mapnum, AbsorbMagic_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                    SendActionMsg mapnum, "ดูดเวทมนต์ !", Pink, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                    DidCast = True
                                End If
                            Else
                                SendAnimation mapnum, PARRY_ANIM, 0, 0, TARGET_TYPE_NPC, Target
                                SendActionMsg mapnum, "เบาไป !", BrightRed, 1, (MapNpc(mapnum).NPC(Target).x * 32), (MapNpc(mapnum).NPC(Target).y * 32) - 16
                                DidCast = True
                            End If
                        End If
                    End If
                    
            End Select
    End Select
    
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) - HPCost)
        Call SendVital(index, Vitals.HP)
        ' send vitals to party if in one
        
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
       
        TempPlayer(index).SpellCD(spellslot) = GetTickCount + (Spell(spellnum).CDTime * 1000)
        ' ส่งค่า ดีเลย์สกิล
        Call SendCooldown(index, spellslot)
        
        If Not Spell(spellnum).Map > 0 Then
            SendActionMsg mapnum, Trim$(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] " & " !", BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32 + 8
            If Spell(spellnum).Passive > 0 Then
                Call PlayerMsg(index, "[สกิลติดตัว] : " & Trim(Spell(spellnum).Name) & " [" & Player(index).skillLV(spellslot) & "] ", BrightGreen)
            End If
        End If
        
        ' แก้ไขบัคสกิลแบบธนู
        If Spell(spellnum).Type = SPELL_TYPE_PROJECTILE Then
        
        ' Spell New type fixed
        If Spell(spellnum).Projectile.Pic > 0 Then
        
        ' Call ProjecTileSpell(index, spellnum)
            
        ' prevent subscript
        If index > MAX_PLAYERS Or index < 1 Then Exit Sub
        
        ' get the players current equipment
        CurEquipment = GetPlayerSpell(index, spellslot)

        ' check if they've got equipment
        If CurEquipment < 1 Or CurEquipment > MAX_SPELLS Then Exit Sub

        ' set the curprojectile
        For i = 1 To MAX_PLAYER_PROJECTILES
            If TempPlayer(index).Projectile(i).Pic = 0 Then
            ' just incase there is left over data
            ClearProjectile index, i
            ' set the curprojtile
            curProjecTile = i
            Exit For
            End If
        Next

        ' check for subscript
        If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

        ' populate the data in the player rec
        With TempPlayer(index).Projectile(curProjecTile)
        .Damage = Vital
        .Direction = GetPlayerDir(index)
        .Pic = Spell(CurEquipment).Projectile.Pic
        .Range = Spell(CurEquipment).Projectile.Range
        .Speed = Spell(CurEquipment).Projectile.Speed
        .x = GetPlayerX(index)
        .y = GetPlayerY(index)
        End With

            ' trololol, they have no more projectile space left
            If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

            ' update the projectile on the map
            SendProjectileToMap index, curProjecTile
        
            End If
        
        ' Send the update
        Call SendStats(index)
        SendPlayerData index
    End If
    
    End If
    
End Sub


Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellnum As Long)
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
    
        SendAnimation GetPlayerMap(index), Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg GetPlayerMap(index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
        
        If increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) + Damage
                If Spell(spellnum).Duration > 0 Then
                    AddHoT_Player index, spellnum
                End If
        ElseIf Not increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) - Damage
        End If
        SendVital index, Vital
    End If

End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellnum As Long, ByVal mapnum As Long, Optional ByVal IsPlayer As Boolean = False)
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
        
                SendAnimation mapnum, Spell(spellnum).spellAnim, 0, 0, TARGET_TYPE_NPC, index
                SendActionMsg mapnum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).x * 32, MapNpc(mapnum).NPC(index).y * 32
                
                ' send the sound
                SendMapSound index, MapNpc(mapnum).NPC(index).x, MapNpc(mapnum).NPC(index).y, SoundEntity.seSpell, spellnum
                
                If increment Then
                        If MapNpc(mapnum).NPC(index).Vital(Vital) + Damage <= GetNpcMaxVital(index, Vitals.HP) Then
                                MapNpc(mapnum).NPC(index).Vital(Vital) = MapNpc(mapnum).NPC(index).Vital(Vital) + Damage
                        Else
                                MapNpc(mapnum).NPC(index).Vital(Vital) = GetNpcMaxVital(index, Vitals.HP)
                        End If
                        
                        If Spell(spellnum).Duration > 0 Then
                                AddHoT_Npc mapnum, index, spellnum
                        End If
                ElseIf Not increment Then
                        MapNpc(mapnum).NPC(index).Vital(Vital) = MapNpc(mapnum).NPC(index).Vital(Vital) - Damage
                End If
                
                ' send update
                SendMapNpcVitals mapnum, index
                
        End If
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal spellnum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
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

Public Sub AddHoT_Player(ByVal index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
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

Public Sub AddDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal spellnum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(index).DoT(i)
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

Public Sub AddHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(index).HoT(i)
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

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
Dim Vital As Long

    With TempPlayer(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
        
        ' set damage for spell not pet spell
        If Spell(.Spell).Type <> SPELL_TYPE_PET Then
   
            Vital = Spell(.Spell).Vital
        
            If Spell(.Spell).PhysicalDmg > 0 Then
                Vital = Vital + ((rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index)) * Spell(.Spell).ATKPer) / 100)
            End If
        
            If Spell(.Spell).MagicDmg > 0 Then
                Vital = Vital + ((rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index)) * Spell(.Spell).MagicPer) / 100)
            End If
        
        End If
      
        ' ตั้งค่า Vital for projectile
        If Spell(.Spell).Type = SPELL_TYPE_PROJECTILE Then
        
            Vital = Spell(.Spell).Projectile.Damage
        
            If Spell(.Spell).PhysicalDmg > 0 Then
                Vital = Vital + ((rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index)) * Spell(.Spell).ATKPer) / 100)
            End If
        
            If Spell(.Spell).MagicDmg > 0 Then
                Vital = Vital + ((rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index)) * Spell(.Spell).MagicPer) / 100)
            End If
        
        End If
    
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, index, True) Then
                    PlayerPassivePlayer .Caster, index, Vital, , Spell(.Spell).spellAnim
                    ' Send the update
                    'SendPlayerData index
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
                        ' Send the update
                        'SendPlayerData index
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
Dim Vital As Long
Dim s(1 To 2) As Long

    With TempPlayer(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
        
        '-------- Spell Damage V 2.0 ---------
   
   ' set damage for spell not pet spell
   If Spell(.Spell).Type <> SPELL_TYPE_PET Then
   
        Vital = Spell(.Spell).Vital
        
        If Spell(.Spell).PhysicalDmg > 0 Then
            s(1) = rand(GetPlayerDamage(index) / 2, GetPlayerDamage(index))
            Vital = Vital + (s(1) * Spell(.Spell).ATKPer / 100)
        End If
        
        If Spell(.Spell).MagicDmg > 0 Then
            s(2) = rand(GetPlayerMATK(index) / 2, GetPlayerMATK(index))
            Vital = Vital + (s(2) * Spell(.Spell).MagicPer / 100)
        End If
        
   End If
       
    ' -------- End Damage --------
        
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg Player(index).Map, "+" & Vital, BrightGreen, ACTIONMSG_SCROLL, Player(index).x * 32, Player(index).y * 32
                If GetPlayerMaxVital(index, HP) > Player(index).Vital(Vitals.HP) + Vital Then
                    Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Vital
                Else
                    Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, HP)
                End If
                SendVital index, Vitals.HP
                .Timer = GetTickCount
                ' Send the update
                'SendPlayerData index
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                        ' Send the update
                        'SendPlayerData index
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal dotNum As Long)
Dim Vital As Long
    
    With MapNpc(mapnum).NPC(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            
        ' set damage for spell not pet spell
        If Spell(.Spell).Type <> SPELL_TYPE_PET Then
   
            Vital = Spell(.Spell).Vital
        
            If Spell(.Spell).PhysicalDmg > 0 Then
                Vital = Vital + ((rand(GetPlayerDamage(.Caster) / 2, GetPlayerDamage(.Caster)) * Spell(.Spell).ATKPer) / 100)
            End If
        
            If Spell(.Spell).MagicDmg > 0 Then
                Vital = Vital + ((rand(GetPlayerMATK(.Caster) / 2, GetPlayerMATK(.Caster)) * Spell(.Spell).MagicPer) / 100)
            End If
        
        End If
      
        ' ตั้งค่า Vital for projectile
        If Spell(.Spell).Type = SPELL_TYPE_PROJECTILE Then
        
            Vital = Spell(.Spell).Projectile.Damage
        
            If Spell(.Spell).PhysicalDmg > 0 Then
                Vital = Vital + ((rand(GetPlayerDamage(.Caster) / 2, GetPlayerDamage(.Caster)) * Spell(.Spell).ATKPer) / 100)
            End If
        
            If Spell(.Spell).MagicDmg > 0 Then
                Vital = Vital + ((rand(GetPlayerMATK(.Caster) / 2, GetPlayerMATK(.Caster)) * Spell(.Spell).MagicPer) / 100)
            End If
        
        End If
            
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, index, True) Then
                    PlayerPassiveNpc .Caster, index, Vital, , True, Spell(.Spell).spellAnim, True
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
                        ' Update
                        'SendPlayerData index
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal hotNum As Long)
Dim Vital As Long
        
        With MapNpc(mapnum).NPC(index).HoT(hotNum)
                If .Used And .Spell > 0 Then
                        
        ' set damage for spell not pet spell
        If Spell(.Spell).Type <> SPELL_TYPE_PET Then
   
            Vital = Spell(.Spell).Vital
        
            If Spell(.Spell).PhysicalDmg > 0 Then
                Vital = Vital + ((rand(NPC(index).Damage / 2, NPC(index).Damage) * Spell(.Spell).ATKPer) / 100)
            End If
        
            If Spell(.Spell).MagicDmg > 0 Then
                Vital = Vital + ((rand(NPC(index).MATK / 2, NPC(index).MATK) * Spell(.Spell).MagicPer) / 100)
            End If
        
        End If
      
        ' ตั้งค่า Vital for projectile
        If Spell(.Spell).Type = SPELL_TYPE_PROJECTILE Then
        
            Vital = Spell(.Spell).Projectile.Damage
        
            If Spell(.Spell).PhysicalDmg > 0 Then
                Vital = Vital + ((rand(NPC(index).Damage / 2, NPC(index).Damage) * Spell(.Spell).ATKPer) / 100)
            End If
        
            If Spell(.Spell).MagicDmg > 0 Then
                Vital = Vital + ((rand(NPC(index).MATK / 2, NPC(index).MATK) * Spell(.Spell).MagicPer) / 100)
            End If
        
        End If
                        
                        ' time to tick?
                        If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                                    SendActionMsg mapnum, "+" & Vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).x * 32, MapNpc(mapnum).NPC(index).y * 32
                                    MapNpc(mapnum).NPC(index).Vital(Vitals.HP) = MapNpc(mapnum).NPC(index).Vital(Vitals.HP) + Vital
                                        
                                    If MapNpc(mapnum).NPC(index).Vital(Vitals.HP) > GetNpcMaxVital(index, Vitals.HP) Then
                                        MapNpc(mapnum).NPC(index).Vital(Vitals.HP) = GetNpcMaxVital(index, Vitals.HP)
                                    End If
                                        
                                    SendMapNpcVitals mapnum, index
                                Else
                                    SendActionMsg mapnum, "+" & Vital, BrightBlue, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).x * 32, MapNpc(mapnum).NPC(index).y * 32
                                    MapNpc(mapnum).NPC(index).Vital(Vitals.MP) = MapNpc(mapnum).NPC(index).Vital(Vitals.MP) + Vital
                                        
                                    If MapNpc(mapnum).NPC(index).Vital(Vitals.MP) > GetNpcMaxVital(index, Vitals.MP) Then
                                        MapNpc(mapnum).NPC(index).Vital(Vitals.MP) = GetNpcMaxVital(index, Vitals.MP)
                                    End If
                                        
                                    SendMapNpcVitals mapnum, index
                                End If
                                
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
                                                ' Update
                                                'SendPlayerData index
                                        End If
                                End If
                        End If
                End If
        End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(index).StunDuration = Spell(spellnum).StunDuration
        TempPlayer(index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned index
        ' tell him he's stunned
        PlayerMsg index, "คุณกำลังถูกสตั้นอยู่.", BrightRed
    End If
End Sub

Public Sub StunPlayerP(ByVal index As Long, ByVal StunTime As Long)
        ' set the values on index
        TempPlayer(index).StunDuration = StunTime
        TempPlayer(index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned index
        ' tell him he's stunned
        PlayerMsg index, "คุณกำลังถูกสตั้นอยู่.", BrightRed
End Sub

Public Sub StunNPC(ByVal index As Long, ByVal mapnum As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(mapnum).NPC(index).StunDuration = Spell(spellnum).StunDuration
        MapNpc(mapnum).NPC(index).StunTimer = GetTickCount
    End If
End Sub

Public Sub StunNPCP(ByVal index As Long, ByVal mapnum As Long, ByVal StunTime As Long)
        ' set the values on index
        MapNpc(mapnum).NPC(index).StunDuration = StunTime
        MapNpc(mapnum).NPC(index).StunTimer = GetTickCount
End Sub

Function CanNpcAttackNpc(ByVal mapnum As Long, ByVal Attacker As Long, ByVal Victim As Long) As Boolean
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
    If MapNpc(mapnum).NPC(Attacker).num <= 0 Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(mapnum).NPC(Victim).num <= 0 Then
        Exit Function
    End If

    aNpcNum = MapNpc(mapnum).NPC(Attacker).num
    vNpcNum = MapNpc(mapnum).NPC(Victim).num
    
    If aNpcNum <= 0 Then Exit Function
    If vNpcNum <= 0 Then Exit Function
    
    If Map(mapnum).Moral <> MAP_MORAL_PETARENA Then
    If MapNpc(mapnum).NPC(Attacker).IsPet = YES Then
    If MapNpc(mapnum).NPC(Victim).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(Attacker).PetData.Owner
        If Not Map(GetPlayerMap(petowner)).Moral = MAP_MORAL_NONE Then
            Call PlayerMsg(petowner, "ที่นี่คือเขตปลอดภัย ไม่สามารถใช้สัตว์เลี้ยงโจมตีผู้เล่นได้.", BrightRed)
            Exit Function
        End If
    End If
    End If
    
    If MapNpc(mapnum).NPC(Attacker).IsPet = YES Then
    If MapNpc(mapnum).NPC(Victim).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(Attacker).PetData.Owner
        If GetPlayerAccess(petowner) > ADMIN_MONITOR Then
            Call PlayerMsg(petowner, "Gm ไม่สามารถโจมตีผู้เล่นด้วยสัตว์เลี้ยงได้.", BrightBlue)
            Exit Function
        End If
    End If
    End If
    
    If MapNpc(mapnum).NPC(Attacker).StunDuration > 0 Then
        'SendActionMsg mapnum, "Stun!", White, 1, (MapNpc(mapnum).NPC(Attacker).x * 32), (MapNpc(mapnum).NPC(Attacker).y * 32)
        Exit Function
    End If

    If MapNpc(mapnum).NPC(Attacker).IsPet = YES Then
    If MapNpc(mapnum).NPC(Victim).IsPet = YES Then
    petowner = MapNpc(mapnum).NPC(Victim).PetData.Owner
        If GetPlayerAccess(petowner) > ADMIN_MONITOR Then
            Call PlayerMsg(petowner, "คุณไม่สามารถใช้สัตว์เลี้ยงโจมตี " & GetPlayerName(petowner) & "[GM] ได้.", BrightBlue)
            CanNpcAttackNpc = False
            Exit Function
        End If
    End If
    End If
    End If

    If MapNpc(mapnum).NPC(Attacker).IsPet = YES Then
    If MapNpc(mapnum).NPC(Victim).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(Attacker).PetData.Owner
        If GetPlayerLevel(petowner) < 10 Then
            Call PlayerMsg(petowner, "คุณมีเลเวลต่ำกว่า 10, ไม่สามารถใช้สัตว์เลี้ยงโจมตีผู้เล่นอื่นได้ !", BrightRed)
            CanNpcAttackNpc = False
            Exit Function
        End If
    End If
    End If
    
    If MapNpc(mapnum).NPC(Attacker).IsPet = YES Then
    If MapNpc(mapnum).NPC(Victim).IsPet = YES Then
        petowner = MapNpc(mapnum).NPC(Victim).PetData.Owner
        If GetPlayerLevel(petowner) < 10 Then
            Call PlayerMsg(petowner, GetPlayerName(petowner) & " มีเลเวลต่ำกว่า 10, คุณไม่สามารถใช้สัตว์เลี้ยงโจมตีเขาได้ !", BrightRed)
            Exit Function
        End If
    End If
    End If
    
    
    ' Make sure the npcs arent already dead
    If MapNpc(mapnum).NPC(Attacker).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(Victim).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Npc aspd
    If NPC(MapNpc(mapnum).NPC(Attacker).num).AttackSpeed > 0 Then
        If GetTickCount < MapNpc(mapnum).NPC(Attacker).AttackTimer + NPC(MapNpc(mapnum).NPC(Attacker).num).AttackSpeed Then
            Exit Function
        End If
    Else
        If GetTickCount < MapNpc(mapnum).NPC(Attacker).AttackTimer + 100 Then
            Exit Function
        End If
    End If
    
    MapNpc(mapnum).NPC(Attacker).AttackTimer = GetTickCount
    
    AttackerX = MapNpc(mapnum).NPC(Attacker).x
    AttackerY = MapNpc(mapnum).NPC(Attacker).y
    VictimX = MapNpc(mapnum).NPC(Victim).x
    VictimY = MapNpc(mapnum).NPC(Victim).y

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

Sub NpcAttackNpc(ByVal mapnum As Long, ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim aNpcNum As Long
    Dim vNpcNum As Long
    Dim n As Long
    Dim petowner As Long
    
    If Attacker <= 0 Or Attacker > MAX_MAP_NPCS Then Exit Sub
    If Victim <= 0 Or Victim > MAX_MAP_NPCS Then Exit Sub
    
    If Damage <= 0 Then Exit Sub
    
    aNpcNum = MapNpc(mapnum).NPC(Attacker).num
    vNpcNum = MapNpc(mapnum).NPC(Victim).num
    
    If aNpcNum <= 0 Then Exit Sub
    If vNpcNum <= 0 Then Exit Sub
    
    'set the victim's target to the pet attacking it
    MapNpc(mapnum).NPC(Victim).targetType = 2 'Npc
    MapNpc(mapnum).NPC(Victim).Target = Attacker
    
    ' Send this packet so they can see the person attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong Attacker
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing

    If Damage >= MapNpc(mapnum).NPC(Victim).Vital(Vitals.HP) Then
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).NPC(Victim).x * 32), (MapNpc(mapnum).NPC(Victim).y * 32)
        SendBlood mapnum, MapNpc(mapnum).NPC(Victim).x, MapNpc(mapnum).NPC(Victim).y
        
        ' npc is dead.
        'Call GlobalMsg(CheckGrammar(Trim$(Npc(vNpcNum).Name), 1) & " has been killed by " & CheckGrammar(Trim$(Npc(aNpcNum).Name)) & "!", BrightRed)

        ' Set NPC target to 0
        MapNpc(mapnum).NPC(Attacker).Target = 0
        MapNpc(mapnum).NPC(Attacker).targetType = 0
        'reset the targetter for the player
        
        If MapNpc(mapnum).NPC(Attacker).IsPet = YES Then
            TempPlayer(MapNpc(mapnum).NPC(Attacker).PetData.Owner).Target = 0
            TempPlayer(MapNpc(mapnum).NPC(Attacker).PetData.Owner).targetType = TARGET_TYPE_NONE
            
            petowner = MapNpc(mapnum).NPC(Attacker).PetData.Owner
            
            SendTarget petowner
            
            'Give the player the pet owner some experience from the kill
            Call SetPlayerExp(petowner, GetPlayerExp(petowner) + NPC(MapNpc(mapnum).NPC(Victim).num).exp)
            CheckPlayerLevelUp petowner
            SendActionMsg mapnum, "+" & NPC(MapNpc(mapnum).NPC(Victim).num).exp & "Exp", White, 1, GetPlayerX(petowner) * 32, GetPlayerY(petowner) * 32
            SendEXP petowner
                      
        ElseIf MapNpc(mapnum).NPC(Victim).IsPet = YES Then
        
            'Set the NPC's target on the pet now
            MapNpc(mapnum).NPC(Attacker).targetType = 2 'npc
            MapNpc(mapnum).NPC(Attacker).Target = Attacker
            'Disband the pet
            PetDisband petowner, GetPlayerMap(petowner)
            
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = GetPlayerMap(petowner) Then
                    Call PlayerWarp(i, GetPlayerMap(petowner), GetPlayerX(i), GetPlayerY(i))
                End If
            End If
        Next
            
        End If
        
        ' Reset victim's stuff so it dies in loop
        MapNpc(mapnum).NPC(Victim).num = 0
        MapNpc(mapnum).NPC(Victim).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(Victim).Vital(Vitals.HP) = 0
               
        ' send npc death packet to map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong Victim
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        If petowner > 0 Then
            PetFollowOwner petowner
        End If
    Else
        ' npc not dead, just do the damage
        MapNpc(mapnum).NPC(Victim).Vital(Vitals.HP) = MapNpc(mapnum).NPC(Victim).Vital(Vitals.HP) - Damage
        
        ' ส่งเสียงของ npc ไปยังแผนที่
        SendMapSound Attacker, MapNpc(mapnum).NPC(Victim).x, MapNpc(mapnum).NPC(Victim).y, SoundEntity.seNpc, aNpcNum
        
        ' ถ้า npc ไม่ตาย จะวาดอนิเมชั่น
        Call SendAnimation(mapnum, NPC(aNpcNum).Animation, MapNpc(mapnum).NPC(Victim).x, MapNpc(mapnum).NPC(Victim).y, 0, 0)
       
        ' Say damage
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).NPC(Victim).x * 32), (MapNpc(mapnum).NPC(Victim).y * 32)
        SendBlood mapnum, MapNpc(mapnum).NPC(Victim).x, MapNpc(mapnum).NPC(Victim).y
    End If
    
    'Send both Npc's Vitals to the client
    SendMapNpcVitals mapnum, Attacker
    SendMapNpcVitals mapnum, Victim

End Sub

Public Sub ProjecTileSpell(ByVal index As Long, ByVal spellslot As Long)
Dim curProjecTile As Long, i As Long, CurEquipment As Long

' prevent subscript
If index > MAX_PLAYERS Or index < 1 Then Exit Sub

' get the players current equipment
CurEquipment = GetPlayerSpell(index, spellslot)

' check if they've got equipment
If CurEquipment < 1 Or CurEquipment > MAX_SPELLS Then Exit Sub

' set the curprojectile
For i = 1 To MAX_PLAYER_PROJECTILES
If TempPlayer(index).Projectile(i).Pic = 0 Then
' just incase there is left over data
ClearProjectile index, i
' set the curprojtile
curProjecTile = i
Exit For
End If
Next

' check for subscript
If curProjecTile < 1 Then Exit Sub

' populate the data in the player rec
With TempPlayer(index).Projectile(curProjecTile)
.Damage = Spell(CurEquipment).Projectile.Damage
.Direction = GetPlayerDir(index)
.Pic = Spell(CurEquipment).Projectile.Pic
.Range = Spell(CurEquipment).Projectile.Range
.Speed = Spell(CurEquipment).Projectile.Speed
.x = GetPlayerX(index)
.y = GetPlayerY(index)
End With


If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub

' update the projectile on the map
SendProjectileToMap index, curProjecTile

End Sub

Private Sub NpcWarp(ByVal mapNpcNum As Long, ByVal PlayerX As Long, ByVal PlayerY As Long, ByVal Dir As Long, ByVal mapnum As Long)
Dim Buffer As clsBuffer

' set npc
MapNpc(mapnum).NPC(mapNpcNum).x = PlayerX
MapNpc(mapnum).NPC(mapNpcNum).y = PlayerY
MapNpc(mapnum).NPC(mapNpcNum).Dir = Dir

'Set Buffer = New clsBuffer
'Buffer.WriteLong SNpcWarp
'Buffer.WriteLong mapNpcNum
'Buffer.WriteLong PlayerX
'Buffer.WriteLong PlayerY
'Buffer.WriteLong Dir
'Buffer.WriteLong mapnum
'SendDataToMap mapnum, Buffer.ToArray()

Set Buffer = New clsBuffer
Buffer.WriteLong SNpcMove
Buffer.WriteLong mapNpcNum
Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
Buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
Buffer.WriteLong MOVING_WALKING
SendDataToMap mapnum, Buffer.ToArray()
Set Buffer = Nothing

End Sub
