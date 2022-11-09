Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GetAddress = FunAddr
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub InitMessages()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerBuff) = GetAddress(AddressOf HandlePlayerBuff)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    'Events
    HandleDataSub(SSpawnEvent) = GetAddress(AddressOf HandleSpawnEventPage)
    HandleDataSub(SEventMove) = GetAddress(AddressOf HandleEventMove)
    HandleDataSub(SEventDir) = GetAddress(AddressOf HandleEventDir)
    HandleDataSub(SEventChat) = GetAddress(AddressOf HandleEventChat)
    
    HandleDataSub(SEventStart) = GetAddress(AddressOf HandleEventStart)
    HandleDataSub(SEventEnd) = GetAddress(AddressOf HandleEventEnd)
    
    HandleDataSub(SPlayBGM) = GetAddress(AddressOf HandlePlayBGM)
    HandleDataSub(SPlaySound) = GetAddress(AddressOf HandlePlaySound)
    HandleDataSub(SFadeoutBGM) = GetAddress(AddressOf HandleFadeoutBGM)
    HandleDataSub(SStopSound) = GetAddress(AddressOf HandleStopSound)
    HandleDataSub(SSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    
    HandleDataSub(SMapEventData) = GetAddress(AddressOf HandleMapEventData)
    ' quests
    HandleDataSub(SQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(SUpdateQuest) = GetAddress(AddressOf HandleUpdateQuest)
    HandleDataSub(SPlayerQuest) = GetAddress(AddressOf HandlePlayerQuest)
    HandleDataSub(SQuestMessage) = GetAddress(AddressOf HandleQuestMessage)
    ' projectile
    HandleDataSub(SHandleProjectile) = GetAddress(AddressOf HandleProjectile)
    ' guilds
    HandleDataSub(SSendGuild) = GetAddress(AddressOf HandleSendGuild)
    HandleDataSub(SAdminGuild) = GetAddress(AddressOf HandleAdminGuild)
    ' pets
    HandleDataSub(SNPCCache) = GetAddress(AddressOf HandleNPCCache)
    'doors
    HandleDataSub(SDoorsEditor) = GetAddress(AddressOf HandleDoorsEditor)
    HandleDataSub(SUpdateDoors) = GetAddress(AddressOf HandleUpdateDoors)
    HandleDataSub(SSpell) = GetAddress(AddressOf HandleSpell)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), MyIndex, Buffer.ReadBytes(Buffer.Length), 0, 0
    
    If Not frmOnline.Visible = True Then
    
    frmOnline.lstOnline.Clear
    
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                frmOnline.lstOnline.AddItem Trim$(Player(i).Name)
            End If
        Next i
    
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Sub HandleProjectile(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PlayerProjectile As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' create a new instance of the buffer
    Set Buffer = New clsBuffer
    
    ' read bytes from data()
    Buffer.WriteBytes Data()
    
    ' recieve projectile number
    PlayerProjectile = Buffer.ReadLong
    Index = Buffer.ReadLong
    
    ' populate the values
    With Player(Index).ProjecTile(PlayerProjectile)
    
        ' set the direction
        .Direction = Buffer.ReadLong
        
        ' set the direction to support file format
        Select Case .Direction
            Case DIR_DOWN
                .Direction = 0
            Case DIR_UP
                .Direction = 1
            Case DIR_RIGHT
                .Direction = 2
            Case DIR_LEFT
                .Direction = 3
        End Select
        
        ' set the pic
        .Pic = Buffer.ReadLong
        ' set the coordinates
        .X = GetPlayerX(Index)
        .Y = GetPlayerY(Index)
        ' get the range
        .Range = Buffer.ReadLong
        ' get the damge
        .Damage = Buffer.ReadLong
        ' get the speed
        .Speed = Buffer.ReadLong
        
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    frmLoad.Visible = False
    frmMenu.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    frmMenu.picMain.Visible = True
    
    Msg = Buffer.ReadString 'Parse(1)
    
    Set Buffer = Nothing
    Call MsgBox(Msg, vbOKOnly, Options.Game_Name)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' save options
    Options.SavePass = frmMenu.chkPass.Value
    Options.Username = Trim$(frmMenu.txtLUser.text)

    If frmMenu.chkPass.Value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMenu.txtLPass.text)
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    
    ' player high index
    Player_HighIndex = Buffer.ReadLong
    
    Set Buffer = Nothing
    frmLoad.Visible = True
    Call SetStatus("กำลังรับข้อมูลเกม..")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Z As Long, X As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString
            .Vital(Vitals.HP) = Buffer.ReadLong
            .Vital(Vitals.MP) = Buffer.ReadLong
            
            ' get array size
            Z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .MaleSprite(X) = Buffer.ReadLong
            Next
            
            ' get array size
            Z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .FemaleSprite(X) = Buffer.ReadLong
            Next
            
            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = Buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set Buffer = Nothing
    
    ' Used for if the player is creating a new character
    frmMenu.Visible = True
    frmMenu.picCharacter.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picRegister.Visible = False
    frmLoad.Visible = False
    frmMenu.cmbClass.Clear
    For i = 1 To Max_Classes
        frmMenu.cmbClass.AddItem Trim$(Class(i).Name)
    Next

    frmMenu.cmbClass.ListIndex = 0
    n = frmMenu.cmbClass.ListIndex + 1
    
    newCharSprite = 0
    NewCharacterBltSprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNewCharClasses", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Z As Long, X As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = Buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = Buffer.ReadLong 'CLng(Parse(n + 2))
            
            ' get array size
            Z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .MaleSprite(X) = Buffer.ReadLong
            Next
            
            ' get array size
            Z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .FemaleSprite(X) = Buffer.ReadLong
            Next
                            
            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = Buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InGame = True
    Call GameInit
    Call GameLoop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
        n = n + 2
    Next
    
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set Buffer = Nothing
    BltInventory
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(3)))
    ' changes, clear drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    BltInventory
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Shield)
    
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    BltInventory
    BltEquipment
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim playerNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    playerNum = Buffer.ReadLong
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Shield)
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.HP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, Buffer.ReadLong)

    If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
        'frmMain.lblHP.Caption = Int(GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP) * 100) & "%"
        frmMain.lblHP.Caption = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
        ' hp bar
        frmMain.imgHPBar.Width = ((GetPlayerVital(MyIndex, Vitals.HP) / HPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPBar_Width)) * HPBar_Width
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerHP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.MP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, Buffer.ReadLong)

    If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
        'frmMain.lblMP.Caption = Int(GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP) * 100) & "%"
        frmMain.lblMP.Caption = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
        ' mp bar
        frmMain.imgMPBar.Width = ((GetPlayerVital(MyIndex, Vitals.MP) / SPRBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / SPRBar_Width)) * SPRBar_Width
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim str, agi, ends, ints, will As Long
    
    str = 0
    agi = 0
    ends = 0
    ints = 0
    will = 0

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' ------------------------ สูตรคำนวนแบบใหม่ !! -----------------------------
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                str = str + Item(GetPlayerEquipment(MyIndex, Weapon)).Add_Stat(1) + (Item(GetPlayerEquipment(MyIndex, Weapon)).Rarity * 5)
                ends = ends + Item(GetPlayerEquipment(MyIndex, Weapon)).Add_Stat(2) + (Item(GetPlayerEquipment(MyIndex, Weapon)).Rarity * 5)
                ints = ints + Item(GetPlayerEquipment(MyIndex, Weapon)).Add_Stat(3) + (Item(GetPlayerEquipment(MyIndex, Weapon)).Rarity * 5)
                agi = agi + Item(GetPlayerEquipment(MyIndex, Weapon)).Add_Stat(4) + (Item(GetPlayerEquipment(MyIndex, Weapon)).Rarity * 5)
                will = will + Item(GetPlayerEquipment(MyIndex, Weapon)).Add_Stat(5) + (Item(GetPlayerEquipment(MyIndex, Weapon)).Rarity * 5)
            End If
        
        If GetPlayerEquipment(MyIndex, Armor) > 0 Then
            str = str + Item(GetPlayerEquipment(MyIndex, Armor)).Add_Stat(1) + (Item(GetPlayerEquipment(MyIndex, Armor)).Rarity * 5)
            ends = ends + Item(GetPlayerEquipment(MyIndex, Armor)).Add_Stat(2) + (Item(GetPlayerEquipment(MyIndex, Armor)).Rarity * 5)
            ints = ints + Item(GetPlayerEquipment(MyIndex, Armor)).Add_Stat(3) + (Item(GetPlayerEquipment(MyIndex, Armor)).Rarity * 5)
            agi = agi + Item(GetPlayerEquipment(MyIndex, Armor)).Add_Stat(4) + (Item(GetPlayerEquipment(MyIndex, Armor)).Rarity * 5)
            will = will + Item(GetPlayerEquipment(MyIndex, Armor)).Add_Stat(5) + (Item(GetPlayerEquipment(MyIndex, Armor)).Rarity * 5)
        End If
        
        If GetPlayerEquipment(MyIndex, Shield) > 0 Then
            str = str + Item(GetPlayerEquipment(MyIndex, Shield)).Add_Stat(1) + (Item(GetPlayerEquipment(MyIndex, Shield)).Rarity * 5)
            ends = ends + Item(GetPlayerEquipment(MyIndex, Shield)).Add_Stat(2) + (Item(GetPlayerEquipment(MyIndex, Shield)).Rarity * 5)
            ints = ints + Item(GetPlayerEquipment(MyIndex, Shield)).Add_Stat(3) + (Item(GetPlayerEquipment(MyIndex, Shield)).Rarity * 5)
            agi = agi + Item(GetPlayerEquipment(MyIndex, Shield)).Add_Stat(4) + (Item(GetPlayerEquipment(MyIndex, Shield)).Rarity * 5)
            will = will + Item(GetPlayerEquipment(MyIndex, Shield)).Add_Stat(5) + (Item(GetPlayerEquipment(MyIndex, Shield)).Rarity * 5)
        End If
        
        If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
            str = str + Item(GetPlayerEquipment(MyIndex, Helmet)).Add_Stat(1) + (Item(GetPlayerEquipment(MyIndex, Helmet)).Rarity * 5)
            ends = ends + Item(GetPlayerEquipment(MyIndex, Helmet)).Add_Stat(2) + (Item(GetPlayerEquipment(MyIndex, Helmet)).Rarity * 5)
            ints = ints + Item(GetPlayerEquipment(MyIndex, Helmet)).Add_Stat(3) + (Item(GetPlayerEquipment(MyIndex, Helmet)).Rarity * 5)
            agi = agi + Item(GetPlayerEquipment(MyIndex, Helmet)).Add_Stat(4) + (Item(GetPlayerEquipment(MyIndex, Helmet)).Rarity * 5)
            will = will + Item(GetPlayerEquipment(MyIndex, Helmet)).Add_Stat(5) + (Item(GetPlayerEquipment(MyIndex, Helmet)).Rarity * 5)
        End If
    
    For i = 1 To Stats.Stat_Count - 1
        Select Case (i)
            Case 1: SetPlayerStat Index, i, Buffer.ReadLong
                            frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i) - str & "(+" & str & ")"
             Case 2: SetPlayerStat Index, i, Buffer.ReadLong
                            frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i) - ends & "(+" & ends & ")"
             Case 3: SetPlayerStat Index, i, Buffer.ReadLong
                            frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i) - ints & "(+" & ints & ")"
             Case 4: SetPlayerStat Index, i, Buffer.ReadLong
                            frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i) - agi & "(+" & agi & ")"
             Case 5: SetPlayerStat Index, i, Buffer.ReadLong
                            frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i) - will & "(+" & will & ")"
          End Select
      Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim TNL As Long
Dim LabelExp
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SetPlayerExp MyIndex, Buffer.ReadLong
    TNL = Buffer.ReadLong
    ' แก้ไขรูปแบบคำสั่ง FormatNumber(ตัวแปร,หลักของทศนิยม)

    LabelExp = (GetPlayerExp(MyIndex) / TNL) * 100  ' ((GetPlayerExp(MyIndex) * 100) / TNL)
    
    If LabelExp >= 0 And LabelExp < 1 Then
        frmMain.lblEXP.Caption = "0" & FormatNumber(LabelExp, 4) & "%" ' แสดง Exp เป็น %
    Else
        frmMain.lblEXP.Caption = FormatNumber(LabelExp, 4) & "%" ' แสดง Exp เป็น %
    End If
    
    frmMain.lblExpShow.Caption = GetPlayerExp(MyIndex) & " / " & TNL
    'frmMain.lblEXP.Caption = GetPlayerExp(MyIndex) & "/" & TNL ' แสดง Exp แบบเป็นหน่วย
    ' mp bar
    frmMain.imgEXPBar.Width = ((GetPlayerExp(MyIndex) / EXPBar_Width) / (TNL / EXPBar_Width)) * EXPBar_Width
    frmMain.imgEXPBar2.Width = ((GetPlayerExp(MyIndex) / EXPBar_Width2) / (TNL / EXPBar_Width2)) * EXPBar_Width2
    
    'If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
    '    frmMain.lblStr.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).Data2 + (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)
    'Else
    '    frmMain.lblStr.Caption = (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)
    'End If
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerExp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long, X As Long
Dim Buffer As clsBuffer
Dim str, agi, ends, ints, will As Long
Dim Class As String
    
    str = 0
    agi = 0
    ends = 0
    ints = 0
    will = 0

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerLevel(i, Buffer.ReadLong)
    Call SetPlayerPOINTS(i, Buffer.ReadLong)
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    
    'Load the Players Message
    Player(i).Message = Buffer.ReadString

    Call SetPlayerClass(i, Buffer.ReadLong)
    
    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, Buffer.ReadLong
    Next
    
    If Buffer.ReadByte = 1 Then
        Player(i).GuildName = Buffer.ReadString
    Else
        Player(i).GuildName = vbNullString
    End If
    
    For X = 1 To MAX_PLAYER_SPELLS
        Player(i).skillLV(X) = Buffer.ReadByte
    Next
    
    For X = 1 To MAX_PLAYER_SPELLS
        Player(i).skillEXP(X) = Buffer.ReadLong
    Next

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
        
    For X = 1 To MAX_PLAYER_SPELLS
        skillLV(X) = Player(i).skillLV(X)
    Next
    
    For X = 1 To MAX_PLAYER_SPELLS
        skillEXP(X) = Player(i).skillEXP(X)
    Next
        
        ' Set the character windows ' และอาชีพ
        frmMain.lblCharName = GetPlayerName(MyIndex) & " - เลเวล " & GetPlayerLevel(MyIndex)
        
        ' Name Class 1.0
        Select Case (GetPlayerClass(MyIndex))
            Case 1: Class = "มนุษย์"
            Case 2: Class = "เอลฟ์"
            Case 3: Class = "การ์เดี้ยน"
            Case 4: Class = "เบอเซิร์ก"
            Case 5: Class = "พาลาดิน"
            Case 6: Class = "วิซาร์ด"
            Case 7: Class = "ซามูไร"
            Case 8: Class = "ฮันเตอร์"
            Case 9: Class = "สไนเปอร์"
            Case 10: Class = "แอสแซสซิน"
            Case 11: Class = "ดาร์คลอร์ด"
        End Select
        
        ' Name Class 2.0
        frmMain.lblNameCls.Caption = "อาชีพ : " & Class
        
        ' ------------------------ สูตรคำนวนแบบใหม่ !! -----------------------------
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                str = str + Item(GetPlayerEquipment(MyIndex, Weapon)).Add_Stat(1) + (Item(GetPlayerEquipment(MyIndex, Weapon)).Rarity * 5)
                ends = ends + Item(GetPlayerEquipment(MyIndex, Weapon)).Add_Stat(2) + (Item(GetPlayerEquipment(MyIndex, Weapon)).Rarity * 5)
                ints = ints + Item(GetPlayerEquipment(MyIndex, Weapon)).Add_Stat(3) + (Item(GetPlayerEquipment(MyIndex, Weapon)).Rarity * 5)
                agi = agi + Item(GetPlayerEquipment(MyIndex, Weapon)).Add_Stat(4) + (Item(GetPlayerEquipment(MyIndex, Weapon)).Rarity * 5)
                will = will + Item(GetPlayerEquipment(MyIndex, Weapon)).Add_Stat(5) + (Item(GetPlayerEquipment(MyIndex, Weapon)).Rarity * 5)
            End If
        
        If GetPlayerEquipment(MyIndex, Armor) > 0 Then
            str = str + Item(GetPlayerEquipment(MyIndex, Armor)).Add_Stat(1) + (Item(GetPlayerEquipment(MyIndex, Armor)).Rarity * 5)
            ends = ends + Item(GetPlayerEquipment(MyIndex, Armor)).Add_Stat(2) + (Item(GetPlayerEquipment(MyIndex, Armor)).Rarity * 5)
            ints = ints + Item(GetPlayerEquipment(MyIndex, Armor)).Add_Stat(3) + (Item(GetPlayerEquipment(MyIndex, Armor)).Rarity * 5)
            agi = agi + Item(GetPlayerEquipment(MyIndex, Armor)).Add_Stat(4) + (Item(GetPlayerEquipment(MyIndex, Armor)).Rarity * 5)
            will = will + Item(GetPlayerEquipment(MyIndex, Armor)).Add_Stat(5) + (Item(GetPlayerEquipment(MyIndex, Armor)).Rarity * 5)
        End If
        
        If GetPlayerEquipment(MyIndex, Shield) > 0 Then
            str = str + Item(GetPlayerEquipment(MyIndex, Shield)).Add_Stat(1) + (Item(GetPlayerEquipment(MyIndex, Shield)).Rarity * 5)
            ends = ends + Item(GetPlayerEquipment(MyIndex, Shield)).Add_Stat(2) + (Item(GetPlayerEquipment(MyIndex, Shield)).Rarity * 5)
            ints = ints + Item(GetPlayerEquipment(MyIndex, Shield)).Add_Stat(3) + (Item(GetPlayerEquipment(MyIndex, Shield)).Rarity * 5)
            agi = agi + Item(GetPlayerEquipment(MyIndex, Shield)).Add_Stat(4) + (Item(GetPlayerEquipment(MyIndex, Shield)).Rarity * 5)
            will = will + Item(GetPlayerEquipment(MyIndex, Shield)).Add_Stat(5) + (Item(GetPlayerEquipment(MyIndex, Shield)).Rarity * 5)
        End If
        
        If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
            str = str + Item(GetPlayerEquipment(MyIndex, Helmet)).Add_Stat(1) + (Item(GetPlayerEquipment(MyIndex, Helmet)).Rarity * 5)
            ends = ends + Item(GetPlayerEquipment(MyIndex, Helmet)).Add_Stat(2) + (Item(GetPlayerEquipment(MyIndex, Helmet)).Rarity * 5)
            ints = ints + Item(GetPlayerEquipment(MyIndex, Helmet)).Add_Stat(3) + (Item(GetPlayerEquipment(MyIndex, Helmet)).Rarity * 5)
            agi = agi + Item(GetPlayerEquipment(MyIndex, Helmet)).Add_Stat(4) + (Item(GetPlayerEquipment(MyIndex, Helmet)).Rarity * 5)
            will = will + Item(GetPlayerEquipment(MyIndex, Helmet)).Add_Stat(5) + (Item(GetPlayerEquipment(MyIndex, Helmet)).Rarity * 5)
        End If
        
        For X = 1 To Stats.Stat_Count - 1
            Select Case (X)
                Case 1: frmMain.lblCharStat(X).Caption = GetPlayerStat(MyIndex, X) - str & "(+" & str & ")"
                Case 2: frmMain.lblCharStat(X).Caption = GetPlayerStat(MyIndex, X) - ends & "(+" & ends & ")"
                Case 3: frmMain.lblCharStat(X).Caption = GetPlayerStat(MyIndex, X) - ints & "(+" & ints & ")"
                Case 4: frmMain.lblCharStat(X).Caption = GetPlayerStat(MyIndex, X) - agi & "(+" & agi & ")"
                Case 5: frmMain.lblCharStat(X).Caption = GetPlayerStat(MyIndex, X) - will & "(+" & will & ")"
            End Select
        Next
        
        ' ซ่อนหน้าต่างเมื่อ Status ตัน
        frmMain.lblPoints.Caption = GetPlayerPOINTS(MyIndex)
        If GetPlayerPOINTS(MyIndex) > 0 Then
            For X = 1 To Stats.Stat_Count - 1
                Select Case (X)
                    Case 1
                    ' สูตรคำนวนสเตตัสใหม่
                    If GetPlayerStat(Index, X) - str < 200 Then
                        frmMain.lblTrainStat(X).Visible = True
                        frmMain.lblCharStat(X).Caption = GetPlayerStat(Index, X) - str & "(+" & str & ")"
                    Else
                        frmMain.lblTrainStat(X).Visible = False
                        frmMain.lblCharStat(X).Caption = GetPlayerStat(Index, X) - str & "(+" & str & ")"
                    End If
                    
                    Case 2
                    ' สูตรคำนวนสเตตัสใหม่
                    If GetPlayerStat(Index, X) - ends < 200 Then
                        frmMain.lblTrainStat(X).Visible = True
                        frmMain.lblCharStat(X).Caption = GetPlayerStat(Index, X) - ends & "(+" & ends & ")"
                    Else
                        frmMain.lblTrainStat(X).Visible = False
                        frmMain.lblCharStat(X).Caption = GetPlayerStat(Index, X) - ends & "(+" & ends & ")"
                    End If

                    Case 3
                    ' สูตรคำนวนสเตตัสใหม่
                    If GetPlayerStat(Index, X) - ints < 200 Then
                        frmMain.lblTrainStat(X).Visible = True
                        frmMain.lblCharStat(X).Caption = GetPlayerStat(Index, X) - ints & "(+" & ints & ")"
                    Else
                        frmMain.lblTrainStat(X).Visible = False
                        frmMain.lblCharStat(X).Caption = GetPlayerStat(Index, X) - ints & "(+" & ints & ")"
                    End If

                    Case 4
                    ' สูตรคำนวนสเตตัสใหม่
                    If GetPlayerStat(Index, X) - agi < 200 Then
                        frmMain.lblTrainStat(X).Visible = True
                        frmMain.lblCharStat(X).Caption = GetPlayerStat(Index, X) - agi & "(+" & agi & ")"
                    Else
                        frmMain.lblTrainStat(X).Visible = False
                        frmMain.lblCharStat(X).Caption = GetPlayerStat(Index, X) - agi & "(+" & agi & ")"
                    End If

                    Case 5
                    ' สูตรคำนวนสเตตัสใหม่
                    If GetPlayerStat(Index, X) - will < 200 Then
                        frmMain.lblTrainStat(X).Visible = True
                        frmMain.lblCharStat(X).Caption = GetPlayerStat(Index, X) - will & "(+" & will & ")"
                    Else
                        frmMain.lblTrainStat(X).Visible = False
                        frmMain.lblCharStat(X).Caption = GetPlayerStat(Index, X) - will & "(+" & will & ")"
                    End If
                    
                End Select
            Next
        Else
            For X = 1 To Stats.Stat_Count - 1
                frmMain.lblTrainStat(X).Visible = False
            Next
        End If
        
        BltFace
    End If

    ' Make sure they aren't walking
    'Player(i).Moving = 0
    'Player(i).xOffset = 0
    'Player(i).yOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Buff
Private Sub HandlePlayerBuff(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long, X As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong

    For X = 1 To MAX_BUFF
        Player(i).BuffStatus(X) = Buffer.ReadByte
    Next
    For X = 1 To MAX_BUFF
        Player(i).BuffTime(X) = Buffer.ReadByte
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerBuff", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim n As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    n = Buffer.ReadLong
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, Dir)
    Player(i).xOffset = 0
    Player(i).yOffset = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)
        Case DIR_UP
            Player(i).yOffset = PIC_Y
        Case DIR_DOWN
            Player(i).yOffset = PIC_Y * -1
        Case DIR_LEFT
            Player(i).xOffset = PIC_X
        Case DIR_RIGHT
            Player(i).xOffset = PIC_X * -1
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim Movement As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MapNpcNum = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Movement = Buffer.ReadLong

    With MapNpc(MapNpcNum)
        .X = X
        .Y = Y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = Movement

        Select Case .Dir
            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcWarp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim mapnum As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MapNpcNum = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    mapnum = Buffer.ReadLong

    With MapNpc(MapNpcNum)
        .X = X
        .Y = Y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0

        Select Case .Dir
            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With
    
    Call BltNpc(MapNpcNum)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcWarp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerDir(i, Dir)

    With Player(i)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Dir = Buffer.ReadLong

    With MapNpc(i)
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, Dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).xOffset = 0
    Player(MyIndex).yOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim Buffer As clsBuffer
Dim thePlayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    thePlayer = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, Dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).xOffset = 0
    Player(thePlayer).yOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXYMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = GetTickCount
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim i As Long
Dim NeedMap As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    
    ' clear the blood
    For i = 1 To MAX_BYTE
        Blood(i).X = 0
        Blood(i).Y = 0
        Blood(i).Sprite = 0
        Blood(i).Timer = 0
    Next
    
    Map.CurrentEvents = 0
    ReDim Map.MapEvents(0)
    
    ' Get map num
    X = Buffer.ReadLong
    ' Get revision
    Y = Buffer.ReadLong

    If FileExist(MAP_PATH & "map" & X & MAP_EXT, False) Then
        Call LoadMap(X)
        ' Check to see if the revisions match
        NeedMap = 1

        If Map.Revision = Y Then
            ' We do so we dont need the map
            'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
            NeedMap = 0
        End If

    Else
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong NeedMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim i As Long
Dim Buffer As clsBuffer
Dim mapnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

    mapnum = Buffer.ReadLong
    Map.Name = Buffer.ReadString
    Map.Music = Buffer.ReadString
    Map.Weather = Buffer.ReadLong
    Map.Revision = Buffer.ReadLong
    Map.Moral = Buffer.ReadByte
    Map.Up = Buffer.ReadLong
    Map.Down = Buffer.ReadLong
    Map.Left = Buffer.ReadLong
    Map.Right = Buffer.ReadLong
    Map.BootMap = Buffer.ReadLong
    Map.BootX = Buffer.ReadByte
    Map.BootY = Buffer.ReadByte
    Map.maxX = Buffer.ReadByte
    Map.maxY = Buffer.ReadByte
    
    ReDim Map.Tile(0 To Map.maxX, 0 To Map.maxY)

    For X = 0 To Map.maxX
        For Y = 0 To Map.maxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Layer(i).X = Buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Y = Buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Tileset = Buffer.ReadLong
            Next
            Map.Tile(X, Y).Type = Buffer.ReadByte
            Map.Tile(X, Y).Data1 = Buffer.ReadLong
            Map.Tile(X, Y).Data2 = Buffer.ReadLong
            Map.Tile(X, Y).Data3 = Buffer.ReadLong
            Map.Tile(X, Y).DirBlock = Buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map.NPC(X) = Buffer.ReadLong
        n = n + 1
    Next

    ClearTempTile
    
    Set Buffer = Nothing
    
    ' Save the map
    Call SaveMap(mapnum)

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_ITEMS
        With MapItem(i)
            .playerName = Buffer.ReadString
            .Num = Buffer.ReadLong
            .Value = Buffer.ReadLong
            .X = Buffer.ReadLong
            .Y = Buffer.ReadLong
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(i)
            .Num = Buffer.ReadLong
            .X = Buffer.ReadLong
            .Y = Buffer.ReadLong
            .Dir = Buffer.ReadLong
            .Vital(HP) = Buffer.ReadLong
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapDone()
Dim i As Long, Tick As Long
Dim MusicFile As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Tick = GetTickCount
    
    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i
    Action_HighIndex = 1
    
    ' load tilesets we need
    LoadTilesets
            
    MusicFile = Trim$(Map.Music)
    If Not MusicFile = Music_Playing Then
        If Not MusicFile = "None." Then
            PlayMidi MusicFile
        Else
            StopMidi
        End If
    End If
    
    ' re-position the map name
    Call UpdateDrawMapName
    
    ' get the npc high index
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).Num > 0 Then
            Npc_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we're not overflowing
    If Npc_HighIndex > MAX_MAP_NPCS Then Npc_HighIndex = MAX_MAP_NPCS

    GettingMap = False
    CanMoveNow = True
    
    '     ' characters
            If NumCharacters > 0 Then
                For i = 1 To NumCharacters    'Check to unload surfaces
                    If CharacterTimer(i) > 0 Then 'Only update surfaces in use
                        If CharacterTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i)))
                            Set DDS_Character(i) = Nothing
                            CharacterTimer(i) = 0
                        End If
                    End If
                Next
            End If
      
      '      ' Paperdolls
            If NumPaperdolls > 0 Then
                For i = 1 To NumPaperdolls    'Check to unload surfaces
                    If PaperdollTimer(i) > 0 Then 'Only update surfaces in use
                        If PaperdollTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i)))
                            Set DDS_Paperdoll(i) = Nothing
                            PaperdollTimer(i) = 0
                        End If
                    End If
                Next
            End If

      '      ' animations
            If NumAnimations > 0 Then
                For i = 1 To NumAnimations    'Check to unload surfaces
                    If AnimationTimer(i) > 0 Then 'Only update surfaces in use
                        If AnimationTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i)))
                            Set DDS_Animation(i) = Nothing
                            AnimationTimer(i) = 0
                        End If
                    End If
                Next
            End If

     '       ' Items
            If NumItems > 0 Then
                For i = 1 To NumItems    'Check to unload surfaces
                    If ItemTimer(i) > 0 Then 'Only update surfaces in use
                        If ItemTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i)))
                            Set DDS_Item(i) = Nothing
                            ItemTimer(i) = 0
                        End If
                    End If
               Next
            End If

            ' Resources
            If NumResources > 0 Then
                For i = 1 To NumResources    'Check to unload surfaces
                    If ResourceTimer(i) > 0 Then 'Only update surfaces in use
                        If ResourceTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i)))
                            Set DDS_Resource(i) = Nothing
                            ResourceTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' spell icons
            If NumSpellIcons > 0 Then
                For i = 1 To NumSpellIcons    'Check to unload surfaces
                    If SpellIconTimer(i) > 0 Then 'Only update surfaces in use
                        If SpellIconTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i)))
                            Set DDS_SpellIcon(i) = Nothing
                            SpellIconTimer(i) = 0
                        End If
                    End If
                Next
            End If
     
            ' faces
            If NumFaces > 0 Then
               For i = 1 To NumFaces    'Check to unload surfaces
                    If FaceTimer(i) > 0 Then 'Only update surfaces in use
                        If FaceTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Face(i)), LenB(DDSD_Face(i)))
                            Set DDS_Face(i) = Nothing
                            FaceTimer(i) = 0
                        End If
                    End If
                Next
            End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapItem(n)
        .playerName = Buffer.ReadString
        .Num = Buffer.ReadLong
        .Value = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleItemEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleItemEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimationEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimationEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    BltInventory
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapNpc(n)
        .Num = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .Dir = Buffer.ReadLong
        .IsPet = Buffer.ReadByte
        .PetData.Name = Buffer.ReadString
        .PetData.Owner = Buffer.ReadLong
        ' Client use only
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Call ClearMapNpc(n)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDead", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(NPC(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    NpcSize = LenB(NPC(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(NPC(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim Buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ResourceNum = Buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    n = Buffer.ReadByte
    TempTile(X, Y).DoorOpen = n
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapKey", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEditMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleShopEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleShopEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim shopnum As Long
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopnum = Buffer.ReadLong
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpellEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpellEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim spellnum As Long
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellnum = Buffer.ReadLong
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i) = Buffer.ReadLong
    Next
    
    BltPlayerSpells
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call ClearPlayer(Buffer.ReadLong)
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Resource_Index = Buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = Buffer.ReadByte
            MapResource(i).X = Buffer.ReadLong
            MapResource(i).Y = Buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
    ' Call DrawPing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleDoorAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    With TempTile(X, Y)
        .DoorFrame = 1
        .DoorAnimate = 1 ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = GetTickCount
    End With
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDoorAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, Message As String, Color As Long, tmpType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    Message = Buffer.ReadString
    Color = Buffer.ReadLong
    tmpType = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong

    Set Buffer = Nothing
    
    CreateActionMsg Message, Color, tmpType, X, Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, Sprite As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong

    Set Buffer = Nothing
    
    ' randomise sprite
    Sprite = Rand(1, BloodCount)
    
    ' make sure tile doesn't already have blood
    For i = 1 To MAX_BYTE
        If Blood(i).X = X And Blood(i).Y = Y Then
            ' already have blood :(
            Exit Sub
        End If
    Next
    
    ' carry on with the set
    BloodIndex = BloodIndex + 1
    If BloodIndex >= MAX_BYTE Then BloodIndex = 1
    
    With Blood(BloodIndex)
        .X = X
        .Y = Y
        .Sprite = Sprite
        .Timer = GetTickCount
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBlood", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .LockType = Buffer.ReadByte
        .lockindex = Buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With
    
    ' play the sound if we've got one
    PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim MapNpcNum As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MapNpcNum = Buffer.ReadLong
    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    SpellCD(Slot) = GetTickCount
    
    BltPlayerSpells
    BltHotbar
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCooldown", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SpellBuffer = 0
    SpellBufferTimer = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Access As Long
Dim Name As String
Dim Message As String
Dim Colour As Long
Dim Header As String
Dim PK As Long
Dim saycolour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    Access = Buffer.ReadLong
    PK = Buffer.ReadLong
    Message = Buffer.ReadString
    Header = Buffer.ReadString
    saycolour = Buffer.ReadLong
    
    ' Check access level
        Select Case Access
            Case 0
                Colour = QBColor(White)
            Case 1
                Colour = QBColor(Yellow)
            Case 2
                Colour = QBColor(Yellow)
            Case 3
                Colour = QBColor(BrightRed)
            Case 4
                Colour = QBColor(BrightRed)
        End Select
    
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = Colour
    frmMain.txtChat.SelText = vbNewLine & Header & Name & " : "
    frmMain.txtChat.SelColor = saycolour
    frmMain.txtChat.SelText = Message
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    ' ReOrderChat Header & Name & " : " & Message, Colour
        
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim shopnum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopnum = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    OpenShop shopnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ShopAction = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResetShopAction", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    StunDuration = Buffer.ReadLong
    StunTime = 0
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_BANK
        Bank.Item(i).Num = Buffer.ReadLong
        Bank.Item(i).Value = Buffer.ReadLong
    Next
    
    InBank = True
    frmMain.picCover.Visible = True
    frmMain.picBank.Visible = True
    BltBank
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    InTrade = Buffer.ReadLong
    frmMain.picCover.Visible = True
    frmMain.picTrade.Visible = True
    BltTrade
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InTrade = 0
    frmMain.picCover.Visible = False
    frmMain.picTrade.Visible = False
    ' DeclineTrade
    frmMain.lblTradeStatus.Caption = vbNullString
    ' re-blt any items we were offering
    BltInventory
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCloseTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim dataType As Byte
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    dataType = Buffer.ReadByte
    
    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).Num = Buffer.ReadLong
            TradeYourOffer(i).Value = Buffer.ReadLong
        Next
        frmMain.lblYourWorth.Caption = Buffer.ReadLong & "g"
        ' remove any items we're offering
        BltInventory
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).Num = Buffer.ReadLong
            TradeTheirOffer(i).Value = Buffer.ReadLong
        Next
        frmMain.lblTheirWorth.Caption = Buffer.ReadLong & "g"
    End If
    
    BltTrade
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim tradeStatus As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeStatus = Buffer.ReadByte
    
    Set Buffer = Nothing
    
    Select Case tradeStatus
        Case 0 ' clear
            frmMain.lblTradeStatus.Caption = vbNullString
        Case 1 ' they've accepted
            frmMain.lblTradeStatus.Caption = "ผู้เล่นอีกฝ่ายได้กดพร้อมแลกเปลี่ยนแล้ว."
        Case 2 ' you've accepted
            frmMain.lblTradeStatus.Caption = "กำลังรอผู้เล่นอีกฝ่ายกดพร้อมการแลกเปลี่ยน."
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    myTarget = Buffer.ReadLong
    myTargetType = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = Buffer.ReadLong
        Hotbar(i).sType = Buffer.ReadByte
    Next
    BltHotbar
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Player_HighIndex = Buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    entityType = Buffer.ReadLong
    entityNum = Buffer.ReadLong
    
    PlayMapSound X, Y, entityType, entityNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    theName = Buffer.ReadString
    
    Dialogue "คำขอแลกเปลี่ยนไอเทม", theName & " ได้ส่งคำขอร้องแลกเปลี่ยนกับคุณ. ต้องการยืนยันหรือไม่?", DIALOGUE_TYPE_TRADE, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    theName = Buffer.ReadString
    
    Dialogue "คำขอเข้าร่วมปาร์ตี้", theName & " ได้เชิญคุณเข้าร่วมปาร์ตี้. ต้องการยืนยันหรือไม่?", DIALOGUE_TYPE_PARTY, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyInvite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, i As Long, inParty As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    inParty = Buffer.ReadByte
    
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(party), LenB(party))
        ' reset the labels
        For i = 1 To MAX_PARTY_MEMBERS
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        Next
        ' exit out early
        Exit Sub
    End If
    
    ' carry on otherwise
    party.Leader = Buffer.ReadLong
    For i = 1 To MAX_PARTY_MEMBERS
        party.Member(i) = Buffer.ReadLong
        If party.Member(i) > 0 Then
            frmMain.lblPartyMember(i).Caption = Trim$(GetPlayerName(party.Member(i)))
            frmMain.imgPartyHealth(i).Visible = True
            frmMain.imgPartySpirit(i).Visible = True
        Else
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        End If
    Next
    party.MemberCount = Buffer.ReadLong
    'party.Num = Buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim playerNum As Long, partyIndex As Long
Dim Buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' which player?
    playerNum = Buffer.ReadLong
    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = Buffer.ReadLong
        Player(playerNum).Vital(i) = Buffer.ReadLong
    Next
    
    ' find the party number
    For i = 1 To MAX_PARTY_MEMBERS
        If party.Member(i) = playerNum Then
            partyIndex = i
        End If
    Next
    
    ' exit out if wrong data
    If partyIndex <= 0 Or partyIndex > MAX_PARTY_MEMBERS Then Exit Sub
    
    ' hp bar
    frmMain.imgPartyHealth(partyIndex).Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
    ' spr bar
    frmMain.imgPartySpirit(partyIndex).Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
    
    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleQuestEditor()
    Dim i As Long
    
    With frmEditor_Quest
        Editor = EDITOR_TASKS
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_QUESTS
            .lstIndex.AddItem i & ": " & Trim$(Quest(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        QuestEditorInit
    End With

End Sub

Private Sub HandleUpdateQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    For i = 1 To MAX_QUESTS
        Player(MyIndex).PlayerQuest(i).Status = Buffer.ReadLong
        Player(MyIndex).PlayerQuest(i).ActualTask = Buffer.ReadLong
        Player(MyIndex).PlayerQuest(i).CurrentCount = Buffer.ReadLong
    Next
    
    RefreshQuestLog
    
    Set Buffer = Nothing
End Sub

Private Sub HandleQuestMessage(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, QuestNum As Long, QuestNumForStart As Long
    Dim Message As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    QuestNum = Buffer.ReadLong
    Message = Trim$(Buffer.ReadString)
    QuestNumForStart = Buffer.ReadLong
    
    frmMain.lblQuestName = Trim$(Quest(QuestNum).Name)
    frmMain.lblQuestSay = Message
    frmMain.picQuestDialogue.Visible = True
    
    If QuestNumForStart > 0 And QuestNumForStart <= MAX_QUESTS Then
        frmMain.lblQuestAccept.Visible = True
        frmMain.lblQuestAccept.Tag = QuestNumForStart
    End If
        
    Set Buffer = Nothing
End Sub

Private Sub HandleNPCCache(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim pIndex As Long
    Dim mapnum As Long
    Dim NPCNum As Long
    
    Dim i As Long
    
    Set Buffer = New clsBuffer
   
    Buffer.WriteBytes Data()
    
    mapnum = Buffer.ReadLong
    NPCNum = Buffer.ReadLong

    Map.NPC(NPCNum) = Buffer.ReadLong
    MapNpc(NPCNum).Num = Buffer.ReadLong

    
    Set Buffer = Nothing

End Sub


Private Sub HandleDoorsEditor()
    Dim i As Long

    With frmEditor_Doors
        Editor = EDITOR_DOORS
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_DOORS
            .lstIndex.AddItem i & ": " & Trim$(Doors(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        DoorEditorInit
    End With

End Sub

Private Sub HandleUpdateDoors(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim DoorNum As Long
Dim Buffer As clsBuffer
Dim DoorSize As Long
Dim DoorData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    DoorNum = Buffer.ReadLong
    
    DoorSize = LenB(Doors(DoorNum))
    ReDim DoorData(DoorSize - 1)
    DoorData = Buffer.ReadBytes(DoorSize)
    
    ClearDoor DoorNum
    
    CopyMemory ByVal VarPtr(Doors(DoorNum)), ByVal VarPtr(DoorData(0)), DoorSize
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateDoors", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnEventPage(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim id As Long, i As Long, Z As Long, X As Long, Y As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    id = Buffer.ReadLong
    
    If id > Map.CurrentEvents Then
        Map.CurrentEvents = id
        ReDim Preserve Map.MapEvents(Map.CurrentEvents)
    End If

    With Map.MapEvents(id)
        .Name = Buffer.ReadString
        .Dir = Buffer.ReadLong
        .ShowDir = .Dir
        .GraphicNum = Buffer.ReadLong
        .GraphicType = Buffer.ReadLong
        .GraphicX = Buffer.ReadLong
        .GraphicX2 = Buffer.ReadLong
        .GraphicY = Buffer.ReadLong
        .GraphicY2 = Buffer.ReadLong
        .MovementSpeed = Buffer.ReadLong
        .Moving = 0
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .xOffset = 0
        .yOffset = 0
        .Position = Buffer.ReadLong
        .Visible = Buffer.ReadLong
        .WalkAnim = Buffer.ReadLong
        .DirFix = Buffer.ReadLong
        .WalkThrough = Buffer.ReadLong
        .ShowName = Buffer.ReadLong
    End With
    
    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnEventPage", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim id As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long, ShowDir As Long
Dim Movement As Long, MovementSpeed As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    id = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    ShowDir = Buffer.ReadLong
    MovementSpeed = Buffer.ReadLong
    
    If id > Map.CurrentEvents Then Exit Sub

    With Map.MapEvents(id)
        .X = X
        .Y = Y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 1
        .ShowDir = ShowDir
        .MovementSpeed = MovementSpeed
        

        Select Case Dir
            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEventMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Dir = Buffer.ReadLong
    
    If i > Map.CurrentEvents Then Exit Sub

    With Map.MapEvents(i)
        .Dir = Dir
        .ShowDir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEventDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventChat(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim Buffer As clsBuffer
Dim choices As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    frmMain.picEventChat.Visible = True
    EventReplyID = Buffer.ReadLong
    EventReplyPage = Buffer.ReadLong
    frmMain.lblEventChat.Caption = Buffer.ReadString
    frmMain.lblEventChat.Caption = Replace(frmMain.lblEventChat.Caption, "/p", Trim$(Player(MyIndex).Name))
    frmMain.picEventChat.Visible = True
    frmMain.lblEventChat.Visible = True
    choices = Buffer.ReadLong
    
    InEvent = True
    
    For i = 1 To 4
        frmMain.lblChoices(i).Visible = False
    Next
    
    frmMain.lblEventChatContinue.Visible = False
    
    If choices = 0 Then
        frmMain.lblEventChatContinue.Visible = True
    Else
        For i = 1 To choices
            frmMain.lblChoices(i).Visible = True
            frmMain.lblChoices(i).Caption = Buffer.ReadString
        Next
    End If
    
    AnotherChat = Buffer.ReadLong
    
    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEventChat", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventStart(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InEvent = True

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEventStart", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventEnd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InEvent = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEventEnd", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayBGM(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    str = Buffer.ReadString
    
    StopMidi
    PlayMidi str
    
    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlaySound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    str = Buffer.ReadString

    PlaySound str
    
    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlaySound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFadeoutBGM(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    'Need to learn how to fadeout :P
    'do later... way later.. like, after release, maybe never
    StopMidi
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStopSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 0 To UBound(Sound()) - 1
        SoundStop (i)
    Next
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = Buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = Buffer.ReadString
    Next
    
    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapEventData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String, i As Long, X As Long, Y As Long, Z As Long, W As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    'Event Data!
    Map.EventCount = Buffer.ReadLong
        
    If Map.EventCount > 0 Then
        ReDim Map.Events(0 To Map.EventCount)
        For i = 1 To Map.EventCount
            With Map.Events(i)
                .Name = Buffer.ReadString
                .Global = Buffer.ReadLong
                .X = Buffer.ReadLong
                .Y = Buffer.ReadLong
                .pageCount = Buffer.ReadLong
            End With
            If Map.Events(i).pageCount > 0 Then
                ReDim Map.Events(i).Pages(0 To Map.Events(i).pageCount)
                For X = 1 To Map.Events(i).pageCount
                    With Map.Events(i).Pages(X)
                        .chkVariable = Buffer.ReadLong
                        .VariableIndex = Buffer.ReadLong
                        .VariableCondition = Buffer.ReadLong
                        .VariableCompare = Buffer.ReadLong
                            
                        .chkSwitch = Buffer.ReadLong
                        .SwitchIndex = Buffer.ReadLong
                        .SwitchCompare = Buffer.ReadLong
                            
                        .chkHasItem = Buffer.ReadLong
                        .HasItemIndex = Buffer.ReadLong
                            
                        .chkSelfSwitch = Buffer.ReadLong
                        .SelfSwitchIndex = Buffer.ReadLong
                        .SelfSwitchCompare = Buffer.ReadLong
                            
                        .GraphicType = Buffer.ReadLong
                        .Graphic = Buffer.ReadLong
                        .GraphicX = Buffer.ReadLong
                        .GraphicY = Buffer.ReadLong
                        .GraphicX2 = Buffer.ReadLong
                        .GraphicY2 = Buffer.ReadLong
                            
                        .MoveType = Buffer.ReadLong
                        .MoveSpeed = Buffer.ReadLong
                        .MoveFreq = Buffer.ReadLong
                            
                        .MoveRouteCount = Buffer.ReadLong
                        
                        .IgnoreMoveRoute = Buffer.ReadLong
                        .RepeatMoveRoute = Buffer.ReadLong
                            
                        If .MoveRouteCount > 0 Then
                            ReDim Map.Events(i).Pages(X).MoveRoute(0 To .MoveRouteCount)
                            For Y = 1 To .MoveRouteCount
                                .MoveRoute(Y).Index = Buffer.ReadLong
                                .MoveRoute(Y).Data1 = Buffer.ReadLong
                                .MoveRoute(Y).Data2 = Buffer.ReadLong
                                .MoveRoute(Y).Data3 = Buffer.ReadLong
                                .MoveRoute(Y).Data4 = Buffer.ReadLong
                                .MoveRoute(Y).Data5 = Buffer.ReadLong
                                .MoveRoute(Y).Data6 = Buffer.ReadLong
                            Next
                        End If
                            
                        .WalkAnim = Buffer.ReadLong
                        .DirFix = Buffer.ReadLong
                        .WalkThrough = Buffer.ReadLong
                        .ShowName = Buffer.ReadLong
                        .Trigger = Buffer.ReadLong
                        .CommandListCount = Buffer.ReadLong
                            
                        .Position = Buffer.ReadLong
                    End With
                        
                    If Map.Events(i).Pages(X).CommandListCount > 0 Then
                        ReDim Map.Events(i).Pages(X).CommandList(0 To Map.Events(i).Pages(X).CommandListCount)
                        For Y = 1 To Map.Events(i).Pages(X).CommandListCount
                            Map.Events(i).Pages(X).CommandList(Y).CommandCount = Buffer.ReadLong
                            Map.Events(i).Pages(X).CommandList(Y).ParentList = Buffer.ReadLong
                            If Map.Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                ReDim Map.Events(i).Pages(X).CommandList(Y).Commands(1 To Map.Events(i).Pages(X).CommandList(Y).CommandCount)
                                For Z = 1 To Map.Events(i).Pages(X).CommandList(Y).CommandCount
                                    With Map.Events(i).Pages(X).CommandList(Y).Commands(Z)
                                        .Index = Buffer.ReadLong
                                        .Text1 = Buffer.ReadString
                                        .Text2 = Buffer.ReadString
                                        .Text3 = Buffer.ReadString
                                        .Text4 = Buffer.ReadString
                                        .Text5 = Buffer.ReadString
                                        .Data1 = Buffer.ReadLong
                                        .Data2 = Buffer.ReadLong
                                        .Data3 = Buffer.ReadLong
                                        .Data4 = Buffer.ReadLong
                                        .Data5 = Buffer.ReadLong
                                        .Data6 = Buffer.ReadLong
                                        .ConditionalBranch.CommandList = Buffer.ReadLong
                                        .ConditionalBranch.Condition = Buffer.ReadLong
                                        .ConditionalBranch.Data1 = Buffer.ReadLong
                                        .ConditionalBranch.Data2 = Buffer.ReadLong
                                        .ConditionalBranch.Data3 = Buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = Buffer.ReadLong
                                        .MoveRouteCount = Buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For W = 1 To .MoveRouteCount
                                                .MoveRoute(W).Index = Buffer.ReadLong
                                                .MoveRoute(W).Data1 = Buffer.ReadLong
                                                .MoveRoute(W).Data2 = Buffer.ReadLong
                                                .MoveRoute(W).Data3 = Buffer.ReadLong
                                                .MoveRoute(W).Data4 = Buffer.ReadLong
                                                .MoveRoute(W).Data5 = Buffer.ReadLong
                                                .MoveRoute(W).Data6 = Buffer.ReadLong
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
    
    
    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapEventData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim spellnum As Long
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte

        If Options.Debug = 1 Then On Error GoTo errorhandler
        
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        spellnum = Buffer.ReadLong
        
        SpellSize = LenB(Spell(spellnum))
        ReDim SpellData(SpellSize - 1)
        SpellData = Buffer.ReadBytes(SpellSize)
        CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
        Set Buffer = Nothing

        Exit Sub
errorhandler:
        HandleError "HandleSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
