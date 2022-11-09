Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public DX7 As New DirectX7  ' Master Object, early binding

Public Sub Main()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' ��ͧ�ѹ�Դ���¨�
    ' Call MultiClient

    ' set loading screen
    loadGUI True
    frmLoad.Visible = True
    
    ' ����Դ��������
    Full_Game = False

    ' load options
    Call SetStatus("���ѧ��Ŵ�;���...")
    LoadOptions

    ' load main menu
    Call SetStatus("���ѧ��Ŵ����...")
    Load frmMenu
    
    ' load gui
    Call SetStatus("���ѧ��Ŵ�Թ������...")
    loadGUI
    
    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\data files\", "graphics"
    ChkDir App.Path & "\data files\graphics\", "animations"
    ChkDir App.Path & "\data files\graphics\", "characters"
    ChkDir App.Path & "\data files\graphics\", "items"
    ChkDir App.Path & "\data files\graphics\", "paperdolls"
    ChkDir App.Path & "\data files\graphics\", "resources"
    ChkDir App.Path & "\data files\graphics\", "spellicons"
    ChkDir App.Path & "\data files\graphics\", "tilesets"
    ChkDir App.Path & "\data files\graphics\", "faces"
    ChkDir App.Path & "\data files\graphics\", "gui"
    ChkDir App.Path & "\data files\graphics\gui\", "menu"
    ChkDir App.Path & "\data files\graphics\gui\", "main"
    ChkDir App.Path & "\data files\graphics\gui\menu\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "bars"
    ChkDir App.Path & "\data files\", "logs"
    ChkDir App.Path & "\data files\", "maps"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"
    
    ' Clear game values
        Call SetStatus("���ѧ���´ҵ����...")
        Call ClearGameData
    
    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34) ' "
    
    ' Update the form with the game's name before it's loaded
    If Full_Game = False Then
        frmMain.Caption = Options.Game_Name
    End If
    
    ' initialize DirectX
    InitDirectDraw
    
    If Not InitDirectDraw Then
        MsgBox "�Դ��Ҵ��û����żŢͧ DirectX7 - DirectDraw."
        MsgBox "����ٻ�Ҿ�Ҩ�٭������������."
        DestroyGame
   End If
    
    ' randomize rnd's seed
    Randomize
    Call SetStatus("���ѧ��駤�� TCP...")
    Call TcpInit
    Call InitMessages
    Call SetStatus("���ѧ��Ŵ�к� DirectX...")
    
    ' DX7 Master Object is already created, early binding
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    ' projectiles
    Call CheckProjectiles
    
    ' temp set music/sound vars
    Music_On = True
    Sound_On = True
    
    ' load music/sound engine
    InitSound
    InitMusic
    
    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMidi Trim$(Options.MenuMusic)
    
    ' Reset values
    Ping = -1
    
    'Load frmMainMenu
    Load frmMenu
    
    ' cache the buttons then reset & render them
    Call SetStatus("���ѧ��Ŵ����...")
    cacheButtons
    resetButtons_Menu
    
    ' we can now see it
    frmMenu.Visible = True
    
    ' hide all pics
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.Weapon
    
    ' hide the load form
    frmLoad.Visible = False
    
    ' turn on the chat
    Chat1(0) = 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub loadGUI(Optional ByVal loadingScreen As Boolean = False)
Dim i As Long

    ' if we can't find the interface
    On Error GoTo errorhandler
    
    ' loading screen
    If loadingScreen Then
        frmLoad.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\loading.jpg")
        Exit Sub
    End If

    ' menu
    frmMenu.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\menu.jpg")
    frmMenu.picMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\main.jpg")
    frmMenu.picLogin.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\login.jpg")
    frmMenu.picRegister.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\register.jpg")
    frmMenu.picCredits.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\credits.jpg")
    frmMenu.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\character.jpg")
    frmMenu.imgButton(1).Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\buttons\login_norm.jpg")
    frmMenu.imgButton(2).Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\buttons\register_norm.jpg")
    frmMenu.imgButton(3).Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\buttons\credits_norm.jpg")
    frmMenu.imgButton(4).Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\buttons\exit_norm.jpg")
    ' main
    frmMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\main.jpg")
    frmMain.picInventory.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\inventory.jpg")
    frmMain.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\character.jpg")
    frmMain.picSpells.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\skills.jpg")
    frmMain.picOptions.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\options.jpg")
    frmMain.picParty.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\party.jpg")
    frmMain.picParty2.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\party2.jpg")
    frmMain.picParty3.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\party3.jpg")
    frmMain.picItemDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_item.jpg")
    frmMain.picSpellDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_spell.jpg")
    frmMain.picTempInv.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTempBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTempSpell.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picShop.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\shop.jpg")
    frmMain.picBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bank.jpg")
    frmMain.picTrade.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\trade.jpg")
    frmMain.picHotbar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\hotbar.jpg")
    ' main - bars
    frmMain.imgHPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\health.jpg")
    frmMain.imgMPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\spirit.jpg")
    frmMain.imgEXPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\experience.jpg")
    frmMain.imgEXPBar2.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\experience.jpg")
    frmMain.imgEXPSKILL.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\experience.jpg")
    ' main - party bars
    For i = 1 To MAX_PARTY_MEMBERS
        frmMain.imgPartyHealth(i).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\party_health.jpg")
        frmMain.imgPartySpirit(i).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\party_spirit.jpg")
    Next
    
    ' store the bar widths for calculations
    HPBar_Width = frmMain.imgHPBar.Width
    SPRBar_Width = frmMain.imgMPBar.Width
    EXPBar_Width = frmMain.imgEXPBar.Width
    EXPBar_Width2 = frmMain.imgEXPBar2.Width
    SKILLBar_Width = frmMain.imgEXPSKILL.Width
    ' party
    Party_HPWidth = frmMain.imgPartyHealth(1).Width
    Party_SPRWidth = frmMain.imgPartySpirit(1).Width
    
    Exit Sub
    
' let them know we can't load the GUI
errorhandler:
    MsgBox "Cannot find one or more interface images." & vbNewLine & "If they exist then you have not extracted the project properly." & vbNewLine & "Please follow the installation instructions fully.", vbCritical
    DestroyGame
    Exit Sub
End Sub

Public Sub MenuState(ByVal state As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmLoad.Visible = True

    Select Case state
        Case MENU_STATE_ADDCHAR
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("���������������, ���ѧ���ҧ����Ф�...")

                If frmMenu.optMale.Value Then
                    Call SendAddChar(frmMenu.txtCName, SEX_MALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
                Else
                    Call SendAddChar(frmMenu.txtCName, SEX_FEMALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
                End If
            End If
            
        Case MENU_STATE_NEWACCOUNT
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("���������������, ���ѧ�觢������ʹ�...")
                Call SendNewAccount(frmMenu.txtRUser.text, frmMenu.txtRPass.text)
            End If

        Case MENU_STATE_LOGIN
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("���������������, ���ѧ��ͤ�Թ��������...")
                Call SendLogin(frmMenu.txtLUser.text, frmMenu.txtLPass.text)
                Exit Sub
            End If
    End Select

    If frmLoad.Visible Then
        If Not IsConnected Then
            frmMenu.Visible = True
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False
            frmLoad.Visible = False
            Call MsgBox("���ɤ��, �Թ�����絢ͧ��ҹ����ա���������� ���� �Կ�����Դ��Ѻ��ا ����ö��Ǩ�ͺʶҹ������� " & GAME_WEBSITE, vbOKOnly, Options.Game_Name)
            frmMenu.picMain.Visible = True
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub logoutGame()
Dim Buffer As clsBuffer, i As Long

    isLogging = True
    InGame = False
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuit
    Buffer.WriteString frmMain.lblPartyMember(2).Caption
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    Call DestroyTCP
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
    DragInvSlotNum = 0
    InvX = 0
    InvY = 0
    EqX = 0
    EqY = 0
    SpellX = 0
    SpellY = 0
    LastItemDesc = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0
    
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    
    ' hide main form stuffs
    frmMenu.picMain.Visible = True
    frmMain.txtChat.text = vbNullString
    frmMain.txtMyChat.text = vbNullString
    frmMain.picCurrency.Visible = False
    frmMain.picDialogue.Visible = False
    frmMain.picInventory.Visible = False
    frmMain.picTrade.Visible = False
    ' DeclineTrade
    frmMain.picCover.Visible = False
    frmMain.picSpells.Visible = False
    frmMain.picCharacter.Visible = False
    frmMain.picOptions.Visible = False
    frmMain.picParty.Visible = False
    frmMain.picParty2.Visible = False
    frmMain.picParty3.Visible = False
    frmMain.picAdmin.Visible = False
    If frmMain.picAdmin.Visible = False Then frmMain.Width = 11340
    frmMain.picBank.Visible = False

End Sub

Sub GameInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EnteringGame = True
    frmMenu.Visible = False
    EnteringGame = False
    
    ' bring all the main gui components to the front
    frmMain.picShop.ZOrder (0)
    frmMain.picBank.ZOrder (0)
    frmMain.picTrade.ZOrder (0)
    
    ' hide gui
    frmMain.picCover.Visible = False
    InBank = False
    InShop = False
    InTrade = False
    Fps_Max = False ' fps  on start
    hasParty = False
    
    If frmMain.picAdmin.Visible = True Then frmMain.Width = 14220
    If frmMain.picAdmin.Visible = False Then frmMain.Width = 11340
    
    ' Set font
    'Call SetFont(FONT_NAME, FONT_SIZE)
    frmMain.Font = "MS Sans Serif"
    frmMain.FontSize = 10
    frmMain.FontBold = True
    
    If Fps_Max = True Then
        frmMain.cmdFps.Caption = "�����礵��"
    Else
        frmMain.cmdFps.Caption = "�������٧"
    End If
    
    ' show the main form
    frmLoad.Visible = False
    frmMain.Show
    
    ' Set the focus
    Call SetFocusOnChat
    frmMain.picScreen.Visible = True
    
    ' Blt inv
    BltInventory
    
    ' blt hotbar
    BltHotbar
    
    ' get ping
    GetPing
    DrawPing
    
    ' set values for amdin panel
    frmMain.scrlAItem.Max = MAX_ITEMS
    frmMain.scrlAItem.Value = 1
    
    'stop the song playing
    StopMidi
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyGame()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' break out of GameLoop
    InGame = False
    Call DestroyTCP
    
    'destroy objects in reverse order
    Call DestroyDirectDraw

    ' destory DirectX7 master object
    If Not DX7 Is Nothing Then
        Set DX7 = Nothing
    End If

    Call UnloadAllForms
    End
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "destroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadAllForms()
Dim frm As Form

Chat1(0) = 1

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For Each frm In VB.Forms
        Unload frm
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStatus(ByVal Caption As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmLoad.lblStatus.Caption = Caption
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NewLine Then
        Txt.text = Txt.text + Msg + vbCrLf
    Else
        Txt.text = Txt.text + Msg
    End If

    Txt.SelStart = Len(Txt.text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetFocusOnChat()

    If chaton = False Then
        SetFocusOnGame
        Exit Sub
    End If

    On Error Resume Next 'prevent RTE5, no way to handle error
    frmMain.txtMyChat.Visible = True
    
    If frmMain.cbMAP.ListIndex < 0 Or frmMain.cbMAP.ListIndex > 3 Then
        frmMain.cbMAP.ListIndex = 0
    End If
    
    frmMain.txtMyChat.SetFocus
    
End Sub

Public Sub SetFocusOnGame()

    If chaton = True Then
        SetFocusOnChat
        Exit Sub
    End If

    On Error Resume Next 'prevent RTE5, no way to handle error
    frmMain.txtMyChat.Visible = False
    frmMain.picScreen.SetFocus
    
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Rand = Int((High - Low + 1) * Rnd) + Low
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub MovePicture(pb As PictureBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim GlobalX As Long
Dim GlobalY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GlobalX = pb.Left
    GlobalY = pb.Top

    If Button = 1 Then
        pb.Left = GlobalX + X - SOffsetX
        pb.Top = GlobalY + Y - SOffsetY
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MovePicture", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent high ascii chars
    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("�������ö���ѡ�� ASCII �� ��س��ͧ����.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If

    Next

    isStringLegal = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' ####################
' ## Buttons - Menu ##
' ####################
Public Sub cacheButtons()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' menu - login
    With MenuButton(1)
        .FileName = "login"
        .state = 0 ' normal
    End With
    
    ' menu - register
    With MenuButton(2)
        .FileName = "register"
        .state = 0 ' normal
    End With
    
    ' menu - credits
    With MenuButton(3)
        .FileName = "credits"
        .state = 0 ' normal
    End With
    
    ' menu - exit
    With MenuButton(4)
        .FileName = "exit"
        .state = 0 ' normal
    End With
    
    ' main - inv
    With MainButton(1)
        .FileName = "inv"
        .state = 0 ' normal
    End With
    
    ' main - skills
    With MainButton(2)
        .FileName = "skills"
        .state = 0 ' normal
    End With
    
    ' main - char
    With MainButton(3)
        .FileName = "char"
        .state = 0 ' normal
    End With
    
    ' main - opt
    With MainButton(4)
        .FileName = "opt"
        .state = 0 ' normal
    End With
    
    ' main - trade
    With MainButton(5)
        .FileName = "trade"
        .state = 0 ' normal
    End With
    
    ' main - party
    With MainButton(6)
        .FileName = "party"
        .state = 0 ' normal
    End With
    
        ' main - quest
    With MainButton(7)
        .FileName = "quest"
        .state = 0 ' normal
    End With
    
        ' main - pet
    With MainButton(8)
        .FileName = "pet"
        .state = 0 ' normal
    End With
    
    ' main - info
    With MainButton(9)
        .FileName = "inf"
        .state = 0 ' normal
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cacheButtons", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' menu specific buttons
Public Sub resetButtons_Menu(Optional ByVal exceptionNum As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_MENUBUTTONS
        ' only change if different and not exception
        If Not MenuButton(i).state = 0 And Not i = exceptionNum Then
            ' reset state and render
            MenuButton(i).state = 0 'normal
            renderButton_Menu i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Menu = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Menu(ByVal buttonNum As Long)
Dim bSuffix As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the suffix
    Select Case MenuButton(buttonNum).state
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMenu.imgButton(buttonNum).Picture = LoadPicture(App.Path & MENUBUTTON_PATH & MenuButton(buttonNum).FileName & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Menu(ByVal buttonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MenuButton(buttonNum).state = bState Then Exit Sub
        ' change and render
        MenuButton(buttonNum).state = bState
        renderButton_Menu buttonNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' main specific buttons
Public Sub resetButtons_Main(Optional ByVal exceptionNum As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_MAINBUTTONS
        ' only change if different and not exception
        If Not MainButton(i).state = 0 And Not i = exceptionNum Then
            ' reset state and render
            MainButton(i).state = 0 'normal
            renderButton_Main i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Main = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Main(ByVal buttonNum As Long)
Dim bSuffix As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the suffix
    Select Case MainButton(buttonNum).state
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMain.imgButton(buttonNum).Picture = LoadPicture(App.Path & MAINBUTTON_PATH & MainButton(buttonNum).FileName & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Main(ByVal buttonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MainButton(buttonNum).state = bState Then Exit Sub
        ' change and render
        MainButton(buttonNum).state = bState
        renderButton_Main buttonNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PopulateLists()
Dim strLoad As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Cache music list
    strLoad = Dir(App.Path & MUSIC_PATH & "*.mid")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop
    
    ' Cache sound list
    strLoad = Dir(App.Path & SOUND_PATH & "*.wav")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop
    
    'Cache MP3
        strLoad = Dir(App.Path & MUSIC_PATH & "*.mp3")
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearGameData()
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        Call ClearNpcs
        Call ClearResources
        Call ClearItems
        Call ClearShops
        Call ClearSpells
        Call ClearAnimations

        ' Error handler
        Exit Sub
errorhandler:
        HandleError "ClearGameData", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
