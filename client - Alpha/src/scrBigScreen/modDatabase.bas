Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim FileName As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\data files\logs\errors.txt"
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "' : " & erDesc & "."
        Print #1, ""
    Close #1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & FileName)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call WritePrivateProfileString$(Header, Var, Value, File)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\Data Files\config.ini"
    
    Call PutVar(FileName, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(FileName, "Options", "Username", Trim$(Options.Username))
    Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
    Call PutVar(FileName, "Options", "SavePass", str(Options.SavePass))
    Call PutVar(FileName, "Options", "IP", Options.IP)
    Call PutVar(FileName, "Options", "Port", str(Options.Port))
    Call PutVar(FileName, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(FileName, "Options", "Music", str(Options.Music))
    Call PutVar(FileName, "Options", "Sound", str(Options.Sound))
    Call PutVar(FileName, "Options", "Debug", str(Options.Debug))
    Call PutVar(FileName, "Options", "MiniMap", str(Options.Minimap))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\Data Files\config.ini"
    
    If Not FileExist(FileName, True) Then
        Options.Game_Name = "Monster war online"
        Options.Password = vbNullString
        Options.SavePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.Port = 401
        Options.MenuMusic = vbNullString
        Options.Music = 1
        Options.Sound = 1
        Options.Debug = 0
        Options.Minimap = 1
        Options.DefaultVolume = 80
        SaveOptions
    Else
        Options.Game_Name = GetVar(FileName, "Options", "Game_Name")
        Options.Username = GetVar(FileName, "Options", "Username")
        Options.Password = GetVar(FileName, "Options", "Password")
        Options.SavePass = Val(GetVar(FileName, "Options", "SavePass"))
        Options.IP = GetVar(FileName, "Options", "IP")
        Options.Port = Val(GetVar(FileName, "Options", "Port"))
        Options.MenuMusic = GetVar(FileName, "Options", "MenuMusic")
        Options.Music = GetVar(FileName, "Options", "Music")
        Options.Sound = GetVar(FileName, "Options", "Sound")
        Options.Debug = GetVar(FileName, "Options", "Debug")
        Options.Minimap = GetVar(FileName, "Options", "MiniMap")
        DefaultVolume = GetVar(FileName, "Options", "DefaultVolume")
    End If
    
    ' show in GUI
    If Options.Music = 0 Then
        frmMain.optMOff.Value = True
    Else
        frmMain.optMOn.Value = True
    End If
    
    If Options.Sound = 0 Then
        frmMain.optSOff.Value = True
    Else
        frmMain.optSOn.Value = True
    End If
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal mapnum As Long)
Dim FileName As String
Dim f As Long
Dim x As Long
Dim y As Long, i As Long, Z As Long, W As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & MAP_PATH & "map" & mapnum & MAP_EXT

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Map.Name
    Put #f, , Map.Music
    Put #f, , Map.Weather
    Put #f, , Map.Revision
    Put #f, , Map.Moral
    Put #f, , Map.Up
    Put #f, , Map.Down
    Put #f, , Map.Left
    Put #f, , Map.Right
    Put #f, , Map.BootMap
    Put #f, , Map.BootX
    Put #f, , Map.BootY
    Put #f, , Map.maxX
    Put #f, , Map.maxY

    For x = 0 To Map.maxX
        For y = 0 To Map.maxY
            Put #f, , Map.Tile(x, y)
        Next

        DoEvents
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , Map.NPC(x)
    Next

    Close #f
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal mapnum As Long)
Dim FileName As String
Dim f As Long
Dim x As Long
Dim y As Long, i As Long, Z As Long, W As Long, p As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & MAP_PATH & "map" & mapnum & MAP_EXT
    ClearMap
    f = FreeFile
    Open FileName For Binary As #f
    Get #f, , Map.Name
    Get #f, , Map.Music
    Get #f, , Map.Weather
    Get #f, , Map.Revision
    Get #f, , Map.Moral
    Get #f, , Map.Up
    Get #f, , Map.Down
    Get #f, , Map.Left
    Get #f, , Map.Right
    Get #f, , Map.BootMap
    Get #f, , Map.BootX
    Get #f, , Map.BootY
    Get #f, , Map.maxX
    Get #f, , Map.maxY
    ' have to set the tile()
    ReDim Map.Tile(0 To Map.maxX, 0 To Map.maxY)

    For x = 0 To Map.maxX
        For y = 0 To Map.maxY
            Get #f, , Map.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Get #f, , Map.NPC(x)
    Next

    Close #f
    ClearTempTile
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTilesets()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "\tilesets\" & i & GFX_EXT)
        NumTileSets = NumTileSets + 1
        i = i + 1
    Wend
    
    If NumTileSets = 0 Then Exit Sub
    
    ReDim DDS_Tileset(1 To NumTileSets)
    ReDim DDSD_Tileset(1 To NumTileSets)
    ReDim TilesetTimer(1 To NumTileSets)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckCharacters()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "characters\" & i & GFX_EXT)
        NumCharacters = NumCharacters + 1
        i = i + 1
    Wend
    
    If NumCharacters = 0 Then Exit Sub

    ReDim DDS_Character(1 To NumCharacters)
    ReDim DDSD_Character(1 To NumCharacters)
    ReDim CharacterTimer(1 To NumCharacters)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPaperdolls()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "paperdolls\" & i & GFX_EXT)
        NumPaperdolls = NumPaperdolls + 1
        i = i + 1
    Wend
    
    If NumPaperdolls = 0 Then Exit Sub

    ReDim DDS_Paperdoll(1 To NumPaperdolls)
    ReDim DDSD_Paperdoll(1 To NumPaperdolls)
    ReDim PaperdollTimer(1 To NumPaperdolls)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "animations\" & i & GFX_EXT)
        NumAnimations = NumAnimations + 1
        i = i + 1
    Wend
    
    If NumAnimations = 0 Then Exit Sub

    ReDim DDS_Animation(1 To NumAnimations)
    ReDim DDSD_Animation(1 To NumAnimations)
    ReDim AnimationTimer(1 To NumAnimations)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Items\" & i & GFX_EXT)
        NumItems = NumItems + 1
        i = i + 1
    Wend
    
    If NumItems = 0 Then Exit Sub

    ReDim DDS_Item(1 To NumItems)
    ReDim DDSD_Item(1 To NumItems)
    ReDim ItemTimer(1 To NumItems)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectiles
Public Sub CheckProjectiles()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Projectiles\" & i & GFX_EXT)
        NumProjectiles = NumProjectiles + 1
        i = i + 1
    Wend
    
    If NumProjectiles = 0 Then Exit Sub

    ReDim DDS_Projectile(1 To NumProjectiles)
    ReDim DDSD_Projectile(1 To NumProjectiles)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Resources\" & i & GFX_EXT)
        NumResources = NumResources + 1
        i = i + 1
    Wend
    
    If NumResources = 0 Then Exit Sub

    ReDim DDS_Resource(1 To NumResources)
    ReDim DDSD_Resource(1 To NumResources)
    ReDim ResourceTimer(1 To NumResources)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckSpellIcons()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "SpellIcons\" & i & GFX_EXT)
        NumSpellIcons = NumSpellIcons + 1
        i = i + 1
    Wend
    
    If NumSpellIcons = 0 Then Exit Sub

    ReDim DDS_SpellIcon(1 To NumSpellIcons)
    ReDim DDSD_SpellIcon(1 To NumSpellIcons)
    ReDim SpellIconTimer(1 To NumSpellIcons)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFaces()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Faces\" & i & GFX_EXT)
        NumFaces = NumFaces + 1
        i = i + 1
    Wend
    
    If NumFaces = 0 Then Exit Sub

    ReDim DDS_Face(1 To NumFaces)
    ReDim DDSD_Face(1 To NumFaces)
    ReDim FaceTimer(1 To NumFaces)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNPC(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpell(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShop(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResource(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projec tile
Sub ClearProjectile(ByVal Index As Long, ByVal PlayerProjectile As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Player(Index).ProjecTile(PlayerProjectile)
        .Direction = 0
        .Pic = 0
        .TravelTime = 0
        .x = 0
        .y = 0
        .Range = 0
        .Damage = 0
        .Speed = 0
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearProjectile", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    Map.maxX = MAX_MAPX
    Map.maxY = MAX_MAPY
    ReDim Map.Tile(0 To Map.maxX, 0 To Map.maxY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Name = Name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Sprite = Sprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).EXP
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).EXP = EXP
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerMaxVital = Player(Index).MaxVital(Vital)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_BYTE Then Value = MAX_BYTE
    Player(Index).Stat(Stat) = Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).POINTS = POINTS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal mapnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = mapnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).x
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).x = x
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).y
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).y = y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Dir = Dir
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invslot).Num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long, ByVal itemnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).Num = itemnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invslot).Value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).Value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    
    'If EquipmentSlot < 1 Then GetPlayerEquipment = 0
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = InvNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearDoor(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Doors(Index)), LenB(Doors(Index)))
    Doors(Index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearDoor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearDoors()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_DOORS
        Call ClearDoor(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
