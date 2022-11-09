Attribute VB_Name = "modGameEditors"
Option Explicit
Public cpEvent As EventRec
Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public EventList() As EventListRec

' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
Dim i As Long
Dim smusic() As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set the width
    frmEditor_Map.Width = 7425
    
    ' we're in the map editor
    InMapEditor = True
    
    ' show the form
    frmEditor_Map.Visible = True
    
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.Max = NumTileSets
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.Value = 1
    
    ' render the tiles
    Call EditorMap_BltTileset
    
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.Max = (frmEditor_Map.picBackSelect.Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.Max = (frmEditor_Map.picBackSelect.Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    MapEditorTileScroll
    
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"
    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).Name
    Next
    
    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorProperties()
Dim X As Long
Dim Y As Long
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    For i = 1 To UBound(musicCache)
        If musicCache(i) <> vbNullString Then
            frmEditor_MapProperties.lstMusic.AddItem musicCache(i)
        End If
    Next
    ' finished populating
    
    With frmEditor_MapProperties
        .txtName.text = Trim$(Map.Name)
        
        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For i = 0 To .lstMusic.ListCount
                If .lstMusic.List(i) = Trim$(Map.Music) Then
                    .lstMusic.ListIndex = i
                End If
            Next
        End If
        
        ' rest of it
        .txtUp.text = CStr(Map.Up)
        .txtDown.text = CStr(Map.Down)
        .txtLeft.text = CStr(Map.Left)
        .txtRight.text = CStr(Map.Right)
        .cmbMoral.ListIndex = Map.Moral
        .txtBootMap.text = CStr(Map.BootMap)
        .txtBootX.text = CStr(Map.BootX)
        .txtBootY.text = CStr(Map.BootY)
        .cmbWeather.ListIndex = CStr(Map.Weather)

        ' show the map npcs
        .lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS
            If Map.NPC(X) > 0 Then
            .lstNpcs.AddItem X & ": " & Trim$(NPC(Map.NPC(X)).Name)
            Else
                .lstNpcs.AddItem X & ": No NPC"
            End If
        Next
        .lstNpcs.ListIndex = 0
        
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"
        For X = 1 To MAX_NPCS
            .cmbNpc.AddItem X & ": " & Trim$(NPC(X).Name)
        Next
        
        ' set the combo box properly
        Dim tmpString() As String
        Dim NPCNum As Long
        tmpString = Split(.lstNpcs.List(.lstNpcs.ListIndex))
        NPCNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = Map.NPC(NPCNum)
    
        ' show the current map
        .lblMap.Caption = "แผนที่ปัจจุบัน : " & GetPlayerMap(MyIndex)
        .txtMaxX.text = Map.maxX
        .txtMaxY.text = Map.maxY
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorProperties", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False)
Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not multitile Then ' single
        With Map.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
        End With
    Else ' multitile
        Y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            X2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.maxX Then
                    If Y >= 0 And Y <= Map.maxY Then
                        With Map.Tile(X, Y)
                            .Layer(CurLayer).X = EditorTileX + X2
                            .Layer(CurLayer).Y = EditorTileY + Y2
                            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                        End With
                    End If
                End If
                X2 = X2 + 1
            Next
            Y2 = Y2 + 1
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSetTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
Dim i As Long
Dim CurLayer As Long
Dim tmpDir As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.Value Then
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer
            Else ' multi tile!
                MapEditorSetTile CurX, CurY, CurLayer, True
            End If
        ElseIf frmEditor_Map.optEvent.Value Then
            If frmEditor_Events.Visible = False Then
                AddEvent CurX, CurY
            End If
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmEditor_Map.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                End If
                ' item spawn
                If frmEditor_Map.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' key
                If frmEditor_Map.optKey.Value Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = 0
                End If
                ' key open
                If frmEditor_Map.optKeyOpen.Value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                End If
                ' resource
                If frmEditor_Map.optResource.Value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' door
                If frmEditor_Map.optDoor.Value Then
                    .Type = TILE_TYPE_DOOR
                    .Data1 = DoorEditorNum
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' npc spawn
                If frmEditor_Map.optNpcSpawn.Value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                End If
                ' shop
                If frmEditor_Map.optShop.Value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' bank
                If frmEditor_Map.optBank.Value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' heal
                If frmEditor_Map.optHeal.Value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                End If
                ' trap
                If frmEditor_Map.optTrap.Value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' slide
                If frmEditor_Map.optSlide.Value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                End If
            
                'Chest
                If frmEditor_Map.OptChest.Value Then
                    .Type = TILE_TYPE_CHEST
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                End If
                
                'sprite
                If frmEditor_Map.optSprite.Value Then
                .Type = TILE_TYPE_SPRITE
                .Data1 = TileSprite
                .Data2 = 0
                .Data3 = 0
                End If
                
                ' animation
                If frmEditor_Map.optAnimation.Value Then
                    .Type = TILE_TYPE_ANIMATION
                    .Data1 = AnimationNumber
                    .Data2 = 0
                    .Data3 = 0
                End If
                
                'Checkpoint
                If frmEditor_Map.optCheckpoint.Value Then
                If movedMouse Then Exit Sub
                    X = X - (CurX * 32)
                    Y = Y - (CurY * 32)
                    .Type = TILE_TYPE_CHECKPOINT
                    .Data1 = GetPlayerMap(MyIndex)
                    .Data2 = CurX
                    .Data3 = CurY
            End If
            
            ' Craft
                If frmEditor_Map.optCraft.Value Then
                    .Type = TILE_TYPE_CRAFT
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
            End If
            
            ' Dungeon
            
             If frmEditor_Map.optOnClick.Value Then
                    .Type = TILE_TYPE_ONCLICK
                    .Data1 = ScriptClick
                    .Data2 = 0
                    .Data3 = 0
                End If
            
            End With
        ElseIf frmEditor_Map.optBlock.Value Then
            If movedMouse Then Exit Sub
            ' find what tile it is
            X = X - ((X \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)
            ' see if it hits an arrow
            For i = 1 To 4
                If X >= DirArrowX(i) And X <= DirArrowX(i) + 8 Then
                    If Y >= DirArrowY(i) And Y <= DirArrowY(i) + 8 Then
                        ' flip the value.
                        setDirBlock Map.Tile(CurX, CurY).DirBlock, CByte(i), Not isDirBlocked(Map.Tile(CurX, CurY).DirBlock, CByte(i))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.Value Then
            With Map.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).X = 0
                .Layer(CurLayer).Y = 0
                .Layer(CurLayer).Tileset = 0
            End With
        ElseIf frmEditor_Map.optEvent.Value Then
            Call DeleteEvent(CurX, CurY)
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With

        End If
    End If

    CacheResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RandomTilePlacement(ByVal X As Long, ByVal Y As Long)
Dim i As Long
Dim CurLayer As Long

' If debug mode, handle error then exit out


' find which layer we're on
For i = 1 To MapLayer.Layer_Count - 1
If frmEditor_Map.optLayer(i).Value Then
CurLayer = i
Exit For
End If
Next

If Not isInBounds Then Exit Sub

    If frmEditor_Map.optLayers.Value Then
        If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
            MapEditorSetTile X, Y, CurLayer
        Else ' multi tile!
            MapEditorSetTile X, Y, CurLayer, True
        End If
    End If

CacheResources

' Error handler
Exit Sub
errorhandler:
HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
        
        frmEditor_Map.shpSelected.Top = EditorTileY * PIC_Y
        frmEditor_Map.shpSelected.Left = EditorTileX * PIC_X
        
        frmEditor_Map.shpSelected.Width = PIC_X
        frmEditor_Map.shpSelected.Height = PIC_Y
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > frmEditor_Map.picBackSelect.Width / PIC_X Then X = frmEditor_Map.picBackSelect.Width / PIC_X
        If Y < 0 Then Y = 0
        If Y > frmEditor_Map.picBackSelect.Height / PIC_Y Then Y = frmEditor_Map.picBackSelect.Height / PIC_Y
        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
        frmEditor_Map.shpSelected.Width = EditorTileWidth * PIC_X
        frmEditor_Map.shpSelected.Height = EditorTileHeight * PIC_Y
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorTileScroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' horizontal scrolling
    If frmEditor_Map.picBackSelect.Width < frmEditor_Map.picBack.Width Then
        frmEditor_Map.scrlPictureX.Enabled = False
    Else
        frmEditor_Map.scrlPictureX.Enabled = True
        frmEditor_Map.picBackSelect.Left = (frmEditor_Map.scrlPictureX.Value * PIC_X) * -1
    End If
    
    ' vertical scrolling
    If frmEditor_Map.picBackSelect.Height < frmEditor_Map.picBack.Height Then
        frmEditor_Map.scrlPictureY.Enabled = False
    Else
        frmEditor_Map.scrlPictureY.Enabled = True
        frmEditor_Map.picBackSelect.Top = (frmEditor_Map.scrlPictureY.Value * PIC_Y) * -1
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorTileScroll", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSend()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SendMap
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSend", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorCancel()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong 1
    SendData Buffer.ToArray()
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearLayer()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next
    
    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("คุณต้องการเคลีย เลเยอร์ นี้ใช่ไหม?", vbYesNo, Options.Game_Name) = vbYes Then
        For X = 0 To Map.maxX
            For Y = 0 To Map.maxY
                Map.Tile(X, Y).Layer(CurLayer).X = 0
                Map.Tile(X, Y).Layer(CurLayer).Y = 0
                Map.Tile(X, Y).Layer(CurLayer).Tileset = 0
            Next
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorFillLayer()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    ' Ground layer
    If MsgBox("คุณต้องการสร้างเลเยอร์นี้ทั้งแผนที่หรือไม่ ?", vbYesNo, Options.Game_Name) = vbYes Then
        For X = 0 To Map.maxX
            For Y = 0 To Map.maxY
                Map.Tile(X, Y).Layer(CurLayer).X = EditorTileX
                Map.Tile(X, Y).Layer(CurLayer).Y = EditorTileY
                Map.Tile(X, Y).Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            Next
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorFillLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearAttribs()
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MsgBox("คุณต้องการเคลียร์ Attributes ในแผนที่นี้หรือไม่ ?", vbYesNo, Options.Game_Name) = vbYes Then

        For X = 0 To Map.maxX
            For Y = 0 To Map.maxY
                Map.Tile(X, Y).Type = 0
            Next
        Next

    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearAttribs", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorLeaveMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorLeaveMap", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Item.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Item.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Item(EditorIndex)
        frmEditor_Item.txtName.text = Trim$(.Name)
        If .Pic > frmEditor_Item.scrlPic.Max Then .Pic = 0
        frmEditor_Item.scrlPic.Value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.Value = .Animation
        frmEditor_Item.txtDesc.text = Trim$(.Desc)
        
        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Item.cmbSound.ListCount
                If frmEditor_Item.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Item.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.Visible = True
            frmEditor_Item.scrlDamage.Value = .Data2
            frmEditor_Item.cmbTool.ListIndex = .Data3
            frmEditor_Item.fraOther.Visible = True
            
        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            frmEditor_Item.lblDamage.Caption = "Damage : " & frmEditor_Item.scrlDamage.Value
            frmEditor_Item.lblMATK.Caption = "เพิ่มโจมตีเวทย์ :"
            frmEditor_Item.txtKick.Visible = True
            frmEditor_Item.lblKick.Visible = True
            frmEditor_Item.txtNDEF.Visible = True
            frmEditor_Item.lblNDEF.Visible = True
            frmEditor_Item.txtCritRate.Visible = True
            frmEditor_Item.lblCritRate.Visible = True
            frmEditor_Item.txtDelayDown.Visible = True
            frmEditor_Item.lblDelayDown.Visible = True
            frmEditor_Item.lblDelayDownP.Visible = True
            frmEditor_Item.lblCritATK.Visible = True
            frmEditor_Item.txtCritATK.Visible = True
            frmEditor_Item.lblVampire.Visible = True
            frmEditor_Item.txtVampire.Visible = True
        Else
            frmEditor_Item.lblDamage.Caption = "Defense : " & frmEditor_Item.scrlDamage.Value
            frmEditor_Item.lblMATK.Caption = "เพิ่มป้องกันเวทย์เวทย์ :"
            frmEditor_Item.txtKick.Visible = False
            frmEditor_Item.lblKick.Visible = False
            frmEditor_Item.txtNDEF.Visible = False
            frmEditor_Item.lblNDEF.Visible = False
            frmEditor_Item.txtCritRate.Visible = False
            frmEditor_Item.lblCritRate.Visible = False
            frmEditor_Item.txtDelayDown.Visible = False
            frmEditor_Item.lblDelayDown.Visible = False
            frmEditor_Item.lblDelayDownP.Visible = False
            frmEditor_Item.lblCritATK.Visible = False
            frmEditor_Item.txtCritATK.Visible = False
            frmEditor_Item.lblVampire.Visible = False
            frmEditor_Item.txtVampire.Visible = False
            frmEditor_Item.lblDodge.Visible = True
            frmEditor_Item.txtDodge.Visible = True
        End If
        
        If Item(EditorIndex).LHand > 0 Then
            frmEditor_Item.lblDamage.Caption = "Damage : " & frmEditor_Item.scrlDamage.Value
            frmEditor_Item.lblMATK.Caption = "เพิ่มโจมตีเวทย์ :"
            frmEditor_Item.txtKick.Visible = True
            frmEditor_Item.lblKick.Visible = True
            frmEditor_Item.txtNDEF.Visible = True
            frmEditor_Item.lblNDEF.Visible = True
            frmEditor_Item.txtCritRate.Visible = True
            frmEditor_Item.lblCritRate.Visible = True
            frmEditor_Item.txtDelayDown.Visible = True
            frmEditor_Item.lblDelayDown.Visible = True
            frmEditor_Item.lblDelayDownP.Visible = True
            frmEditor_Item.lblCritATK.Visible = True
            frmEditor_Item.txtCritATK.Visible = True
            frmEditor_Item.lblVampire.Visible = True
            frmEditor_Item.txtVampire.Visible = True
        End If
        
            ' Check Speed
            If .Speed > 5000 Then .Speed = 5000
            frmEditor_Item.scrlSpeed.text = .Speed
            If .SpeedLow > 5000 Then .SpeedLow = 5000
            frmEditor_Item.scrlSpeedLow.text = .SpeedLow
            
            ' 2 hand and dagger
            If Item(EditorIndex).isTwoHanded Then
                frmEditor_Item.chkTwoh.Value = 1
            Else
                frmEditor_Item.chkTwoh.Value = 0
            End If

            If Item(EditorIndex).isDagger Then
                frmEditor_Item.ChkDagger.Value = 1
                frmEditor_Item.ScrlDagPdoll.Visible = True
                frmEditor_Item.lblDagPdoll.Visible = True
            Else
                frmEditor_Item.ChkDagger.Value = 0
                frmEditor_Item.ScrlDagPdoll.Visible = False
                frmEditor_Item.lblDagPdoll.Visible = False
            End If

            frmEditor_Item.ScrlDagPdoll.Value = Item(EditorIndex).Daggerpdoll
            
            ' loop for stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(i).Value = .Add_Stat(i)
            Next
            
            frmEditor_Item.scrlPaperdoll = .Paperdoll
            
            ' projectiles
        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
                frmEditor_Item.Frame4.Visible = True
                With Item(EditorIndex).ProjecTile
                    frmEditor_Item.scrlProjectileDamage.Value = .Damage
                    frmEditor_Item.scrlProjectilePic.Value = .Pic
                    frmEditor_Item.scrlProjectileRange.Value = .Range
                    frmEditor_Item.scrlProjectileSpeed.Value = .Speed
                End With
            End If
        Else
            frmEditor_Item.fraEquipment.Visible = False
            frmEditor_Item.Frame4.Visible = False
            frmEditor_Item.fraOther.Visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.Visible = True
            frmEditor_Item.scrlAddHp.Value = .AddHP
            frmEditor_Item.scrlAddMP.Value = .AddMP
            frmEditor_Item.scrlAddExp.Value = .AddEXP
            frmEditor_Item.scrlCastSpell.Value = .CastSpell
            frmEditor_Item.chkInstant.Value = .instaCast
        Else
            frmEditor_Item.fraVitals.Visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.Visible = True
            frmEditor_Item.scrlSpell.Value = .Data1
        Else
            frmEditor_Item.fraSpell.Visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_RECIPE) Then
            frmEditor_Item.fraRecipe.Visible = True
            frmEditor_Item.scrlItem1.Value = .Data1
            frmEditor_Item.scrlItem2.Value = .Data2
            frmEditor_Item.scrlResult.Value = .Data3
            frmEditor_Item.cmbCToolReq.ListIndex = .ToolReq
        Else
            frmEditor_Item.fraRecipe.Visible = False
        End If

frmEditor_Item.cmbCTool.ListIndex = .Tool

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.Value = .AccessReq
        frmEditor_Item.scrlLevelReq.Value = .LevelReq
        frmEditor_Item.scrlToolpower.Value = .Toolpower
        
        ' loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(i).Value = .Stat_Req(i)
        Next
        
        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(i).Name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.scrlPrice.Value = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.Value = .Rarity
        
        ' New Funtion by Allstar
        frmEditor_Item.txtKick.text = .Kick
        frmEditor_Item.txtMATK.text = .MATK
        frmEditor_Item.txtNDEF.text = .NDEF
        frmEditor_Item.txtCritRate.text = .CritRate
        
        If frmEditor_Item.txtDelayDown.text = 0 Then
            frmEditor_Item.txtDelayDown.text = 1
        End If
        
        frmEditor_Item.txtDelayDown.text = .DelayDown
        frmEditor_Item.txtCritATK.text = .CritATK
        frmEditor_Item.txtDelayUse.text = .DelayUse
        frmEditor_Item.chkReUse.Value = .ReUse
        
        frmEditor_Item.txtHP.text = .HP
        frmEditor_Item.txtMP.text = .MP
         
        frmEditor_Item.chkAdd1.Value = .Add1
        frmEditor_Item.chkSub1.Value = .Sub1
        frmEditor_Item.txtHPCase.text = .HPCase
        frmEditor_Item.chkAdd2.Value = .Add2
        frmEditor_Item.chkSub2.Value = .Sub2
        frmEditor_Item.txtMPCase.text = .MPCase
        frmEditor_Item.txtVampire.text = .Vampire
        frmEditor_Item.txtDodge.text = .Dodge
        frmEditor_Item.chkDropOnDeath.Value = .DropOnDeath
        
        ' 03/03/2558 Patch
        
        frmEditor_Item.txtRegenHp.text = .RegenHp
        frmEditor_Item.txtRegenMp.text = .RegenMp
        frmEditor_Item.chkPer1.Value = .Per1
        frmEditor_Item.chkPer2.Value = .Per2
        frmEditor_Item.chkClassR1.Value = .ClassR1
        frmEditor_Item.chkClassR2.Value = .ClassR2
        frmEditor_Item.chkClassR3.Value = .ClassR3
        frmEditor_Item.chkClassR4.Value = .ClassR4
        frmEditor_Item.chkClassR5.Value = .ClassR5
        frmEditor_Item.chkClassR6.Value = .ClassR6
        frmEditor_Item.chkClassR7.Value = .ClassR7
        frmEditor_Item.chkClassR8.Value = .ClassR8
        frmEditor_Item.chkClassR9.Value = .ClassR9
        frmEditor_Item.chkClassR10.Value = .ClassR10
        frmEditor_Item.chkClassR11.Value = .ClassR11
        
        frmEditor_Item.txtHP.text = .HP
        frmEditor_Item.txtMP.text = .MP
        
        frmEditor_Item.chkLHand.Value = .LHand
        frmEditor_Item.txtDMGLow.text = .DmgLow
        frmEditor_Item.txtMagicLow.text = .MagicLow
        frmEditor_Item.txtDMGHigh.text = .DmgHigh
        frmEditor_Item.txtMagicHigh.text = .MagicHigh
        frmEditor_Item.txtReflect.text = .Reflect
        frmEditor_Item.txtDmgReflect.text = .DmgReflect
        frmEditor_Item.txtAbsorbMagic.text = .AbsorbMagic
        
        ' buff item
        For i = 2 To 8
            frmEditor_Item.txtBuff(i).text = .Buff(i)
            frmEditor_Item.txtBuffTime(i).text = .BuffTime(i)
        Next
        
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With

    Call EditorItem_BltItem
    Call EditorItem_BltPaperdoll
    Item_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If
    Next
    
    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Item()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Item", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Animation.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Animation.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.text = Trim$(.Name)
        
        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Animation.cmbSound.ListCount
                If frmEditor_Animation.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If
        
        For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).Value = .Sprite(i)
            frmEditor_Animation.scrlFrameCount(i).Value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).Value = .LoopCount(i)
            
            If .looptime(i) > 0 Then
                frmEditor_Animation.scrlLoopTime(i).Value = .looptime(i)
            Else
                frmEditor_Animation.scrlLoopTime(i).Value = 45
            End If
            
        Next
        
        frmEditor_Animation.scrlAlpha.Value = .alpha
         
        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With

    Call EditorAnim_BltAnim
    Animation_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If
    Next
    
    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Animation()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Animation", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
Dim i As Long
Dim SoundSet As Boolean
Dim DropIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_NPC.Visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    DropIndex = frmEditor_NPC.scrlDrop.Value
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_NPC
        .txtName.text = Trim$(NPC(EditorIndex).Name)
        
        'If .txtAttackSay.text <> vbNullString Then
            .txtAttackSay.text = Trim$(NPC(EditorIndex).AttackSay)
        'Else
            '.txtAttackSay.text = " "
        'End If
        
        If NPC(EditorIndex).Sprite < 0 Or NPC(EditorIndex).Sprite > .scrlSprite.Max Then NPC(EditorIndex).Sprite = 0
        .scrlSprite.Value = NPC(EditorIndex).Sprite
        .txtSpawnSecs.text = CStr(NPC(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = NPC(EditorIndex).Behaviour
        .scrlRange.Value = NPC(EditorIndex).Range
        .txtChance.text = CStr(NPC(EditorIndex).DropChance(DropIndex))
        .scrlNum.Value = NPC(EditorIndex).DropItem(DropIndex)
        .scrlValue.Value = NPC(EditorIndex).DropItemValue(DropIndex)
        .scrlAnimation.Value = NPC(EditorIndex).Animation
        .txtHP.text = NPC(EditorIndex).HP
        .txtEXP.text = NPC(EditorIndex).EXP
        .txtEXP_max.text = NPC(EditorIndex).EXP_max
        .txtLevel.text = NPC(EditorIndex).Level
        .txtDamage.text = NPC(EditorIndex).Damage
        .scrlSpellNum.Max = MAX_NPC_SPELLS
        .scrlSpellNum.Value = 1
        .scrlSpell = NPC(EditorIndex).Spell(frmEditor_NPC.scrlSpellNum.Value) ' fixed
        .chkQuest.Value = NPC(EditorIndex).Quest
        .scrlQuest.Value = NPC(EditorIndex).QuestNum
        .txtBossNum.text = NPC(EditorIndex).BossNum
        ' Npc aspd
        .scrlAttackSpeed.text = NPC(EditorIndex).AttackSpeed
        .txtCrit.text = NPC(EditorIndex).CritRate
        .scrlCritChange.text = NPC(EditorIndex).CritChange
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(NPC(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).Value = NPC(EditorIndex).Stat(i)
        Next
        
    ' New system Allstar
    .txtDEF.text = NPC(EditorIndex).Def
    .txtDodge.text = NPC(EditorIndex).Dodge
    .txtBlock.text = NPC(EditorIndex).Block
    .txtRegenHp.text = NPC(EditorIndex).RegenHp
    .txtRegenMp.text = NPC(EditorIndex).RegenMp
    .txtMATK.text = NPC(EditorIndex).MATK
        
    .txtReflectDmg.text = NPC(EditorIndex).ReflectDmg
    .txtAbsorbMagic.text = NPC(EditorIndex).AbsorbMagic
    
    .scrlAlpha.Value = NPC(EditorIndex).alpha
        
    End With
    
    Call EditorNpc_BltSprite
    NPC_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If
    Next
    
    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_NPC()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_NPC", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
Dim i As Long
Dim SoundSet As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Resource.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Resource
        .scrlExhaustedPic.Max = NumResources
        .scrlNormalPic.Max = NumResources
        .scrlAnimation.Max = MAX_ANIMATIONS
        
        .txtName.text = Trim$(Resource(EditorIndex).Name)
        .txtMessage.text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlTool.Value = Resource(EditorIndex).ToolRequired
        .scrlHealth.Value = Resource(EditorIndex).health
        .scrlRespawn.Value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.Value = Resource(EditorIndex).Animation
        .scrlPwrReq.Value = Resource(EditorIndex).ToolpowerReq
        If Resource(EditorIndex).SuccessRate <= 0 Then
        frmEditor_Resource.scrlSuccess.Value = 1
    Else
        frmEditor_Resource.scrlSuccess.Value = Resource(EditorIndex).SuccessRate
    End If
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Resource(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    
    End With
        
    Call EditorResource_BltSprite
    
    Resource_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If
    Next
    
    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Resource()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Resource", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Shop.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    
    frmEditor_Shop.txtName.text = Trim$(Shop(EditorIndex).Name)
    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.Value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.Value = 100
    End If
    
    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(Item(i).Name)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(Item(i).Name)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    
    UpdateShopTrade
    
    Shop_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(i)
            ' if none, show as none
            If .Item = 0 And .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).Name) & " สำหรับ " & .CostValue & "x " & Trim$(Item(.CostItem).Name)
            End If
        End With
    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateShopTrade", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If
    Next
    
    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Shop()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Shop", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Spell.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."
    
    For i = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(i)
    Next
    
    ' finished populating

    With Spell(EditorIndex).ProjecTile
         frmEditor_Spell.scrlProjectileDamage.Value = .Damage
         frmEditor_Spell.scrlProjectilePic.Value = .Pic
         frmEditor_Spell.scrlProjectileRange.Value = .Range
         frmEditor_Spell.scrlProjectileSpeed.Value = .Speed
    End With
    
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.Max = MAX_ANIMATIONS
        .scrlAnim.Max = MAX_ANIMATIONS
        .scrlAOE.Max = MAX_BYTE
        .scrlRange.Max = MAX_BYTE
        .scrlMap.Max = MAX_MAPS
        .scrlIcon.Max = NumSpellIcons
        
        ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"
        For i = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(i).Name)
        Next
        
        ' fixed by allstar
        '.cmbClass.ListIndex = 0
        If Spell(EditorIndex).ClassReq > -1 And Spell(EditorIndex).ClassReq <= Max_Classes Then
            .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        End If
        
        ' set values
        .txtName.text = Trim$(Spell(EditorIndex).Name)
        .txtDesc.text = Trim$(Spell(EditorIndex).Desc)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.Value = Spell(EditorIndex).MPCost
        .scrlLevel.Value = Spell(EditorIndex).LevelReq
        .scrlAccess.Value = Spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        .scrlCast.Value = Spell(EditorIndex).CastTime
        .scrlCool.Value = Spell(EditorIndex).CDTime
        .scrlIcon.Value = Spell(EditorIndex).Icon
        .scrlMap.Value = Spell(EditorIndex).Map
        .scrlX.Value = Spell(EditorIndex).X
        .scrlY.Value = Spell(EditorIndex).Y
        .scrlDir.Value = Spell(EditorIndex).Dir
        .scrlVital.Value = Spell(EditorIndex).Vital
        .scrlDuration.Value = Spell(EditorIndex).Duration
        .scrlInterval.Value = Spell(EditorIndex).Interval
        .scrlRange.Value = Spell(EditorIndex).Range
        
        ' New Funtion by Allstar
        .chkPhysicalDmg.Value = Spell(EditorIndex).PhysicalDmg
        .chkMagicDmg.Value = Spell(EditorIndex).MagicDmg
        .txtATKPer.text = Spell(EditorIndex).ATKPer
        .txtMagicPer.text = Spell(EditorIndex).MagicPer
        
        If Spell(EditorIndex).Passive = 1 Then
            .fraPassive.Visible = True
        Else
            .fraPassive.Visible = False
        End If
        
        .chkPassive.Value = Spell(EditorIndex).Passive
        .chkPATK.Value = Spell(EditorIndex).PATK
        .chkPDEF.Value = Spell(EditorIndex).PDEF
        .txtPerSkill.text = Spell(EditorIndex).PerSkill
        .chkCanMove.Value = Spell(EditorIndex).CanMove
        
        If Spell(EditorIndex).IsAoE Then
            .chkAOE.Value = 1
        Else
            .chkAOE.Value = 0
        End If
        .scrlAOE.Value = Spell(EditorIndex).AoE
        .scrlAnimCast.Value = Spell(EditorIndex).CastAnim
        .scrlAnim.Value = Spell(EditorIndex).SpellAnim
        .scrlStun.Value = Spell(EditorIndex).StunDuration
        .scrlHP.Value = Spell(EditorIndex).HPCost
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Spell(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        .chkCanCancle.Value = Spell(EditorIndex).CanCancle
        .txtS1.text = Spell(EditorIndex).S1
        .txtS2.text = Spell(EditorIndex).S2
        .txtS3.text = Spell(EditorIndex).S3
        .txtS4.text = Spell(EditorIndex).S4
        
    End With
    
    EditorSpell_BltIcon
    
    Spell_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If
    Next
    
    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Spell()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Spell", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAttributeDialogue()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Map.fraNpcSpawn.Visible = False
    frmEditor_Map.fraResource.Visible = False
    frmEditor_Map.fraMapItem.Visible = False
    frmEditor_Map.fraMapKey.Visible = False
    frmEditor_Map.fraKeyOpen.Visible = False
    frmEditor_Map.fraMapWarp.Visible = False
    frmEditor_Map.fraShop.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAttributeDialogue", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'/////////
'//DOORS//
'/////////

Public Sub DoorEditorInit()
    If frmEditor_Doors.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Doors.lstIndex.ListIndex + 1
    
    With frmEditor_Doors
    
        .txtName.text = Doors(EditorIndex).Name
        If Doors(EditorIndex).DoorType = 0 Then
           .optDoor(0).Value = True
        Else
            .optDoor(1).Value = True
        End If
        .scrlKey.Value = Doors(EditorIndex).key
        .scrlSwitch.Value = Doors(EditorIndex).Switch
        .scrlMap.Value = Doors(EditorIndex).WarpMap
        .scrlX.Value = Doors(EditorIndex).WarpX
        .scrlY.Value = Doors(EditorIndex).WarpY
        If Doors(EditorIndex).UnlockType = 0 Then
            .OptUnlock(0).Value = True
        ElseIf Doors(EditorIndex).UnlockType = 1 Then
            .OptUnlock(1).Value = True
        Else
            .OptUnlock(2).Value = True
        End If
        
    End With
    Door_Changed(EditorIndex) = True
End Sub

Public Sub DoorEditorOk()
    Dim i As Long

    For i = 1 To MAX_DOORS
        If Door_Changed(i) Then
            Call SendSavedoor(i)
        End If
    Next
    
    Unload frmEditor_Doors
    Editor = 0
    ClearChanged_Doors
End Sub

Public Sub DoorEditorCancel()
    Editor = 0
    Unload frmEditor_Doors
    ClearChanged_Doors
    ClearDoors
    SendRequestDoors
End Sub

Public Sub ClearChanged_Doors()
    ZeroMemory Door_Changed(1), MAX_DOORS * 2 ' 2 = boolean length
End Sub

'Event Editor Stuffz Also includes event functions from the map editor (copy/paste/delete)

Sub CopyEvent_Map(X As Long, Y As Long)
Dim Count As Long, i As Long
    Count = Map.EventCount
    If Count = 0 Then Exit Sub
    
    For i = 1 To Count
        If Map.Events(i).X = X And Map.Events(i).Y = Y Then
            ' copy it
            'CopyMemory ByVal VarPtr(cpEvent), ByVal VarPtr(Map.Events(i)), LenB(Map.Events(i))
            cpEvent = Map.Events(i)
            ' exit
            Exit Sub
        End If
    Next
End Sub

Sub PasteEvent_Map(X As Long, Y As Long)
Dim Count As Long, i As Long, eventNum As Long
    Count = Map.EventCount
    
    If Count > 0 Then
        For i = 1 To Count
            If Map.Events(i).X = X And Map.Events(i).Y = Y Then
                ' already an event - paste over it
                eventNum = i
            End If
        Next
    End If
    
    ' couldn't find one - create one
    If eventNum = 0 Then
        ' increment count
        AddEvent X, Y, True
        eventNum = Count + 1
    End If
    
    ' copy it
    'CopyMemory ByVal VarPtr(Map.Events(eventNum)), ByVal VarPtr(cpEvent), LenB(cpEvent)
    Map.Events(eventNum) = cpEvent
    
    ' set position
    Map.Events(eventNum).X = X
    Map.Events(eventNum).Y = Y
End Sub

Sub DeleteEvent(X As Long, Y As Long)
Dim Count As Long, i As Long, lowIndex As Long
    If Not InMapEditor Then Exit Sub
    
    Count = Map.EventCount
    For i = 1 To Count
        If Map.Events(i).X = X And Map.Events(i).Y = Y Then
            ' delete it
            ClearEvent i
            lowIndex = i
            Exit For
        End If
    Next
    
    ' not found anything
    If lowIndex = 0 Then Exit Sub
    
    ' move everything down an index
    For i = lowIndex To Count - 1
        CopyEvent i + 1, i
    Next
    ' delete the last index
    ClearEvent Count
    ' set the new count
    Map.EventCount = Count - 1
End Sub

Sub AddEvent(X As Long, Y As Long, Optional ByVal cancelLoad As Boolean = False)
Dim Count As Long, pageCount As Long, i As Long
    Count = Map.EventCount + 1
    ' make sure there's not already an event
    If Count - 1 > 0 Then
        For i = 1 To Count - 1
            If Map.Events(i).X = X And Map.Events(i).Y = Y Then
                ' already an event - edit it
                If Not cancelLoad Then EventEditorInit i
                Exit Sub
            End If
        Next
    End If
    ' increment count
    Map.EventCount = Count
    ReDim Preserve Map.Events(0 To Count)
    ' set the new event
    Map.Events(Count).X = X
    Map.Events(Count).Y = Y
    ' give it a new page
    pageCount = Map.Events(Count).pageCount + 1
    Map.Events(Count).pageCount = pageCount
    ReDim Preserve Map.Events(Count).Pages(pageCount)
    ' load the editor
    If Not cancelLoad Then EventEditorInit Count
End Sub

Sub ClearEvent(eventNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map.Events(eventNum)), LenB(Map.Events(eventNum)))
End Sub

Sub CopyEvent(original As Long, newone As Long)
    CopyMemory ByVal VarPtr(Map.Events(newone)), ByVal VarPtr(Map.Events(original)), LenB(Map.Events(original))
End Sub

Sub EventEditorInit(eventNum As Long)
Dim i As Long
    EditorEvent = eventNum
    ' copy the event data to the temp event
    'CopyMemory ByVal VarPtr(tmpEvent), ByVal VarPtr(Map.Events(eventNum)), LenB(Map.Events(eventNum))
    tmpEvent = Map.Events(eventNum)
    ' populate form
    With frmEditor_Events
        ' set the tabs
        .tabPages.Tabs.Clear
        For i = 1 To tmpEvent.pageCount
            .tabPages.Tabs.Add , , str(i)
        Next
        ' items
        .cmbHasItem.Clear
        .cmbHasItem.AddItem "None"
        For i = 1 To MAX_ITEMS
            .cmbHasItem.AddItem i & ": " & Trim$(Item(i).Name)
        Next
        
        ' variables
        .cmbPlayerVar.Clear
        .cmbPlayerVar.AddItem "None"
        For i = 1 To MAX_VARIABLES
            .cmbPlayerVar.AddItem i & ". " & Variables(i)
        Next
        
        ' variables
        .cmbPlayerSwitch.Clear
        .cmbPlayerSwitch.AddItem "None"
        For i = 1 To MAX_SWITCHES
            .cmbPlayerSwitch.AddItem i & ". " & Switches(i)
        Next
        
        
        ' name
        .txtName.text = tmpEvent.Name
        ' enable delete button
        If tmpEvent.pageCount > 1 Then
            .cmdDeletePage.Enabled = True
        Else
            .cmdDeletePage.Enabled = False
        End If
        .cmdPastePage.Enabled = False
        ' Load page 1 to start off with
        curPageNum = 1
        EventEditorLoadPage curPageNum
    End With
    ' show the editor
    frmEditor_Events.Show
End Sub

Sub EventEditorLoadPage(pageNum As Long)
    ' populate form
    With tmpEvent.Pages(pageNum)
        GraphicSelX = .GraphicX
        GraphicSelY = .GraphicY
        GraphicSelX2 = .GraphicX2
        GraphicSelY2 = .GraphicY2
        frmEditor_Events.cmbGraphic.ListIndex = .GraphicType
        
        frmEditor_Events.cmbHasItem.ListIndex = .HasItemIndex
        frmEditor_Events.cmbMoveFreq.ListIndex = .MoveFreq
        frmEditor_Events.cmbMoveSpeed.ListIndex = .MoveSpeed
        frmEditor_Events.cmbMoveType.ListIndex = .MoveType
        
        frmEditor_Events.cmbPlayerVar.ListIndex = .VariableIndex
        frmEditor_Events.cmbPlayerSwitch.ListIndex = .SwitchIndex
        frmEditor_Events.cmbSelfSwitch.ListIndex = .SelfSwitchIndex
        frmEditor_Events.cmbSelfSwitchCompare.ListIndex = .SelfSwitchCompare
        frmEditor_Events.cmbPlayerSwitchCompare.ListIndex = .SwitchCompare
        frmEditor_Events.cmbPlayerVarCompare.ListIndex = .VariableCompare
        
        
        frmEditor_Events.chkGlobal.Value = tmpEvent.Global
        
        frmEditor_Events.cmbTrigger.ListIndex = .Trigger
        frmEditor_Events.chkDirFix.Value = .DirFix
        frmEditor_Events.chkHasItem.Value = .chkHasItem
        frmEditor_Events.chkPlayerVar.Value = .chkVariable
        frmEditor_Events.chkPlayerSwitch.Value = .chkSwitch
        frmEditor_Events.chkSelfSwitch.Value = .chkSelfSwitch
        frmEditor_Events.chkWalkAnim.Value = .WalkAnim
        frmEditor_Events.chkWalkThrough.Value = .WalkThrough
        frmEditor_Events.chkShowName.Value = .ShowName
        frmEditor_Events.txtPlayerVariable = .VariableCondition
        frmEditor_Events.scrlGraphic.Value = .Graphic
        
        If .chkHasItem = 0 Then
            frmEditor_Events.cmbHasItem.Enabled = False
        Else
            frmEditor_Events.cmbHasItem.Enabled = True
        End If
        
        
        If .chkSelfSwitch = 0 Then
            frmEditor_Events.cmbSelfSwitch.Enabled = False
            frmEditor_Events.cmbSelfSwitchCompare.Enabled = False
        Else
            frmEditor_Events.cmbSelfSwitch.Enabled = True
            frmEditor_Events.cmbSelfSwitchCompare.Enabled = True
        End If
        
        If .chkSwitch = 0 Then
            frmEditor_Events.cmbPlayerSwitch.Enabled = False
            frmEditor_Events.cmbPlayerSwitchCompare.Enabled = False
        Else
            frmEditor_Events.cmbPlayerSwitch.Enabled = True
            frmEditor_Events.cmbPlayerSwitchCompare.Enabled = True
        End If
        
        
        If .chkVariable = 0 Then
            frmEditor_Events.cmbPlayerVar.Enabled = False
            frmEditor_Events.txtPlayerVariable.Enabled = False
            frmEditor_Events.cmbPlayerVarCompare.Enabled = False
        Else
            frmEditor_Events.cmbPlayerVar.Enabled = True
            frmEditor_Events.txtPlayerVariable.Enabled = True
            frmEditor_Events.cmbPlayerVarCompare.Enabled = True
        End If
        
        If frmEditor_Events.cmbMoveType.ListIndex = 2 Then
            frmEditor_Events.cmdMoveRoute.Enabled = True
        Else
            frmEditor_Events.cmdMoveRoute.Enabled = False
        End If
        
        frmEditor_Events.cmbPositioning.ListIndex = .Position
        
        ' show the commands
        EventListCommands
    End With
End Sub

Sub EventEditorOK()
    ' copy the event data from the temp event
    'CopyMemory ByVal VarPtr(Map.Events(EditorEvent)), ByVal VarPtr(tmpEvent), LenB(tmpEvent)
    Map.Events(EditorEvent) = tmpEvent
    ' unload the form
    Unload frmEditor_Events
End Sub

Public Sub EventListCommands()
Dim i As Long, curlist As Long, oldI As Long, X As Long, indent As String, listleftoff() As Long, conditionalstage() As Long
    frmEditor_Events.lstCommands.Clear
    If tmpEvent.Pages(curPageNum).CommandListCount > 0 Then
    ReDim listleftoff(1 To tmpEvent.Pages(curPageNum).CommandListCount)
    ReDim conditionalstage(1 To tmpEvent.Pages(curPageNum).CommandListCount)
        'Start Up at 1
        curlist = 1
        X = -1
newlist:
        For i = 1 To tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
            If listleftoff(curlist) > 0 Then
                If (tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(listleftoff(curlist)).Index = EventType.evCondition Or tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(listleftoff(curlist)).Index = EventType.evShowChoices) And conditionalstage(curlist) <> 0 Then
                    i = listleftoff(curlist)
                ElseIf listleftoff(curlist) >= i Then
                    i = listleftoff(curlist) + 1
                End If
            End If
            If i <= tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
                If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Index = EventType.evCondition Then
                    X = X + 1
                    Select Case conditionalstage(curlist)
                        Case 0
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = i
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Condition
                                Case 0
                                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2
                                        Case 0
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 1
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] >= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 2
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] <= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 3
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] > " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 4
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] < " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 5
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] != " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                    End Select
                                Case 1
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] == " & "True"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] == " & "False"
                                    End If
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Has Item [" & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1).Name) & "]"
                                Case 3
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Class Is [" & Trim$(Class(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1).Name) & "]"
                                Case 4
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Knows Skill [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1).Name) & "]"
                                Case 5
                                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2
                                        Case 0
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 1
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is >= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 2
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is <= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 3
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is > " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 4
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is < " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 5
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is NOT " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                    End Select
                                Case 6
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2 = 0 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [A] == " & "True"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [B] == " & "True"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [C] == " & "True"
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [D] == " & "True"
                                        End Select
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2 = 1 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [A] == " & "False"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [B] == " & "False"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [C] == " & "False"
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [D] == " & "False"
                                        End Select
                                    End If
                            End Select
                            
                            indent = indent & "       "
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 1
                            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.CommandList
                            GoTo newlist
                        Case 1
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "Else"
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 2
                            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.ElseCommandList
                            GoTo newlist
                        Case 2
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "End Branch"
                            indent = Mid(indent, 1, Len(indent) - 7)
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 0
                    End Select
                ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Index = EventType.evShowChoices Then
                    X = X + 1
                    Select Case conditionalstage(curlist)
                        Case 0
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = i
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Choices - Prompt: " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "..."
                            
                            indent = indent & "       "
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 1
                            GoTo newlist
                        Case 1
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text2) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text2) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 2
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 2
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 2
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text3) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text3) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 3
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 3
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 3
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text4) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text4) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 4
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 4
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 4
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text5) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text5) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 5
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 5
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 5
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "Branch End"
                            indent = Mid(indent, 1, Len(indent) - 7)
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 0
                    End Select
                Else
                    X = X + 1
                    ReDim Preserve EventList(X)
                    EventList(X).CommandList = curlist
                    EventList(X).CommandNum = i
                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Index
                        Case EventType.evAddText
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2
                                Case 0
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Add Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " - Chat Type : Player"
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Add Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " - Chat Type : Map"
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Add Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " - Chat Type : Global"
                            End Select
                        Case EventType.evShowText
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "..."
                        Case EventType.evPlayerVar
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2
                        Case EventType.evPlayerSwitch
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] == True"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] == False"
                            End If
                        Case EventType.evSelfSwitch
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                                Case 0
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [A] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [A] to OFF"
                                    End If
                                Case 1
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [B] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [B] to OFF"
                                    End If
                                Case 2
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [C] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [C] to OFF"
                                    End If
                                Case 3
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [D] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [D] to OFF"
                                    End If
                            End Select
                        Case EventType.evExitProcess
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Exit Event Processing"
                        
                        Case EventType.evChangeItems
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Item Amount of [" & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "] to " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Give Player " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " " & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "(s)"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 2 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Take " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " " & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "(s) from Player."
                            End If
                            
                        Case EventType.evRestoreHP
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Restore Player HP"
                        Case EventType.evRestoreMP
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Restore Player MP"
                        Case EventType.evLevelUp
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Level Up Player"
                        Case EventType.evChangeLevel
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Level to " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                        Case EventType.evChangeSkills
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Teach Player Skill [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Remove Player Skill [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]"
                            End If
                        Case EventType.evChangeClass
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Class to " & Trim$(Class(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name)
                        Case EventType.evChangeSprite
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Sprite to " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                        Case EventType.evChangeSex
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Sex to Male."
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Sex to Female."
                            End If
                        Case EventType.evChangePK
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player PK to No."
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player PK to Yes."
                            End If
                        Case EventType.evWarpPlayer
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") while retaining direction."
                            Else
                                Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4 - 1
                                    Case DIR_UP
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing upward."
                                    Case DIR_DOWN
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing downward."
                                    Case DIR_LEFT
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing left."
                                    Case DIR_RIGHT
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing right."
                                End Select
                            End If
                        Case EventType.evSetMoveRoute
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 <= Map.EventCount Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Move Route for Event #" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]"
                            Else
                               frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Move Route for COULD NOT FIND EVENT!"
                            End If
                        Case EventType.evPlayAnimation
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]" & " on Player"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]" & " on Event #" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " [" & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3).Name) & "]"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 2 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]" & " on Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4 & ")"
                            End If
                        Case EventType.evCustomScript
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Execute Custom Script Case: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                        Case EventType.evPlayBGM
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play BGM [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1) & "]"
                        Case EventType.evFadeoutBGM
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Fadeout BGM"
                        Case EventType.evPlaySound
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Sound [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1) & "]"
                        Case EventType.evStopSound
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Stop Sound"
                        Case EventType.evOpenBank
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Open Bank"
                        Case EventType.evOpenShop
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Open Shop [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & ". " & Trim$(Shop(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).Name) & "]"
                        Case EventType.evSetAccess
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Access [" & frmEditor_Events.cmbSetAccess.List(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "]"
                        Case EventType.evGiveExp
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Give Player " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Experience."
                        Case Else
                            'Ghost
                            X = X - 1
                            If X = -1 Then
                                ReDim EventList(0)
                            Else
                                ReDim Preserve EventList(X)
                            End If
                    End Select
                End If
            End If
        Next
        If curlist > 1 Then
            X = X + 1
            ReDim Preserve EventList(X)
            EventList(X).CommandList = curlist
            EventList(X).CommandNum = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount + 1
            frmEditor_Events.lstCommands.AddItem indent & "@> "
            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList
            GoTo newlist
        End If
    End If
    
    frmEditor_Events.lstCommands.AddItem indent & "@> "
End Sub

Sub ListCommandAdd(s As String)
Static X As Long
    frmEditor_Events.lstCommands.AddItem s
    ' scrollbar
    If X < frmEditor_Events.TextWidth(s & "  ") Then
       X = frmEditor_Events.TextWidth(s & "  ")
      If frmEditor_Events.ScaleMode = vbTwips Then X = X / Screen.TwipsPerPixelX ' if twips change to pixels
      SendMessageByNum frmEditor_Events.lstCommands.hWnd, LB_SETHORIZONTALEXTENT, X, 0
    End If
End Sub

Sub AddCommand(Index As Long)
    Dim curlist As Long, i As Long, X As Long, curslot As Long, p As Long, oldCommandList As CommandListRec
    If tmpEvent.Pages(curPageNum).CommandListCount = 0 Then
        tmpEvent.Pages(curPageNum).CommandListCount = 1
        ReDim tmpEvent.Pages(curPageNum).CommandList(1)
    End If
    
    If frmEditor_Events.lstCommands.ListIndex = frmEditor_Events.lstCommands.ListCount - 1 Then
        curlist = 1
    Else
        curlist = EventList(frmEditor_Events.lstCommands.ListIndex).CommandList
    End If
        
    If tmpEvent.Pages(curPageNum).CommandListCount = 0 Then
        tmpEvent.Pages(curPageNum).CommandListCount = 1
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist)
    End If
    
    oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
    tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount + 1
    p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
    If p <= 0 Then
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
    Else
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(1 To p)
        tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
        For i = 1 To p - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i) = oldCommandList.Commands(i)
        Next
    End If
    
    If frmEditor_Events.lstCommands.ListIndex = frmEditor_Events.lstCommands.ListCount - 1 Then
        curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
    Else
        i = EventList(frmEditor_Events.lstCommands.ListIndex).CommandNum
        If i < tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
            For X = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1 To i Step -1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X + 1) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X)
            Next
            curslot = EventList(frmEditor_Events.lstCommands.ListIndex).CommandNum
        Else
            curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        End If
    End If
    
    
    Select Case Index
        Case EventType.evAddText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtAddText_Text.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlAddText_Colour.Value
            If frmEditor_Events.optAddText_Map.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optAddText_Global.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optAddText_Player.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
        Case EventType.evCondition
            'This is the part where the whole entire source goes to hell :D
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandListCount = tmpEvent.Pages(curPageNum).CommandListCount + 2
            ReDim Preserve tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.CommandList = tmpEvent.Pages(curPageNum).CommandListCount - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.ElseCommandList = tmpEvent.Pages(curPageNum).CommandListCount
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.CommandList).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.ElseCommandList).ParentList = curlist
            For i = 0 To 6
                If frmEditor_Events.optCondition_Index(i).Value = True Then X = i
            Next
            
            Select Case X
                Case 0 'Player Var
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 0
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = Val(frmEditor_Events.txtCondition_PlayerVarCondition.text)
                Case 1 'Player Switch
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex
                Case 2 'Has Item
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 2
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_HasItem.ListIndex + 1
                Case 3 'Class Is
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 3
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_ClassIs.ListIndex + 1
                Case 4 'Learnt Skill
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 4
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_LearntSkill.ListIndex + 1
                Case 5 'Level Is
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 5
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = Val(frmEditor_Events.txtCondition_LevelAmount.text)
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_LevelCompare.ListIndex
                Case 6 'Self Switch
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 6
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SelfSwitch.ListIndex
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex
            End Select
        Case EventType.evShowText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtShowText.text
        Case EventType.evShowChoices
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChoicePrompt.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2 = frmEditor_Events.txtChoices(1).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3 = frmEditor_Events.txtChoices(2).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4 = frmEditor_Events.txtChoices(3).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5 = frmEditor_Events.txtChoices(4).text
            
            tmpEvent.Pages(curPageNum).CommandListCount = tmpEvent.Pages(curPageNum).CommandListCount + 4
            ReDim Preserve tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = tmpEvent.Pages(curPageNum).CommandListCount - 3
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = tmpEvent.Pages(curPageNum).CommandListCount - 2
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = tmpEvent.Pages(curPageNum).CommandListCount - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = tmpEvent.Pages(curPageNum).CommandListCount
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 3).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 2).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 1).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount).ParentList = curlist
        Case EventType.evPlayerVar
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbVariable.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = Val(frmEditor_Events.txtVariable.text)
        Case EventType.evPlayerSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSwitch.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbPlayerSwitchSet.ListIndex
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetSelfSwitch.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbSetSelfSwitchTo.ListIndex
        Case EventType.evExitProcess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evChangeItems
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeItemIndex.ListIndex + 1
            If frmEditor_Events.optChangeItemSet.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeItemAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optChangeItemRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtChangeItemsAmount.text)
        Case EventType.evRestoreHP
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evRestoreMP
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evLevelUp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evChangeLevel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeLevel.Value
        Case EventType.evChangeSkills
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeSkills.ListIndex + 1
            If frmEditor_Events.optChangeSkillsAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeSkillsRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            End If
        Case EventType.evChangeClass
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeClass.ListIndex + 1
        Case EventType.evChangeSprite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeSprite.Value
        Case EventType.evChangeSex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            If frmEditor_Events.optChangeSexMale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            ElseIf frmEditor_Events.optChangeSexFemale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            End If
        Case EventType.evChangePK
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            If frmEditor_Events.optChangePKYes.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            ElseIf frmEditor_Events.optChangePKNo.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            End If
        Case EventType.evWarpPlayer
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWPMap.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlWPX.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlWPY.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.cmbWarpPlayerDir.ListIndex
        Case EventType.evSetMoveRoute
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = ListOfEvents(frmEditor_Events.cmbEvent.ListIndex)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.chkIgnoreMove.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.chkRepeatRoute.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount = TempMoveRouteCount
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute = TempMoveRoute
        Case EventType.evPlayAnimation
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbPlayAnim.ListIndex + 1
            If frmEditor_Events.optPlayAnimPlayer.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optPlayAnimEvent.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.cmbPlayAnimEvent.ListIndex + 1
            ElseIf frmEditor_Events.optPlayAnimTile.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlPlayAnimTileX.Value
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlPlayAnimTileY.Value
            End If
        Case EventType.evCustomScript
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlCustomScript.Value
        Case EventType.evPlayBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = musicCache(frmEditor_Events.cmbPlayBGM.ListIndex + 1)
        Case EventType.evFadeoutBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evPlaySound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = soundCache(frmEditor_Events.cmbPlaySound.ListIndex + 1)
        Case EventType.evStopSound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evOpenBank
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evOpenShop
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbOpenShop.ListIndex + 1
        Case EventType.evSetAccess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetAccess.ListIndex
        Case EventType.evGiveExp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlGiveExp.Value
    End Select
    EventListCommands
End Sub

Public Sub EditEventCommand()
    Dim i As Long, X As Long, Z As Long, curlist As Long, curslot As Long
    i = frmEditor_Events.lstCommands.ListIndex
    If i = -1 Then Exit Sub
    
    If i > UBound(EventList) Then Exit Sub

    curlist = EventList(i).CommandList
    curslot = EventList(i).CommandNum
    
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    
    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index
        Case EventType.evAddText
            isEdit = True
            frmEditor_Events.txtAddText_Text.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.scrlAddText_Colour.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
                Case 0
                    frmEditor_Events.optAddText_Player.Value = True
                Case 1
                    frmEditor_Events.optAddText_Map.Value = True
                Case 2
                    frmEditor_Events.optAddText_Global.Value = True
            End Select
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraAddText.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evCondition
            isEdit = True
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraConditionalBranch.Visible = True
            frmEditor_Events.fraCommands.Visible = False
            frmEditor_Events.ClearConditionFrame
            frmEditor_Events.optCondition_Index(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition).Value = True
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition
                Case 0
                    frmEditor_Events.cmbCondition_PlayerVarIndex.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerVarCompare.Enabled = True
                    frmEditor_Events.txtCondition_PlayerVarCondition.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                    frmEditor_Events.txtCondition_PlayerVarCondition.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3
                Case 1
                    frmEditor_Events.cmbCondition_PlayerSwitch.Enabled = True
                    frmEditor_Events.cmbCondtion_PlayerSwitchCondition.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                Case 2
                    frmEditor_Events.cmbCondition_HasItem.Enabled = True
                    frmEditor_Events.cmbCondition_HasItem.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                Case 3
                    frmEditor_Events.cmbCondition_ClassIs.Enabled = True
                    frmEditor_Events.cmbCondition_ClassIs.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                Case 4
                    frmEditor_Events.cmbCondition_LearntSkill.Enabled = True
                    frmEditor_Events.cmbCondition_LearntSkill.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                Case 5
                    frmEditor_Events.cmbCondition_LevelCompare.Enabled = True
                    frmEditor_Events.txtCondition_LevelAmount.Enabled = True
                    frmEditor_Events.txtCondition_LevelAmount.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                    frmEditor_Events.cmbCondition_LevelCompare.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1
                Case 6
                    frmEditor_Events.cmbCondition_SelfSwitch.Enabled = True
                    frmEditor_Events.cmbCondition_SelfSwitchCondition.Enabled = True
                    frmEditor_Events.cmbCondition_SelfSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1
                    frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
            End Select
        Case EventType.evShowText
            isEdit = True
            frmEditor_Events.txtShowText.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraShowText.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evShowChoices
            isEdit = True
            frmEditor_Events.txtChoicePrompt.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.txtChoices(1).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2
            frmEditor_Events.txtChoices(2).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3
            frmEditor_Events.txtChoices(3).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4
            frmEditor_Events.txtChoices(4).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraShowChoices.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayerVar
            isEdit = True
            frmEditor_Events.cmbVariable.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.txtVariable.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraPlayerVar.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayerSwitch
            isEdit = True
            frmEditor_Events.cmbSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.cmbPlayerSwitchSet.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraPlayerSwitch.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSelfSwitch
            isEdit = True
            frmEditor_Events.cmbSetSelfSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.cmbSetSelfSwitchTo.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraSelfSwitch.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeItems
            isEdit = True
            frmEditor_Events.cmbChangeItemIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0 Then
                frmEditor_Events.optChangeItemSet.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1 Then
                frmEditor_Events.optChangeItemAdd.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2 Then
                frmEditor_Events.optChangeItemRemove.Value = True
            End If
            frmEditor_Events.txtChangeItemsAmount.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraChangeItems.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeLevel
            isEdit = True
            frmEditor_Events.scrlChangeLevel.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraChangeLevel.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSkills
            isEdit = True
            frmEditor_Events.cmbChangeSkills.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0 Then
                frmEditor_Events.optChangeSkillsAdd.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1 Then
                frmEditor_Events.optChangeSkillsRemove.Value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraChangeSkills.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeClass
            isEdit = True
            frmEditor_Events.cmbChangeClass.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraChangeClass.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSprite
            isEdit = True
            frmEditor_Events.scrlChangeSprite.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraChangeSprite.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSex
            isEdit = True
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0 Then
                frmEditor_Events.optChangeSexMale.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1 Then
                frmEditor_Events.optChangeSexFemale.Value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraChangeSex.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangePK
            isEdit = True
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1 Then
                frmEditor_Events.optChangePKYes.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0 Then
                frmEditor_Events.optChangePKNo.Value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraChangePK.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evWarpPlayer
            isEdit = True
            frmEditor_Events.scrlWPMap.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlWPX.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.scrlWPY.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.cmbWarpPlayerDir.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraWarpPlayer.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetMoveRoute
            isEdit = True
            frmEditor_Events.fraMoveRoute.Visible = True
            frmEditor_Events.lstMoveRoute.Clear
            frmEditor_Events.cmbEvent.Clear
            ReDim ListOfEvents(0 To Map.EventCount)
            ListOfEvents(0) = EditorEvent
            frmEditor_Events.cmbEvent.AddItem "This Event"
            frmEditor_Events.cmbEvent.ListIndex = 0
            frmEditor_Events.cmbEvent.Enabled = True
            
            For i = 1 To Map.EventCount
                If i <> EditorEvent Then
                    frmEditor_Events.cmbEvent.AddItem Trim$(Map.Events(i).Name)
                    X = X + 1
                    ListOfEvents(X) = i
                    If i = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 Then frmEditor_Events.cmbEvent.ListIndex = X
                End If
            Next
            
                
            IsMoveRouteCommand = True
                
            frmEditor_Events.chkIgnoreMove.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.chkRepeatRoute.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                
            TempMoveRouteCount = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount
            TempMoveRoute = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute
            
            For i = 1 To TempMoveRouteCount
                Select Case TempMoveRoute(i).Index
                    Case 1
                        frmEditor_Events.lstMoveRoute.AddItem "Move Up"
                    Case 2
                        frmEditor_Events.lstMoveRoute.AddItem "Move Down"
                    Case 3
                        frmEditor_Events.lstMoveRoute.AddItem "Move Left"
                    Case 4
                        frmEditor_Events.lstMoveRoute.AddItem "Move Right"
                    Case 5
                        frmEditor_Events.lstMoveRoute.AddItem "Move Randomly"
                    Case 6
                        frmEditor_Events.lstMoveRoute.AddItem "Move Towards Player"
                    Case 7
                        frmEditor_Events.lstMoveRoute.AddItem "Move Away From Player"
                    Case 8
                        frmEditor_Events.lstMoveRoute.AddItem "Step Forward"
                    Case 9
                        frmEditor_Events.lstMoveRoute.AddItem "Step Back"
                    Case 10
                        frmEditor_Events.lstMoveRoute.AddItem "Wait 100ms"
                    Case 11
                        frmEditor_Events.lstMoveRoute.AddItem "Wait 500ms"
                    Case 12
                        frmEditor_Events.lstMoveRoute.AddItem "Wait 1000ms"
                    Case 13
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Up"
                    Case 14
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Down"
                    Case 15
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Left"
                    Case 16
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Right"
                    Case 17
                        frmEditor_Events.lstMoveRoute.AddItem "Turn 90 Degrees To the Right"
                    Case 18
                        frmEditor_Events.lstMoveRoute.AddItem "Turn 90 Degrees To the Left"
                    Case 19
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Around 180 Degrees"
                    Case 20
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Randomly"
                    Case 21
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Towards Player"
                    Case 22
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Away from Player"
                    Case 23
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 8x Slower"
                    Case 24
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 4x Slower"
                    Case 25
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 2x Slower"
                    Case 26
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed to Normal"
                    Case 27
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 2x Faster"
                    Case 28
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 4x Faster"
                    Case 29
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Lowest"
                    Case 30
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Lower"
                    Case 31
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Normal"
                    Case 32
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Higher"
                    Case 33
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Highest"
                    Case 34
                        frmEditor_Events.lstMoveRoute.AddItem "Turn On Walking Animation"
                    Case 35
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Off Walking Animation"
                    Case 36
                        frmEditor_Events.lstMoveRoute.AddItem "Turn On Fixed Direction"
                    Case 37
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Off Fixed Direction"
                    Case 38
                        frmEditor_Events.lstMoveRoute.AddItem "Turn On Walk Through"
                    Case 39
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Off Walk Through"
                    Case 40
                        frmEditor_Events.lstMoveRoute.AddItem "Set Position Below Player"
                    Case 41
                        frmEditor_Events.lstMoveRoute.AddItem "Set Position at Player Level"
                    Case 42
                        frmEditor_Events.lstMoveRoute.AddItem "Set Position Above Player"
                    Case 43
                        frmEditor_Events.lstMoveRoute.AddItem "Set Graphic"
                End Select
            Next
                
            frmEditor_Events.fraMoveRoute.Width = 841
            frmEditor_Events.fraMoveRoute.Height = 609
            frmEditor_Events.fraMoveRoute.Visible = True
            
            frmEditor_Events.fraDialogue.Visible = False
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayAnimation
            isEdit = True
            frmEditor_Events.lblPlayAnimX.Visible = False
            frmEditor_Events.lblPlayAnimY.Visible = False
            frmEditor_Events.scrlPlayAnimTileX.Visible = False
            frmEditor_Events.scrlPlayAnimTileY.Visible = False
            frmEditor_Events.cmbPlayAnimEvent.Visible = False
            frmEditor_Events.cmbPlayAnim.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.cmbPlayAnimEvent.Clear
            For i = 1 To Map.EventCount
                frmEditor_Events.cmbPlayAnimEvent.AddItem i & ". " & Trim$(Map.Events(i).Name)
            Next
            frmEditor_Events.cmbPlayAnimEvent.ListIndex = 0
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0 Then
                frmEditor_Events.optPlayAnimPlayer.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1 Then
                frmEditor_Events.optPlayAnimEvent.Value = True
                frmEditor_Events.cmbPlayAnimEvent.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 - 1
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2 Then
                frmEditor_Events.optPlayAnimTile.Value = True
                frmEditor_Events.scrlPlayAnimTileX.Max = Map.maxX
                frmEditor_Events.scrlPlayAnimTileY.Max = Map.maxY
                frmEditor_Events.scrlPlayAnimTileX.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                frmEditor_Events.scrlPlayAnimTileY.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraPlayAnimation.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evCustomScript
            isEdit = True
            frmEditor_Events.scrlCustomScript.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCustomScript.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayBGM
            isEdit = True
            For i = 1 To UBound(musicCache())
                If musicCache(i) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 Then
                    frmEditor_Events.cmbPlayBGM.ListIndex = i - 1
                End If
            Next
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraPlayBGM.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlaySound
            isEdit = True
            For i = 1 To UBound(soundCache())
                If soundCache(i) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 Then
                    frmEditor_Events.cmbPlaySound.ListIndex = i - 1
                End If
            Next
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraPlaySound.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evOpenShop
            isEdit = True
            frmEditor_Events.cmbOpenShop.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraOpenShop.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetAccess
            isEdit = True
            frmEditor_Events.cmbSetAccess.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraSetAccess.Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evGiveExp
            isEdit = True
            frmEditor_Events.scrlGiveExp.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.lblGiveExp.Caption = "Give Exp: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraGiveExp.Visible = True
            frmEditor_Events.fraCommands.Visible = False
    End Select
    
End Sub

Public Sub DeleteEventCommand()
    Dim i As Long, X As Long, Z As Long, curlist As Long, curslot As Long, p As Long, oldCommandList As CommandListRec
    i = frmEditor_Events.lstCommands.ListIndex
    If i = -1 Then Exit Sub
    
    curlist = EventList(i).CommandList
    curslot = EventList(i).CommandNum
    
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    
    If curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1
        p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        If p <= 0 Then
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
        Else
            oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(p)
            X = 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
            tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
            For i = 1 To p + 1
                If i <> curslot Then
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X) = oldCommandList.Commands(i)
                    X = X + 1
                End If
            Next
        End If
    Else
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1
        p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
        X = 1
        If p <= 0 Then
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
        Else
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(p)
            tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
            tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
            For i = 1 To p + 1
                If i <> curslot Then
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X) = oldCommandList.Commands(i)
                    X = X + 1
                End If
            Next
        End If
    End If
    
    EventListCommands
    
End Sub

Public Sub ClearEventCommands()
    ReDim tmpEvent.Pages(curPageNum).CommandList(1)
    tmpEvent.Pages(curPageNum).CommandListCount = 1
    EventListCommands
End Sub

Public Sub EditCommand()
    Dim i As Long, X As Long, Z As Long, curlist As Long, curslot As Long
    i = frmEditor_Events.lstCommands.ListIndex
    If i = -1 Then Exit Sub
    
    If i > UBound(EventList) Then Exit Sub

    curlist = EventList(i).CommandList
    curslot = EventList(i).CommandNum
    
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    
    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index
        Case EventType.evAddText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtAddText_Text.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlAddText_Colour.Value
            If frmEditor_Events.optAddText_Player.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optAddText_Map.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optAddText_Global.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
        Case EventType.evCondition
            If frmEditor_Events.optCondition_Index(0).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 0
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = Val(frmEditor_Events.txtCondition_PlayerVarCondition.text)
            ElseIf frmEditor_Events.optCondition_Index(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_HasItem.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(3).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 3
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_ClassIs.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(4).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 4
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_LearntSkill.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(5).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 5
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = Val(frmEditor_Events.txtCondition_LevelAmount.text)
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_LevelCompare.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(6).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 6
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SelfSwitch.ListIndex
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex
            End If
        Case EventType.evShowText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtShowText.text
        Case EventType.evShowChoices
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChoicePrompt.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2 = frmEditor_Events.txtChoices(1).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3 = frmEditor_Events.txtChoices(2).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4 = frmEditor_Events.txtChoices(3).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5 = frmEditor_Events.txtChoices(4).text
        Case EventType.evPlayerVar
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbVariable.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = Val(frmEditor_Events.txtVariable.text)
        Case EventType.evPlayerSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSwitch.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbPlayerSwitchSet.ListIndex
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetSelfSwitch.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbSetSelfSwitchTo.ListIndex
        Case EventType.evChangeItems
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeItemIndex.ListIndex + 1
            If frmEditor_Events.optChangeItemSet.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeItemAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optChangeItemRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtChangeItemsAmount.text)
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeLevel.Value
        Case EventType.evChangeSkills
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeSkills.ListIndex + 1
            If frmEditor_Events.optChangeSkillsAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeSkillsRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            End If
        Case EventType.evChangeClass
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeClass.ListIndex + 1
        Case EventType.evChangeSprite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeSprite.Value
        Case EventType.evChangeSex
            If frmEditor_Events.optChangeSexMale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            ElseIf frmEditor_Events.optChangeSexFemale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            End If
        Case EventType.evChangePK
            If frmEditor_Events.optChangePKYes.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            ElseIf frmEditor_Events.optChangePKNo.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            End If
        Case EventType.evWarpPlayer
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWPMap.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlWPX.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlWPY.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.cmbWarpPlayerDir.ListIndex
        Case EventType.evSetMoveRoute
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = ListOfEvents(frmEditor_Events.cmbEvent.ListIndex)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.chkIgnoreMove.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.chkRepeatRoute.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount = TempMoveRouteCount
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute = TempMoveRoute
        Case EventType.evPlayAnimation
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbPlayAnim.ListIndex + 1
            If frmEditor_Events.optPlayAnimPlayer.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optPlayAnimEvent.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.cmbPlayAnimEvent.ListIndex + 1
            ElseIf frmEditor_Events.optPlayAnimTile.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlPlayAnimTileX.Value
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlPlayAnimTileY.Value
            End If
        Case EventType.evCustomScript
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlCustomScript.Value
        Case EventType.evPlayBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = musicCache(frmEditor_Events.cmbPlayBGM.ListIndex + 1)
        Case EventType.evPlaySound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = soundCache(frmEditor_Events.cmbPlaySound.ListIndex + 1)
        Case EventType.evOpenShop
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbOpenShop.ListIndex + 1
        Case EventType.evSetAccess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetAccess.ListIndex
        Case EventType.evGiveExp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlGiveExp.Value
    End Select
    
    EventListCommands
End Sub

Sub RequestSwitchesAndVariables()
Dim i As Long, Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CRequestSwitchesAndVariables
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Sub SendSwitchesAndVariables()
Dim i As Long, Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwitchesAndVariables
    For i = 1 To MAX_SWITCHES
        Buffer.WriteString Switches(i)
    Next
    For i = 1 To MAX_VARIABLES
        Buffer.WriteString Variables(i)
    Next
    SendData Buffer.ToArray
Set Buffer = Nothing
End Sub
