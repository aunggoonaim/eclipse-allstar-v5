Attribute VB_Name = "modDirectDraw7"
Option Explicit
' Cursor
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private oldCursor As Long
Private newCursor As Long
Private Const GCL_HCURSOR = (-12)

' **********************
' ** Renders graphics **
' **********************
' DirectDraw7 Object
Public DD As DirectDraw7
' Clipper object
Public DD_Clip As DirectDrawClipper

' primary surface
Public DDS_Primary As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

' back buffer
Public DDS_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

' Used for pre-rendering
Public DDS_Map As DirectDrawSurface7
Public DDSD_Map As DDSURFACEDESC2

' gfx buffers
Public DDS_Item() As DirectDrawSurface7 ' arrays
Public DDS_Character() As DirectDrawSurface7
Public DDS_Paperdoll() As DirectDrawSurface7
Public DDS_Tileset() As DirectDrawSurface7
Public DDS_Resource() As DirectDrawSurface7
Public DDS_Animation() As DirectDrawSurface7
Public DDS_SpellIcon() As DirectDrawSurface7
Public DDS_Face() As DirectDrawSurface7
Public DDS_Projectile() As DirectDrawSurface7 ' projectiles
Public DDS_Door As DirectDrawSurface7 ' singes
Public DDS_Blood As DirectDrawSurface7
Public DDS_Misc As DirectDrawSurface7
Public DDS_Direction As DirectDrawSurface7
Public DDS_Target As DirectDrawSurface7
Public DDS_Bars As DirectDrawSurface7
Public DDS_Snow As DirectDrawSurface7
Public DDS_Bird As DirectDrawSurface7
Public DDS_Sand As DirectDrawSurface7
Public DDS_Fire As DirectDrawSurface7
Public DDS_MiniMap As DirectDrawSurface7
Public DDS_Event As DirectDrawSurface7

' descriptions
Public DDSD_Temp As DDSURFACEDESC2 ' arrays
Public DDSD_Item() As DDSURFACEDESC2
Public DDSD_Character() As DDSURFACEDESC2
Public DDSD_Paperdoll() As DDSURFACEDESC2
Public DDSD_Tileset() As DDSURFACEDESC2
Public DDSD_Resource() As DDSURFACEDESC2
Public DDSD_Animation() As DDSURFACEDESC2
Public DDSD_SpellIcon() As DDSURFACEDESC2
Public DDSD_Face() As DDSURFACEDESC2
Public DDSD_Projectile() As DDSURFACEDESC2 ' projectiles
Public DDSD_Door As DDSURFACEDESC2 ' singles
Public DDSD_Blood As DDSURFACEDESC2
Public DDSD_Misc As DDSURFACEDESC2
Public DDSD_Direction As DDSURFACEDESC2
Public DDSD_Target As DDSURFACEDESC2
Public DDSD_Bars As DDSURFACEDESC2
Public DDSD_Snow As DDSURFACEDESC2
Public DDSD_Bird As DDSURFACEDESC2
Public DDSD_Sand As DDSURFACEDESC2
Public DDSD_MiniMap As DDSURFACEDESC2
Public DDSD_Event As DDSURFACEDESC2

' timers
Public Const SurfaceTimerMax As Long = 10000
Public CharacterTimer() As Long
Public PaperdollTimer() As Long
Public ItemTimer() As Long
Public ResourceTimer() As Long
Public AnimationTimer() As Long
Public SpellIconTimer() As Long
Public FaceTimer() As Long

' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public NumItems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Private cImage As c32bppDIB
Public NumProjectiles As Long ' projectiles


' ********************
' ** Initialization **
' ********************
Public Function InitDirectDraw() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear DD7
    Call DestroyDirectDraw
    
    ' Init Direct Draw
    Set DD = DX7.DirectDrawCreate(vbNullString)
    
    ' Windowed
    DD.SetCooperativeLevel frmMain.hwnd, DDSCL_NORMAL

    ' Init type and set the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        .lBackBufferCount = 1
    End With
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMain.picScreen.hwnd
    
    ' Have the blits to the screen clipped to the picture box
    DDS_Primary.SetClipper DD_Clip
    
    ' Initialise the surfaces
    InitSurfaces
    
    ' We're done
    InitDirectDraw = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InitDirectDraw", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub InitSurfaces()
Dim rec As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' DirectDraw Surface memory management setting
    DDSD_Temp.lFlags = DDSD_CAPS
    DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    
    ' clear out everything for re-init
    Set DDS_BackBuffer = Nothing

    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (MAX_MAPX + 3) * PIC_X
        .lHeight = (MAX_MAPY + 3) * PIC_Y
    End With
    Set DDS_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    ' load persistent surfaces
    If FileExist(App.Path & "\data files\graphics\door.png", True) Then Call InitDDSurf("door", DDSD_Door, DDS_Door)
    If FileExist(App.Path & "\data files\graphics\direction.png", True) Then Call InitDDSurf("direction", DDSD_Direction, DDS_Direction)
    If FileExist(App.Path & "\data files\graphics\target.png", True) Then Call InitDDSurf("target", DDSD_Target, DDS_Target)
    If FileExist(App.Path & "\data files\graphics\misc.png", True) Then Call InitDDSurf("misc", DDSD_Misc, DDS_Misc)
    If FileExist(App.Path & "\data files\graphics\blood.png", True) Then Call InitDDSurf("blood", DDSD_Blood, DDS_Blood)
    If FileExist(App.Path & "\data files\graphics\bars.png", True) Then Call InitDDSurf("bars", DDSD_Bars, DDS_Bars)
    If FileExist(App.Path & "\data files\graphics\snow.png", True) Then Call InitDDSurf("snow", DDSD_Snow, DDS_Snow)
    If FileExist(App.Path & "\data files\graphics\bird.png", True) Then Call InitDDSurf("bird", DDSD_Bird, DDS_Bird)
    If FileExist(App.Path & "\data files\graphics\sand.png", True) Then Call InitDDSurf("sand", DDSD_Sand, DDS_Sand)
    If FileExist(App.Path & "\data files\graphics\fire.png", True) Then Call InitDDSurf("fire", DDSD_Sand, DDS_Fire)
    If FileExist(App.Path & "\data files\graphics\minimap.png", True) Then Call InitDDSurf("minimap", DDSD_MiniMap, DDS_MiniMap)
    If FileExist(App.Path & "\data files\graphics\event.bmp", True) Then Call InitDDSurf("event", DDSD_Event, DDS_Event)
    
   
    ' count the blood sprites
    BloodCount = DDSD_Blood.lWidth / 32
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' This sub gets the mask color from the surface loaded from a bitmap image
Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With TmpR
        .Left = X
        .Top = Y
        .Right = X
        .Bottom = Y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .Low = TheSurface.GetLockedPixel(X, Y)
        .High = .Low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    TheSurface.Unlock TmpR
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetMaskColorFromPixel", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Initializing a surface, using a bitmap
Public Sub InitDDSurf(FileName As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Set path
    FileName = App.Path & GFX_PATH & FileName & GFX_EXT

    ' Destroy surface if it exist
    If Not Surf Is Nothing Then
        Set Surf = Nothing
        Call ZeroMemory(ByVal VarPtr(SurfDesc), LenB(SurfDesc))
    End If

        ' set flags
    SurfDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    SurfDesc.ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps

        ' init object
        ' Set Surf = DD.CreateSurfaceFromFile(FileName, SurfDesc)
    Set cImage = New c32bppDIB
    cImage.LoadPicture_File (FileName)
    SurfDesc.lWidth = cImage.Width
    SurfDesc.lHeight = cImage.Height
    Set Surf = DD.CreateSurface(SurfDesc)
    Dim DC
    DC = Surf.GetDC
    cImage.Render (DC)
    Surf.ReleaseDC (DC)
    cImage.DestroyDIB
    
    ' Set mask
    Call SetMaskColorFromPixel(Surf, 0, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitDDSurf", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function CheckSurfaces() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if we need to restore surfaces
    If Not DD.TestCooperativeLevel = DD_OK Then
        CheckSurfaces = False
    Else
        CheckSurfaces = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "CheckSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function NeedToRestoreSurfaces() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not DD.TestCooperativeLevel = DD_OK Then
        NeedToRestoreSurfaces = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "NeedToRestoreSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub ReInitDD()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call InitDirectDraw
    
    LoadTilesets
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ReInitDD", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyDirectDraw()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Unload DirectDraw
    Set DDS_Misc = Nothing
    
    For i = 1 To NumTileSets
        Set DDS_Tileset(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i))
    Next

    For i = 1 To NumItems
        Set DDS_Item(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i))
    Next

    For i = 1 To NumCharacters
        Set DDS_Character(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i))
    Next
    
    For i = 1 To NumPaperdolls
        Set DDS_Paperdoll(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i))
    Next
    
    For i = 1 To NumResources
        Set DDS_Resource(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i))
    Next
    
    For i = 1 To NumAnimations
        Set DDS_Animation(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i))
    Next
    
    For i = 1 To NumSpellIcons
        Set DDS_SpellIcon(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i))
    Next
    
    For i = 1 To NumFaces
        Set DDS_Face(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Face(i)), LenB(DDSD_Face(i))
    Next
    
    ' projectiles
    For i = 1 To NumProjectiles
        Set DDS_Projectile(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Projectile(i)), LenB(DDSD_Projectile(i))
    Next
    
    Set DDS_Blood = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Blood), LenB(DDSD_Blood)
    
    Set DDS_Door = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Door), LenB(DDSD_Door)
    
    Set DDS_Direction = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Direction), LenB(DDSD_Direction)
    
    Set DDS_Target = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Target), LenB(DDSD_Target)
    
    Set DDS_MiniMap = Nothing
    ZeroMemory ByVal VarPtr(DDSD_MiniMap), LenB(DDSD_MiniMap)

    Set DDS_BackBuffer = Nothing
    Set DDS_Primary = Nothing
    Set DD_Clip = Nothing
    Set DD = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyDirectDraw", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Blitting **
' **************
Public Sub Engine_BltFast(ByVal dX As Long, ByVal dY As Long, ByRef ddS As DirectDrawSurface7, srcRect As RECT, trans As CONST_DDBLTFASTFLAGS)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler


    If Not ddS Is Nothing Then
        Call DDS_BackBuffer.BltFast(dX, dY, ddS, srcRect, trans)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Engine_BltFast", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function Engine_BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Clear Then
        picBox.Cls
    End If

    Call Surface.BltToDC(picBox.hDC, sRECT, dRECT)
    picBox.Refresh
    Engine_BltToDC = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Engine_BltToDC", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub BltDirection(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' render grid
    rec.Top = 24
    rec.Left = 0
    rec.Right = rec.Left + 32
    rec.Bottom = rec.Top + 32
    Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Direction, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' render dir blobs
    For i = 1 To 4
        rec.Left = (i - 1) * 8
        rec.Right = rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(X, Y).DirBlock, CByte(i)) Then
            rec.Top = 8
        Else
            rec.Top = 16
        End If
        rec.Bottom = rec.Top + 8
        'render!
        Call Engine_BltFast(ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), DDS_Direction, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltDirection", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTarget(ByVal X As Long, ByVal Y As Long)
Dim sRECT As DxVBLib.RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DDS_Target Is Nothing Then Exit Sub
    
    Width = DDSD_Target.lWidth / 2
    Height = DDSD_Target.lHeight

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Target, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTarget", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltHover(ByVal tType As Long, ByVal target As Long, ByVal X As Long, ByVal Y As Long)
Dim sRECT As DxVBLib.RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DDS_Target Is Nothing Then Exit Sub
    
    ' ย้อนกลับเมาส์เป็นปกติ
    
    If oldCursor > 0 Then
        Call SetClassLong(frmMain.hwnd, GCL_HCURSOR, oldCursor)
        Call SetClassLong(frmMain.picScreen.hwnd, GCL_HCURSOR, oldCursor)
        DestroyCursor newCursor
        ' Cursor_Status = "None"
    End If
    
    ' cursor ภาพเมาส์ บน npc / ผู้เล่น
       newCursor = LoadCursorFromFile("data files\graphics\cursor\Attack.ani")
       
       If newCursor > 0 Then
            oldCursor = SetClassLong(frmMain.hwnd, GCL_HCURSOR, newCursor)
            oldCursor = SetClassLong(frmMain.picScreen.hwnd, GCL_HCURSOR, newCursor)
            ' Cursor_Status = "Hover"
       End If
    
    Width = DDSD_Target.lWidth / 2
    Height = DDSD_Target.lHeight

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = Width
        .Right = .Left + Width
    End With
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Target, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHover", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapTile(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Ground To MapLayer.Mask2
            ' skip tile?
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                ' sort out rec
                rec.Top = .Layer(i).Y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).X * PIC_X
                rec.Right = rec.Left + PIC_X
                
                'If i Mod 2 = 0 Then
                    'If LayerAnim Then
                        'render'
                        'Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    'End If
                'Else
                    'Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                'End If
                
                ' render'
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "BltMapTile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapFringeTile(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            ' skip tile if tileset isn't set
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                ' sort out rec
                rec.Top = .Layer(i).Y * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = .Layer(i).X * PIC_X
                rec.Right = rec.Left + PIC_X
                
                'If i Mod 2 = 0 Then
                    'If LayerAnim Then
                        ' render'
                        'Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    'End If
                'Else
                    'Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                'End If
            
                ' render'
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapFringeTile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltDoor(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim X2 As Long, Y2 As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' sort out animation
    With TempTile(X, Y)
        If .DoorAnimate = 1 Then ' opening
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame < 4 Then
                    .DoorFrame = .DoorFrame + 1
                Else
                    .DoorAnimate = 2 ' set to closing
                End If
                .DoorTimer = GetTickCount
            End If
        ElseIf .DoorAnimate = 2 Then ' closing
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame > 1 Then
                    .DoorFrame = .DoorFrame - 1
                Else
                    .DoorAnimate = 0 ' end animation
                End If
                .DoorTimer = GetTickCount
            End If
        End If
        
        If .DoorFrame = 0 Then .DoorFrame = 1
    End With

    With rec
        .Top = 0
        .Bottom = DDSD_Door.lHeight
        .Left = ((TempTile(X, Y).DoorFrame - 1) * (DDSD_Door.lWidth / 4))
        .Right = .Left + (DDSD_Door.lWidth / 4)
    End With

    X2 = (X * PIC_X)
    Y2 = (Y * PIC_Y) - (DDSD_Door.lHeight / 2) + 4
    Call DDS_BackBuffer.BltFast(ConvertMapX(X2), ConvertMapY(Y2), DDS_Door, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltDoor", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltBlood(ByVal Index As Long)
Dim rec As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Blood(Index)
        ' check if we should be seeing it
        If .Timer + 20000 < GetTickCount Then Exit Sub
        
        rec.Top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        Engine_BltFast ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), DDS_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBlood", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltAnimation(ByVal Index As Long, ByVal Layer As Long)
Dim Sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
Dim i As Long
Dim Width As Long, Height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim X As Long, Y As Long
Dim lockindex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    AnimationTimer(Sprite) = GetTickCount + SurfaceTimerMax
    
    If DDS_Animation(Sprite) Is Nothing Then
        Call InitDDSurf("animations\" & Sprite, DDSD_Animation(Sprite), DDS_Animation(Sprite))
    End If
    
    ' total width divided by frame count
    Width = DDSD_Animation(Sprite).lWidth / FrameCount
    Height = DDSD_Animation(Sprite).lHeight
    
    sRECT.Top = 0
    sRECT.Bottom = Height
    sRECT.Left = (AnimInstance(Index).FrameIndex(Layer) - 1) * Width
    sRECT.Right = sRECT.Left + Width
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).xOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).Num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).xOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).yOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    ' Clip to screen
    If Y < 0 Then

        With sRECT
            .Top = .Top - Y
        End With

        Y = 0
    End If

    If X < 0 Then

        With sRECT
            .Left = .Left - X
        End With

        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    
    Call Engine_BltFast(X, Y, DDS_Animation(Sprite), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltAnimation", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltItem(ByVal itemnum As Long)
Dim PicNum As Long
Dim rec As DxVBLib.RECT
Dim MaxFrames As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if it's not us then don't render
    If MapItem(itemnum).playerName <> vbNullString Then
        If MapItem(itemnum).playerName <> Trim$(GetPlayerName(MyIndex)) Then Exit Sub
    End If
    
    ' get the picture
    PicNum = Item(MapItem(itemnum).Num).Pic

    If PicNum < 1 Or PicNum > NumItems Then Exit Sub
    ItemTimer(PicNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(PicNum) Is Nothing Then
        Call InitDDSurf("items\" & PicNum, DDSD_Item(PicNum), DDS_Item(PicNum))
    End If

    If DDSD_Item(PicNum).lWidth > 64 Then ' has more than 1 frame
        With rec
            .Top = 0
            .Bottom = 32
            .Left = (MapItem(itemnum).Frame * 32)
            .Right = .Left + 32
        End With
    Else
        With rec
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If

    Call Engine_BltFast(ConvertMapX(MapItem(itemnum).X * PIC_X), ConvertMapY(MapItem(itemnum).Y * PIC_Y), DDS_Item(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' player Projectiles
Public Sub BltProjectile(ByVal Index As Long, ByVal PlayerProjectile As Long)
Dim X As Long, Y As Long, PicNum As Long, i As Long
Dim rec As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check for subscript error
    If Index < 1 Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
    
    ' check to see if it's time to move the Projectile
    If GetTickCount > Player(Index).ProjecTile(PlayerProjectile).TravelTime Then
        With Player(Index).ProjecTile(PlayerProjectile)
            ' set next travel time and the current position and then set the actual direction based on RMXP arrow tiles.
            Select Case .Direction
                ' down
                Case 0
                    .Y = .Y + 1
                    ' check if they reached maxrange
                    If .Y = (GetPlayerY(Index) + .Range) + 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' up
                Case 1
                    .Y = .Y - 1
                    ' check if they reached maxrange
                    If .Y = (GetPlayerY(Index) - .Range) - 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' right
                Case 2
                    .X = .X + 1
                    ' check if they reached max range
                    If .X = (GetPlayerX(Index) + .Range) + 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' left
                Case 3
                    .X = .X - 1
                    ' check if they reached maxrange
                    If .X = (GetPlayerX(Index) - .Range) - 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
            End Select
            .TravelTime = GetTickCount + .Speed
        End With
    End If
    
    ' set the x, y & pic values for future reference
    X = Player(Index).ProjecTile(PlayerProjectile).X
    Y = Player(Index).ProjecTile(PlayerProjectile).Y
    PicNum = Player(Index).ProjecTile(PlayerProjectile).Pic
    
    ' check if left map
    If X > Map.maxX Or Y > Map.maxY Or X < 0 Or Y < 0 Then
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if we hit a block
    If Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
    ' check for player hit
    For i = 1 To Player_HighIndex
        If X = GetPlayerX(i) And Y = GetPlayerY(i) Then
            ' they're hit, remove it
            If Not X = Player(MyIndex).X Or Not Y = GetPlayerY(MyIndex) Then
                ClearProjectile Index, PlayerProjectile
                Exit Sub
            End If
        End If
    Next
    
    ' check for npc hit
    For i = 1 To MAX_MAP_NPCS
        If X = MapNpc(i).X And Y = MapNpc(i).Y Then
            ' they're hit, remove it
            ClearProjectile Index, PlayerProjectile
            Exit Sub
        End If
    Next
    
    ' if projectile is not loaded, load it, bitch.
    If DDS_Projectile(PicNum) Is Nothing Then
        Call InitDDSurf("projectiles\" & PicNum, DDSD_Projectile(PicNum), DDS_Projectile(PicNum))
    End If
    
    ' get positioning in the texture
    With rec
        .Top = 0
        .Bottom = SIZE_Y
        .Left = Player(Index).ProjecTile(PlayerProjectile).Direction * SIZE_X
        .Right = .Left + SIZE_X
    End With

    ' blt the projectile
    Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Projectile(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltProjectile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ScreenshotMap()
Dim X As Long, Y As Long, i As Long, rec As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' clear the surface
    Set DDS_Map = Nothing
    
    ' Initialize it
    With DDSD_Map
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (Map.maxX + 1) * 32
        .lHeight = (Map.maxY + 1) * 32
    End With
    Set DDS_Map = DD.CreateSurface(DDSD_Map)
    
    ' render the tiles
    For X = 0 To Map.maxX
        For Y = 0 To Map.maxY
            With Map.Tile(X, Y)
                For i = MapLayer.Ground To MapLayer.Mask2
                    ' skip tile?
                    If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                        ' sort out rec
                        rec.Top = .Layer(i).Y * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        rec.Left = .Layer(i).X * PIC_X
                        rec.Right = rec.Left + PIC_X
                        ' render
                        DDS_Map.BltFast X * PIC_X, Y * PIC_Y, DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Next
            End With
        Next
    Next
    
    ' render the resources
    For Y = 0 To Map.maxY
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = Y Then
                            Call BltMapResource(i, True)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' render the tiles
    For X = 0 To Map.maxX
        For Y = 0 To Map.maxY
            With Map.Tile(X, Y)
                For i = MapLayer.Fringe To MapLayer.Fringe2
                    ' skip tile?
                    If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                        ' sort out rec
                        rec.Top = .Layer(i).Y * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        rec.Left = .Layer(i).X * PIC_X
                        rec.Right = rec.Left + PIC_X
                        ' render
                        DDS_Map.BltFast X * PIC_X, Y * PIC_Y, DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Next
            End With
        Next
    Next
    
    ' dump and save
    frmMain.picSSMap.Width = DDSD_Map.lWidth
    frmMain.picSSMap.Height = DDSD_Map.lHeight
    rec.Top = 0
    rec.Left = 0
    rec.Bottom = DDSD_Map.lHeight
    rec.Right = DDSD_Map.lWidth
    Engine_BltToDC DDS_Map, rec, rec, frmMain.picSSMap
    SavePicture frmMain.picSSMap.Image, App.Path & "\map" & GetPlayerMap(MyIndex) & ".jpg"
    
    ' let them know we did it
    AddText "Screenshot of map #" & GetPlayerMap(MyIndex) & " saved.", BrightGreen
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotMap", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' DrawMapResource

Public Sub BltMapResource(ByVal Resource_num As Long, Optional ByVal screenShot As Boolean = False)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As DxVBLib.RECT
Dim X As Long, Y As Long
Dim i As Long, Alpha As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).X > Map.maxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.maxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).Data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' Load early
    If DDS_Resource(Resource_sprite) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource_sprite, DDSD_Resource(Resource_sprite), DDS_Resource(Resource_sprite))
    End If

    ' src rect
    With rec
        .Top = 0
        .Bottom = DDSD_Resource(Resource_sprite).lHeight
        .Left = 0
        .Right = DDSD_Resource(Resource_sprite).lWidth
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (DDSD_Resource(Resource_sprite).lWidth / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - DDSD_Resource(Resource_sprite).lHeight + 32
    
    ' render it
    If Not screenShot Then
        Call BltResource(Resource_sprite, X, Y, rec)
    Else
        Call ScreenshotResource(Resource_sprite, X, Y, rec)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltResource(ByVal Resource As Long, ByVal dX As Long, dY As Long, rec As DxVBLib.RECT)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long
Dim destRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    ResourceTimer(Resource) = GetTickCount + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If

    X = ConvertMapX(dX)
    Y = ConvertMapY(dY)
    
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If

    ' End clipping
    Call Engine_BltFast(X, Y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScreenshotResource(ByVal Resource As Long, ByVal X As Long, Y As Long, rec As DxVBLib.RECT)
Dim Width As Long
Dim Height As Long
Dim destRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    ResourceTimer(Resource) = GetTickCount + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If
    
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_Map.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_Map.lHeight)
    End If

    If X + Width > DDSD_Map.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_Map.lWidth)
    End If

    ' End clipping
    'Call Engine_BltFast(x, y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    DDS_Map.BltFast X, Y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRECT As RECT
Dim barWidth As Long
Dim i As Long, NPCNum As Long, partyIndex As Long
Dim SkillDelay As Double
Dim wPower As Double

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' dynamic bar calculations
    sWidth = DDSD_Bars.lWidth
    sHeight = DDSD_Bars.lHeight / 4
    
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        NPCNum = MapNpc(i).Num
        ' exists?
        If NPCNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < NPC(NPCNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).X * PIC_X + MapNpc(i).xOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(i).Y * PIC_Y + MapNpc(i).yOffset + 35
                
                ' calculate the width to fill
                barWidth = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (NPC(NPCNum).HP / sWidth)) * sWidth
                
                ' draw bar background
                With sRECT
                    .Top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .Top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                
                ' draw the bar proper
                With sRECT
                    .Top = 0 ' HP bar
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .Top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).yOffset + 35 + sHeight + 1
            
            SkillDelay = 1
            
            ' ตรวจสอบเงื่อนไข
                If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(MyIndex, Weapon)).DelayDown > 0 And Item(GetPlayerEquipment(MyIndex, Weapon)).DelayDown < 1 Then
                        SkillDelay = Item(GetPlayerEquipment(MyIndex, Weapon)).DelayDown
                        If SkillDelay <= 0 Then SkillDelay = 1
                    End If
                End If
            
            wPower = 1 + (GetPlayerStat(MyIndex, willpower) / 50)
            
            ' สูตรคำนวนเวลาร่ายสกิล V2
            barWidth = ((GetTickCount - SpellBufferTimer) / (((Spell(PlayerSpells(SpellBuffer)).CastTime * 100) * SkillDelay / wPower))) * sWidth ' ((Spell(PlayerSpells(SpellBuffer)).CastTime * 100)) * SkillDelay / (1 + Int(GetPlayerStat(MyIndex, willpower) / 50))) * sWidth
           ' Call AddText("Client Skill : ((" & Int(GetPlayerStat(MyIndex, willpower)) / 50 & "))", White)
            
            ' draw bar background
            With sRECT
                .Top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            
            ' draw the bar proper
            With sRECT
                .Top = sHeight * 2 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .Top + sHeight
            End With
            Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).yOffset + 35
       
        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
       
        ' draw bar background
        With sRECT
            .Top = sHeight * 1 ' HP bar background
            .Left = 0
            .Right = .Left + sWidth
            .Bottom = .Top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
       
        ' draw the bar proper
        With sRECT
            .Top = 0 ' HP bar
            .Left = 0
            .Right = .Left + barWidth
            .Bottom = .Top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).xOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).yOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' draw bar background
                    With sRECT
                        .Top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    
                    ' draw the bar proper
                    With sRECT
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .Top + sHeight
                    End With
                    Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            End If
        Next
    End If
                    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBars", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltHotbar()
Dim sRECT As RECT, dRECT As RECT, i As Long, Num As String, n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.picHotbar.Cls
    
    For i = 1 To MAX_HOTBAR
        With dRECT
            .Top = HotbarTop
            .Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
            .Bottom = .Top + 32
            .Right = .Left + 32
        End With
        
        With sRECT
            .Top = 0
            .Left = 32
            .Bottom = 32
            .Right = 64
        End With
        
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(i).Slot).Name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 Then
                        If DDS_Item(Item(Hotbar(i).Slot).Pic) Is Nothing Then
                            Call InitDDSurf("Items\" & Item(Hotbar(i).Slot).Pic, DDSD_Item(Item(Hotbar(i).Slot).Pic), DDS_Item(Item(Hotbar(i).Slot).Pic))
                        End If
                        Engine_BltToDC DDS_Item(Item(Hotbar(i).Slot).Pic), sRECT, dRECT, frmMain.picHotbar, False
                    End If
                End If
            Case 2 ' spell
                With sRECT
                    .Top = 0
                    .Left = 0
                    .Bottom = 32
                    .Right = 32
                End With
                If Len(Spell(Hotbar(i).Slot).Name) > 0 Then
                    If Spell(Hotbar(i).Slot).Icon > 0 Then
                        If DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon) Is Nothing Then
                            Call InitDDSurf("Spellicons\" & Spell(Hotbar(i).Slot).Icon, DDSD_SpellIcon(Spell(Hotbar(i).Slot).Icon), DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon))
                        End If
                        ' check for cooldown
                        For n = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(n) = (Hotbar(i).Slot) Then
                                ' has spell
                                If Not SpellCD(n) = 0 Then
                                    sRECT.Left = 32
                                    sRECT.Right = 64
                                    ' ถ้าสกิลดีเลย์
                                End If
                            End If
                        Next
                        Engine_BltToDC DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon), sRECT, dRECT, frmMain.picHotbar, False
                    End If
                End If
        End Select
        
        ' แก้ไขคีย์ลัด by Allstar
        If i <= 10 Then
            Num = "" & str(i) ' ใน " " ปกติคือ F เช่นพวก F1 - F12
            DrawText frmMain.picHotbar.hDC, dRECT.Left + 1, dRECT.Top + 1, Num, QBColor(White)
        ElseIf i = 11 Then
            Num = "-" ' & str(i)
            DrawText frmMain.picHotbar.hDC, dRECT.Left + 1, dRECT.Top + 1, Num, QBColor(White)
        ElseIf i = 12 Then
            Num = "=" ' & str(i)
            DrawText frmMain.picHotbar.hDC, dRECT.Left + 1, dRECT.Top + 1, Num, QBColor(White)
        End If
        
    Next
    
    frmMain.picHotbar.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHotbar", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPlayer(ByVal Index As Long)
Dim anim As Byte, i As Long, X As Long, Y As Long
Dim Sprite As Long, spritetop As Long
Dim rec As DxVBLib.RECT
Dim AttackSpeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = GetPlayerSprite(Index)

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    ' ความเร็วการโจมตี
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        AttackSpeed = ((2000 + Item(GetPlayerEquipment(Index, Weapon)).Speed) - Item(GetPlayerEquipment(Index, Weapon)).SpeedLow) - ((GetPlayerStat(Index, Stats.Agility) * 5))
    Else
        AttackSpeed = 2000 - ((GetPlayerStat(Index, Stats.Agility) * 5))
    End If
    
    ' Fixed bug attackspeed high
        If AttackSpeed < 200 Then
            AttackSpeed = 200
        End If

    ' Reset frame
    If Player(Index).Step = 3 Then
        anim = 0
    ElseIf Player(Index).Step = 1 Then
        anim = 2
    End If
    
    ' Check for attacking animation
    If Player(Index).AttackTimer + (AttackSpeed / 2) > GetTickCount Then
        If Player(Index).Attacking = 1 Then
            anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).yOffset > 8) Then anim = Player(Index).Step
            Case DIR_DOWN
                If (Player(Index).yOffset < -8) Then anim = Player(Index).Step
            Case DIR_LEFT
                If (Player(Index).xOffset > 8) Then anim = Player(Index).Step
            Case DIR_RIGHT
                If (Player(Index).xOffset < -8) Then anim = Player(Index).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = spritetop * (DDSD_Character(Sprite).lHeight / 4)
        .Bottom = .Top + (DDSD_Character(Sprite).lHeight / 4)
        .Left = anim * (DDSD_Character(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 4)
    End With

    ' Calculate the X
    X = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((DDSD_Character(Sprite).lWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (DDSD_Character(Sprite).lHeight) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((DDSD_Character(Sprite).lHeight / 4) - 32)
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
    End If

    ' render the actual sprite
    Call BltSprite(Sprite, X, Y, rec)
    
    ' Fixed by Allstar
    If GetPlayerEquipment(Index, Shield) <> Player(Index).WieldDagger Then ' Get player daggers [ 1]
        If GetPlayerEquipment(Index, Armor) > 0 Then
            'Call DrawPaperdoll(x, Y, Item(GetPlayerEquipment(Index, Armor)).Paperdoll, anim, spritetop)
            Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Armor)).Paperdoll, anim, spritetop)
        End If
        
        If GetPlayerEquipment(Index, Helmet) > 0 Then
            Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Helmet)).Paperdoll, anim, spritetop)
        End If

        If Item(GetPlayerEquipment(Index, Shield)).Daggerpdoll > 0 Then
            If GetPlayerEquipment(Index, Shield) > 0 Then
                Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Shield)).Paperdoll, anim, spritetop)
            End If
        Else
            Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Shield)).Paperdoll, anim, spritetop)
        End If
    
        If GetPlayerEquipment(Index, Weapon) > 0 Then
            Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Weapon)).Paperdoll, anim, spritetop)
        End If
    
        Exit Sub
    End If
    
    'check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, anim, spritetop)
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayer", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltNpc(ByVal MapNpcNum As Long)
Dim anim As Byte, i As Long, X As Long, Y As Long, Sprite As Long, spritetop As Long
Dim rec As DxVBLib.RECT
Dim AttackSpeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).Num = 0 Then Exit Sub ' no npc set
    
    Sprite = NPC(MapNpc(MapNpcNum).Num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    If NPC(MapNpc(MapNpcNum).Num).AttackSpeed > 0 Then
        AttackSpeed = NPC(MapNpc(MapNpcNum).Num).AttackSpeed
    Else
        AttackSpeed = 3000
    End If
    
    ' Fixed bug attackspeed high
        If AttackSpeed < 100 Then
            AttackSpeed = 100
        End If
    
    ' Reset frame
    anim = 0
    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (AttackSpeed / 2) > GetTickCount Then
        If MapNpc(MapNpcNum).Attacking = 1 Then
            anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).yOffset > 8) Then anim = MapNpc(MapNpcNum).Step
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).yOffset < -8) Then anim = MapNpc(MapNpcNum).Step
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).xOffset > 8) Then anim = MapNpc(MapNpcNum).Step
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).xOffset < -8) Then anim = MapNpc(MapNpcNum).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (DDSD_Character(Sprite).lHeight / 4) * spritetop
        .Bottom = .Top + DDSD_Character(Sprite).lHeight / 4
        .Left = anim * (DDSD_Character(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 4)
    End With

    ' Calculate the X
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).xOffset - ((DDSD_Character(Sprite).lWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (DDSD_Character(Sprite).lHeight / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((DDSD_Character(Sprite).lHeight / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset
    End If

    Call BltSprite(Sprite, X, Y, rec)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltNpc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub BltPaperdoll(ByVal X2 As Long, ByVal Y2 As Long, ByVal Sprite As Long, ByVal anim As Long, ByVal spritetop As Long)
Dim rec As DxVBLib.RECT
Dim X As Long, Y As Long
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    If DDS_Paperdoll(Sprite) Is Nothing Then
        Call InitDDSurf("Paperdolls\" & Sprite, DDSD_Paperdoll(Sprite), DDS_Paperdoll(Sprite))
    End If
    
    With rec
        .Top = spritetop * (DDSD_Paperdoll(Sprite).lHeight / 4)
        .Bottom = .Top + (DDSD_Paperdoll(Sprite).lHeight / 4)
        .Left = anim * (DDSD_Paperdoll(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Paperdoll(Sprite).lWidth / 4)
    End With
    
    ' clipping
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' Clip to screen
    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Paperdoll(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPaperdoll", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltSprite(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, rec As DxVBLib.RECT)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' clipping
    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Character(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltAnimatedInvItems()
Dim i As Long
Dim itemnum As Long, itempic As Long
Dim X As Long, Y As Long
Dim MaxFrames As Byte
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For i = 1 To MAX_MAP_ITEMS

        If MapItem(i).Num > 0 Then
            itempic = Item(MapItem(i).Num).Pic

            If itempic < 1 Or itempic > NumItems Then Exit Sub
            MaxFrames = (DDSD_Item(itempic).lWidth / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(i).Frame < MaxFrames - 1 Then
                MapItem(i).Frame = MapItem(i).Frame + 1
            Else
                MapItem(i).Frame = 0
            End If
        End If

    Next

    For i = 1 To MAX_INV
        itemnum = GetPlayerInvItemNum(MyIndex, i)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= NumItems Then
                If DDSD_Item(itempic).lWidth > 64 Then
                    MaxFrames = (DDSD_Item(itempic).lWidth / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(i) < MaxFrames - 1 Then
                        InvItemFrame(i) = InvItemFrame(i) + 1
                    Else
                        InvItemFrame(i) = 0
                    End If

                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = (DDSD_Item(itempic).lWidth / 2) + (InvItemFrame(i) * 32) ' middle to get the start of inv gfx, then +32 for each frame
                        .Right = .Left + 32
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' Load item if not loaded, and reset timer
                    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_Item(itempic) Is Nothing Then
                        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                    End If

                    ' We'll now re-blt the item, and place the currency value over it again :P
                    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picInventory, False

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                        ' Draw currency but with k, m, b etc. using a convertion function
                        DrawText frmMain.picInventory.hDC, X, Y, ConvertCurrency(Amount), QBColor(Yellow)

                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If

    Next

    frmMain.picInventory.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltAnimatedInvItems", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltFace()
Dim rec As RECT, rec_pos As RECT, faceNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub
    
    frmMain.picFace.Cls
    
    faceNum = GetPlayerSprite(MyIndex)
    
    If faceNum <= 0 Or faceNum > NumFaces Then Exit Sub

    With rec
        .Top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    With rec_pos
        .Top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    ' Load face if not loaded, and reset timer
    FaceTimer(faceNum) = GetTickCount + SurfaceTimerMax

    If DDS_Face(faceNum) Is Nothing Then
        Call InitDDSurf("Faces\" & faceNum, DDSD_Face(faceNum), DDS_Face(faceNum))
    End If

    Engine_BltToDC DDS_Face(faceNum), rec, rec_pos, frmMain.picFace, False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltFace", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltEquipment()
Dim i As Long, itemnum As Long, itempic As Long
Dim rec As RECT, rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumItems = 0 Then Exit Sub
    
    frmMain.picCharacter.Cls

    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(MyIndex, i)

        If itemnum > 0 Then
            itempic = Item(itemnum).Pic

            With rec
                .Top = 0
                .Bottom = 32
                .Left = 32
                .Right = 64
            End With

        ' พิกัดวาดรูปตอนใส่ item ใน Charecter ของ frmMain

            With rec_pos
                .Top = EqTop
                .Bottom = .Top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            ' Load item if not loaded, and reset timer
            ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

            If DDS_Item(itempic) Is Nothing Then
                Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
            End If

            Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picCharacter, False
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltEquipment", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltInventory()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT
Dim Colour As Long
Dim tmpItem As Long, amountModifier As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' reset gold label
    frmMain.lblGold.Caption = "0g"
    
    frmMain.picInventory.Cls

    For i = 1 To MAX_INV
        itemnum = GetPlayerInvItemNum(MyIndex, i)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            
            amountModifier = 0
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).Num)
                    If TradeYourOffer(X).Num = i Then
                        ' check if currency
                        If Not Item(tmpItem).Type = ITEM_TYPE_CURRENCY Then
                            ' normal item, exit out
                            GoTo NextLoop
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(X).Value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(X).Value
                            End If
                        End If
                    End If
                Next
            End If

            If itempic > 0 And itempic <= NumItems Then
                If DDSD_Item(itempic).lWidth <= 64 Then ' more than 1 frame is handled by anim sub

                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' Load item if not loaded, and reset timer
                    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_Item(itempic) Is Nothing Then
                        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                    End If

                    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picInventory, False

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        
                        Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            Colour = QBColor(White)
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            Colour = QBColor(Yellow)
                        ElseIf Amount > 10000000 Then
                            Colour = QBColor(BrightGreen)
                        End If
                        
                        DrawText frmMain.picInventory.hDC, X, Y, Format$(ConvertCurrency(str(Amount)), "#,###,###,###"), Colour

                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If
NextLoop:
    Next
    
    frmMain.picInventory.Refresh
    'update animated items
    BltAnimatedInvItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventory", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltTrade()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT
Dim Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    frmMain.picYourTrade.Cls
    frmMain.picTheirTrade.Cls
    
    For i = 1 To MAX_INV
        ' blt your own offer
        itemnum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= NumItems Then
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With rec_pos
                    .Top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If

                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picYourTrade, False

                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).Value > 1 Then
                    Y = rec_pos.Top + 22
                    X = rec_pos.Left - 4
                    
                    Amount = TradeYourOffer(i).Value
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        Colour = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        Colour = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        Colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picYourTrade.hDC, X, Y, ConvertCurrency(str(Amount)), Colour
                End If
            End If
        End If
            
        ' blt their offer
        itemnum = TradeTheirOffer(i).Num

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= NumItems Then
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With rec_pos
                    .Top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If

                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picTheirTrade, False

                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).Value > 1 Then
                    Y = rec_pos.Top + 22
                    X = rec_pos.Left - 4
                    
                    Amount = TradeTheirOffer(i).Value
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        Colour = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        Colour = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        Colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picTheirTrade.hDC, X, Y, ConvertCurrency(str(Amount)), Colour
                End If
            End If
        End If
    Next
    
    frmMain.picYourTrade.Refresh
    frmMain.picTheirTrade.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTrade", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltPlayerSpells()
Dim i As Long, X As Long, Y As Long, spellnum As Long, spellicon As Long
Dim Amount As String
Dim rec As RECT, rec_pos As RECT
Dim Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    frmMain.picSpells.Cls

    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i)

        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellicon = Spell(spellnum).Icon

            If spellicon > 0 And spellicon <= NumSpellIcons Then
            
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 0
                    .Right = 32
                End With
                
                If Not SpellCD(i) = 0 Then
                    rec.Left = 32
                    rec.Right = 64
                End If

                With rec_pos
                    .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load spellicon if not loaded, and reset timer
                SpellIconTimer(spellicon) = GetTickCount + SurfaceTimerMax

                If DDS_SpellIcon(spellicon) Is Nothing Then
                    Call InitDDSurf("SpellIcons\" & spellicon, DDSD_SpellIcon(spellicon), DDS_SpellIcon(spellicon))
                End If

                Engine_BltToDC DDS_SpellIcon(spellicon), rec, rec_pos, frmMain.picSpells, False
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayerSpells", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltShop()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As String
Dim rec As RECT, rec_pos As RECT
Dim Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    frmMain.picShopItems.Cls

    For i = 1 To MAX_TRADES
        itemnum = Shop(InShop).TradeItem(i).Item 'GetPlayerInvItemNum(MyIndex, i)
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            If itempic > 0 And itempic <= NumItems Then
            
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With
                
                With rec_pos
                    .Top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                    .Right = .Left + PIC_X
                End With
                
                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax
                
                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If
                
                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picShopItems, False
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    Y = rec_pos.Top + 22
                    X = rec_pos.Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = QBColor(White)
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = QBColor(Yellow)
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picShopItems.hDC, X, Y, ConvertCurrency(Amount), Colour
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltShop", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltInventoryItem(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT
Dim itemnum As Long, itempic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)

    If itemnum > 0 And itemnum <= MAX_ITEMS Then
        itempic = Item(itemnum).Pic
        
        If itempic = 0 Then Exit Sub

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = DDSD_Item(itempic).lWidth / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 2
            .Bottom = .Top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        ' Load item if not loaded, and reset timer
        ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

        If DDS_Item(itempic) Is Nothing Then
            Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
        End If

        Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picTempInv, False

        With frmMain.picTempInv
            .Top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventoryItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltDraggedSpell(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT
Dim spellnum As Long, spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = PlayerSpells(DragSpell)

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon
        
        If spellpic = 0 Then Exit Sub

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 2
            .Bottom = .Top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        ' Load item if not loaded, and reset timer
        SpellIconTimer(spellpic) = GetTickCount + SurfaceTimerMax

        If DDS_SpellIcon(spellpic) Is Nothing Then
            Call InitDDSurf("Spellicons\" & spellpic, DDSD_SpellIcon(spellpic), DDS_SpellIcon(spellpic))
        End If

        Engine_BltToDC DDS_SpellIcon(spellpic), rec, rec_pos, frmMain.picTempSpell, False

        With frmMain.picTempSpell
            .Top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventoryItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltItemDesc(ByVal itemnum As Long)
Dim rec As RECT, rec_pos As RECT
Dim itempic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.picItemDescPic.Cls
    
    If itemnum > 0 And itemnum <= MAX_ITEMS Then
        itempic = Item(itemnum).Pic

        If itempic = 0 Then Exit Sub
        
        ' Load item if not loaded, and reset timer
        ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

        If DDS_Item(itempic) Is Nothing Then
            Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
        End If

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = DDSD_Item(itempic).lWidth / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picItemDescPic, False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltItemDesc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltSpellDesc(ByVal spellnum As Long)
Dim rec As RECT, rec_pos As RECT
Dim spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.picSpellDescPic.Cls

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon

        If spellpic <= 0 Or spellpic > NumSpellIcons Then Exit Sub
        
        ' Load item if not loaded, and reset timer
        SpellIconTimer(spellpic) = GetTickCount + SurfaceTimerMax

        If DDS_SpellIcon(spellpic) Is Nothing Then
            Call InitDDSurf("SpellIcons\" & spellpic, DDSD_SpellIcon(spellpic), DDS_SpellIcon(spellpic))
        End If

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        Engine_BltToDC DDS_SpellIcon(spellpic), rec, rec_pos, frmMain.picSpellDescPic, False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSpellDesc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_BltTileset()
Dim Height As Long
Dim Width As Long
Dim Tileset As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    ' make sure it's loaded
    If DDS_Tileset(Tileset) Is Nothing Then
        Call InitDDSurf("tilesets\" & Tileset, DDSD_Tileset(Tileset), DDS_Tileset(Tileset))
    End If
    
    Height = DDSD_Tileset(Tileset).lHeight
    Width = DDSD_Tileset(Tileset).lWidth
    
    dRECT.Top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width
    
    frmEditor_Map.picBackSelect.Height = Height
    frmEditor_Map.picBackSelect.Width = Width
    
    Call Engine_BltToDC(DDS_Tileset(Tileset), sRECT, dRECT, frmEditor_Map.picBackSelect)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltTileset", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTileOutline()
Dim rec As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optBlock.Value Then Exit Sub

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    Call Engine_BltFast(ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), DDS_Misc, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTileOutline", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NewCharacterBltSprite()
Dim Sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If frmMenu.optMale.Value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("Characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If
    
    Width = DDSD_Character(Sprite).lWidth / 4
    Height = DDSD_Character(Sprite).lHeight / 4
    
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    sRECT.Top = 0
    sRECT.Bottom = sRECT.Top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    
    dRECT.Top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width
    
    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmMenu.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NewCharacterBltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_BltMapItem()
Dim itemnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = Item(frmEditor_Map.scrlMapItem.Value).Pic

    If itemnum < 1 Or itemnum > NumItems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    ItemTimer(itemnum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(itemnum) Is Nothing Then
        Call InitDDSurf("Items\" & itemnum, DDSD_Item(itemnum), DDS_Item(itemnum))
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(itemnum), sRECT, dRECT, frmEditor_Map.picMapItem)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltMapItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_BltKey()
Dim itemnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = Item(frmEditor_Map.scrlMapKey.Value).Pic

    If itemnum < 1 Or itemnum > NumItems Then
        frmEditor_Map.picMapKey.Cls
        Exit Sub
    End If

    ItemTimer(itemnum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(itemnum) Is Nothing Then
        Call InitDDSurf("Items\" & itemnum, DDSD_Item(itemnum), DDS_Item(itemnum))
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(itemnum), sRECT, dRECT, frmEditor_Map.picMapKey)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltKey", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltItem()
Dim itemnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = frmEditor_Item.scrlPic.Value

    If itemnum < 1 Or itemnum > NumItems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    ItemTimer(itemnum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(itemnum) Is Nothing Then
        Call InitDDSurf("Items\" & itemnum, DDSD_Item(itemnum), DDS_Item(itemnum))
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRECT = sRECT
    Call Engine_BltToDC(DDS_Item(itemnum), sRECT, dRECT, frmEditor_Item.picItem)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltPaperdoll()
Dim Sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Item.picPaperdoll.Cls
    
    Sprite = frmEditor_Item.scrlPaperdoll.Value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    PaperdollTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Paperdoll(Sprite) Is Nothing Then
        Call InitDDSurf("paperdolls\" & Sprite, DDSD_Paperdoll(Sprite), DDS_Paperdoll(Sprite))
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = DDSD_Paperdoll(Sprite).lHeight
    sRECT.Left = 0
    sRECT.Right = DDSD_Paperdoll(Sprite).lWidth
    ' same for destination as source
    dRECT = sRECT
    
    Call Engine_BltToDC(DDS_Paperdoll(Sprite), sRECT, dRECT, frmEditor_Item.picPaperdoll)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltPaperdoll", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_BltIcon()
Dim iconnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    iconnum = frmEditor_Spell.scrlIcon.Value
    
    If iconnum < 1 Or iconnum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    SpellIconTimer(iconnum) = GetTickCount + SurfaceTimerMax
    
    If DDS_SpellIcon(iconnum) Is Nothing Then
        Call InitDDSurf("SpellIcons\" & iconnum, DDSD_SpellIcon(iconnum), DDS_SpellIcon(iconnum))
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    
    Call Engine_BltToDC(DDS_SpellIcon(iconnum), sRECT, dRECT, frmEditor_Spell.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_BltIcon", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_BltAnim()
Dim Animationnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
Dim i As Long
Dim Width As Long, Height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim ShouldRender As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).Value
        
        If Animationnum < 1 Or Animationnum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                frmEditor_Animation.picSprite(i).Cls
            
                AnimationTimer(Animationnum) = GetTickCount + SurfaceTimerMax
                
                If DDS_Animation(Animationnum) Is Nothing Then
                    Call InitDDSurf("animations\" & Animationnum, DDSD_Animation(Animationnum), DDS_Animation(Animationnum))
                End If
                
                If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then
                    ' total width divided by frame count
                    Width = DDSD_Animation(Animationnum).lWidth / frmEditor_Animation.scrlFrameCount(i).Value
                    Height = DDSD_Animation(Animationnum).lHeight
                    
                    sRECT.Top = 0
                    sRECT.Bottom = Height
                    sRECT.Left = (AnimEditorFrame(i) - 1) * Width
                    sRECT.Right = sRECT.Left + Width
                    
                    dRECT.Top = 0
                    dRECT.Bottom = Height
                    dRECT.Left = 0
                    dRECT.Right = Width
                    
                    Call Engine_BltToDC(DDS_Animation(Animationnum), sRECT, dRECT, frmEditor_Animation.picSprite(i))
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_BltAnim", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_BltSprite()
Dim Sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
Dim X As Long, Y As Long, Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_NPC.scrlSprite.Value

    If Sprite < 1 Or Sprite > NumCharacters Then
        ' Clear the screen so it doesn't leave lingering images.
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    ' Calculate the locations and Render the graphic
    X = (frmEditor_NPC.picSprite.ScaleWidth / 2) - (DDSD_Character(Sprite).lWidth / 4) / 2
    Y = (frmEditor_NPC.picSprite.ScaleHeight / 2) - (DDSD_Character(Sprite).lHeight / 4) / 2
    Width = DDSD_Character(Sprite).lWidth / 4
    Height = DDSD_Character(Sprite).lHeight / 4

    sRECT.Top = 0
    sRECT.Bottom = sRECT.Top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    
    dRECT.Top = X
    dRECT.Bottom = Height + dRECT.Top
    dRECT.Left = Y
    dRECT.Right = Width + dRECT.Left
    
    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmEditor_NPC.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_BltSprite()
Dim Sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        ResourceTimer(Sprite) = GetTickCount + SurfaceTimerMax
        If DDS_Resource(Sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & Sprite, DDSD_Resource(Sprite), DDS_Resource(Sprite))
        End If
        sRECT.Top = 0
        sRECT.Bottom = DDSD_Resource(Sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(Sprite).lWidth
        dRECT.Top = 0
        dRECT.Bottom = DDSD_Resource(Sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(Sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(Sprite), sRECT, dRECT, frmEditor_Resource.picNormalPic)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        ResourceTimer(Sprite) = GetTickCount + SurfaceTimerMax
        If DDS_Resource(Sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & Sprite, DDSD_Resource(Sprite), DDS_Resource(Sprite))
        End If
        sRECT.Top = 0
        sRECT.Bottom = DDSD_Resource(Sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(Sprite).lWidth
        dRECT.Top = 0
        dRECT.Bottom = DDSD_Resource(Sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(Sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(Sprite), sRECT, dRECT, frmEditor_Resource.picExhaustedPic)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Render_Graphics()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim n As Long
Dim rec As DxVBLib.RECT
Dim rec_pos As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check if automation is screwed
    If Not CheckSurfaces Then
        ' exit out and let them know we need to re-init
        ReInitSurfaces = True
        Exit Sub
    Else
        ' if we need to fix the surfaces then do so
        If ReInitSurfaces Then
            ReInitSurfaces = False
            ReInitDD
        End If
    End If
    
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub
    
    ' update the viewpoint
    UpdateCamera
    
    ' update animation editor
    If Editor = EDITOR_ANIMATION Then
        EditorAnim_BltAnim
    End If
    
    ' fill it with black
    DDS_BackBuffer.BltColorFill rec_pos, 0
    
    ' blit lower tiles
    If NumTileSets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapTile(X, Y)
                End If
            Next
        Next
    End If

    ' render the decals
    For i = 1 To MAX_BYTE
        Call BltBlood(i)
    Next

    ' Blit out the items
    If NumItems > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).Num > 0 Then
                Call BltItem(i)
            End If
        Next
    End If
    
    If Map.CurrentEvents > 0 Then
        For i = 1 To Map.CurrentEvents
            If Map.MapEvents(i).Position = 0 Then
                BltEvent i
            End If
        Next
    End If
    
    ' draw animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                BltAnimation i, 0
            End If
        Next
    End If
    
    ' projec tile
    ' blt projec tiles for each player
    For i = 1 To Player_HighIndex
        For n = 1 To MAX_PLAYER_PROJECTILES
            If Player(i).ProjecTile(X).Pic > 0 Then
                BltProjectile i, X
            End If
        Next
    Next

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For Y = 0 To Map.maxY
        
        If NumCharacters > 0 Then
        
            ' Event
            If Map.CurrentEvents > 0 Then
                For i = 1 To Map.CurrentEvents
                    If Map.MapEvents(i).Position = 1 Then
                        If Y = Map.MapEvents(i).Y Then
                            BltEvent i
                        End If
                    End If
                Next
            End If
            
            ' Players
            For i = 1 To Player_HighIndex
                
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Y = Player(i).Y Then
                        If Not i = MyIndex Then
                            Call BltPlayer(i)
                        End If
                    End If
                End If
                
                ' Render our sprite now so it's always at the top
                If Player(MyIndex).Y = Y Then
                    Call BltPlayer(i)
                End If
                
                ' Npcs
                For n = 1 To Npc_HighIndex
                    If IsPlaying(i) And GetPlayerMap(i) = MapNpc(n).Num Then
                        If Y = MapNpc(n).Y Then
                            Call BltNpc(n)
                        End If
                    End If
                Next
                            
            Next
            
        End If
        
        ' Resources
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = Y Then
                            Call BltMapResource(i)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                BltAnimation i, 1
            End If
        Next
    End If

    ' blit out upper tiles
    If NumTileSets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapFringeTile(X, Y)
                End If
            Next
        Next
    End If
    
    If Map.CurrentEvents > 0 Then
        
        For i = 1 To Map.CurrentEvents
            If Map.MapEvents(i).Position = 2 Then
                BltEvent i
            End If
        Next
    End If
    
    ' วาดบาร์
    BltBars
    
    ' สภาพอากาศ
    BltWeather
    
    ' blit out a square at mouse cursor
    If InMapEditor Then
        If frmEditor_Map.optBlock.Value = True Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.Top To TileView.Bottom
                    If IsValidMapPoint(X, Y) Then
                        Call BltDirection(X, Y)
                    End If
                Next
            Next
        End If
        Call BltTileOutline
    End If
    
    ' minimap
    If Options.Minimap = 1 Then BltMiniMap
    
    ' Blt the target icon
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
            BltTarget (Player(myTarget).X * 32) + Player(myTarget).xOffset, (Player(myTarget).Y * 32) + Player(myTarget).yOffset
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            BltTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).xOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).yOffset
        End If
    End If
    
    ' blt the hover icon
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Map = Player(MyIndex).Map Then
                If CurX = Player(i).X And CurY = Player(i).Y Then
                    If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                        ' dont render lol
                    Else
                        BltHover TARGET_TYPE_PLAYER, i, (Player(i).X * 32) + Player(i).xOffset, (Player(i).Y * 32) + Player(i).yOffset
                    End If
                End If
            End If
        End If
    Next
    
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            If CurX = MapNpc(i).X And CurY = MapNpc(i).Y Then
                If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                    ' dont render lol
                Else
                    BltHover TARGET_TYPE_NPC, i, (MapNpc(i).X * 32) + MapNpc(i).xOffset, (MapNpc(i).Y * 32) + MapNpc(i).yOffset
                End If
            End If
        End If
    Next
    
    If frmEditor_Events.Visible Then
        EditorEvent_BltGraphic
    End If
If InMapEditor Then
        If frmEditor_Map.optEvent.Value = True Then
            BltEvents
        End If
    End If

    ' Lock the backbuffer so we can draw text and names
    TexthDC = DDS_BackBuffer.GetDC

    ' draw FPS
    If BFPS Then
        Call DrawText(TexthDC, Camera.Right - (Len("FPS : " & GameFPS) * 8), Camera.Top + 1, Trim$("FPS : " & GameFPS), QBColor(Yellow))
    End If
    
    '
    If DMAP Then
        Call DrawText(TexthDC, Camera.Left + 1, Camera.Top + 1, Trim$("กดปุ่ม m เพื่อเปิด/ปิด แผนที่"), QBColor(BrightGreen))
    End If

    ' draw cursor, player X and Y locations
    If BLoc Then
        Call DrawText(TexthDC, Camera.Left, Camera.Top + 1, Trim$("เมาส์ x : " & CurX & " y : " & CurY), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.Top + 15, Trim$("พิกัดยืน x : " & GetPlayerX(MyIndex) & " y : " & GetPlayerY(MyIndex)), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.Top + 27, Trim$(" (แผนที่ #" & GetPlayerMap(MyIndex) & ")"), QBColor(Yellow))
    End If
    
    For i = 1 To Map.CurrentEvents
        If Map.MapEvents(i).Visible = 1 Then
            If Map.MapEvents(i).ShowName = 1 Then
                DrawEventName (i)
            End If
        End If
    Next

    ' draw player names
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            Call DrawPlayerName(i)
        End If
    Next
    
    ' draw npc names
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            Call DrawNpcName(i)
        End If
    Next
    
    For i = 1 To Action_HighIndex
        Call BltActionMsg(i)
    Next i

    ' Blit out map attributes
    If InMapEditor Then
        Call BltMapAttributes
    End If

    ' Release DC
    DDS_BackBuffer.ReleaseDC TexthDC
    
    ' Get rec
    With rec
        .Top = Camera.Top
        .Bottom = .Top + ScreenY
        .Left = Camera.Left
        .Right = .Left + ScreenX
    End With
    
    ' rec_pos
    With rec_pos
        .Bottom = ((MAX_MAPY + 1) * PIC_Y)
        .Right = ((MAX_MAPX + 1) * PIC_X)
    End With
    
    ' Flip and render
    DX7.GetWindowRect frmMain.picScreen.hwnd, rec_pos
    DDS_Primary.Blt rec_pos, DDS_BackBuffer, rec, DDBLT_WAIT
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "Render_Graphics", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateCamera()
Dim offsetX As Long
Dim offsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    offsetX = Player(MyIndex).xOffset + PIC_X
    offsetY = Player(MyIndex).yOffset + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        offsetX = 0
        If StartX = -1 Then
            If Player(MyIndex).xOffset > 0 Then
                offsetX = Player(MyIndex).xOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0
        If StartY = -1 Then
            If Player(MyIndex).yOffset > 0 Then
                offsetY = Player(MyIndex).yOffset
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.maxX Then
        offsetX = 32
        If EndX = Map.maxX + 1 Then
            If Player(MyIndex).xOffset < 0 Then
                offsetX = Player(MyIndex).xOffset + PIC_X
            End If
        End If
        EndX = Map.maxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.maxY Then
        offsetY = 32
        If EndY = Map.maxY + 1 Then
            If Player(MyIndex).yOffset < 0 Then
                offsetY = Player(MyIndex).yOffset + PIC_Y
            End If
        End If
        EndY = Map.maxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .Bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
    
    UpdateDrawMapName

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = X - (TileView.Left * PIC_X)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = Y - (TileView.Top * PIC_Y)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.Top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    InViewPort = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.maxX Then Exit Function
    If Y > Map.maxY Then Exit Function
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTilesets()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim tilesetInUse() As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim tilesetInUse(0 To NumTileSets)
    
    For X = 0 To Map.maxX
        For Y = 0 To Map.maxY
            For i = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If Map.Tile(X, Y).Layer(i).Tileset > 0 And Map.Tile(X, Y).Layer(i).Tileset <= NumTileSets Then
                    tilesetInUse(Map.Tile(X, Y).Layer(i).Tileset) = True
                End If
            Next
        Next
    Next
    
    For i = 1 To NumTileSets
        If tilesetInUse(i) Then
            ' load tileset
            If DDS_Tileset(i) Is Nothing Then
                Call InitDDSurf("tilesets\" & i, DDSD_Tileset(i), DDS_Tileset(i))
            End If
        Else
            ' unload tileset
            Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            Set DDS_Tileset(i) = Nothing
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltBank()
Dim i As Long, X As Long, Y As Long, itemnum As Long
Dim Amount As String
Dim sRECT As RECT, dRECT As RECT
Dim Sprite As Long, Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMain.picBank.Visible = True Then
        frmMain.picBank.Cls
                
        For i = 1 To MAX_BANK
            itemnum = GetBankItemNum(i)
            If itemnum > 0 And itemnum <= MAX_ITEMS Then
            
                Sprite = Item(itemnum).Pic
                
                If Sprite <= 0 Or Sprite > NumItems Then Exit Sub
                
                If DDS_Item(Sprite) Is Nothing Then
                    Call InitDDSurf("Items\" & Sprite, DDSD_Item(Sprite), DDS_Item(Sprite))
                End If
            
                With sRECT
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = DDSD_Item(Sprite).lWidth / 2
                    .Right = .Left + PIC_X
                End With
                
                With dRECT
                    .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                    .Right = .Left + PIC_X
                End With
                
                Engine_BltToDC DDS_Item(Sprite), sRECT, dRECT, frmMain.picBank, False

                ' If item is a stack - draw the amount you have
                If GetBankItemValue(0, i) > 1 Then
                    Y = dRECT.Top + 22
                    X = dRECT.Left - 4
                
                    Amount = CStr(GetBankItemValue(0, i))
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = QBColor(White)
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = QBColor(Yellow)
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = QBColor(BrightGreen)
                    End If
                    DrawText frmMain.picBank.hDC, X, Y, ConvertCurrency(Amount), Colour
                End If
            End If
        Next
    
        frmMain.picBank.Refresh
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBank", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltBankItem(ByVal X As Long, ByVal Y As Long)
Dim sRECT As RECT, dRECT As RECT
Dim itemnum As Long
Dim Sprite As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = GetBankItemNum(DragBankSlotNum)
    Sprite = Item(GetBankItemNum(DragBankSlotNum)).Pic
    
    If DDS_Item(Sprite) Is Nothing Then
        Call InitDDSurf("Items\" & Sprite, DDSD_Item(Sprite), DDS_Item(Sprite))
    End If
    
    If itemnum > 0 Then
        If itemnum <= MAX_ITEMS Then
            With sRECT
                .Top = 0
                .Bottom = .Top + PIC_Y
                .Left = DDSD_Item(Sprite).lWidth / 2
                .Right = .Left + PIC_X
            End With
        End If
    End If
    
    With dRECT
        .Top = 2
        .Bottom = .Top + PIC_Y
        .Left = 2
        .Right = .Left + PIC_X
    End With

    Engine_BltToDC DDS_Item(Sprite), sRECT, dRECT, frmMain.picTempBank
    
    With frmMain.picTempBank
        .Top = Y
        .Left = X
        .Visible = True
        .ZOrder (0)
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBankItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltEvents()
Dim sRECT As DxVBLib.RECT
Dim Width As Long, Height As Long, i As Long, X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Map.EventCount <= 0 Then Exit Sub
    
    For i = 1 To Map.EventCount
        If Map.Events(i).pageCount <= 0 Then
                sRECT.Top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            GoTo nextevent
        End If
        
        Width = 32
        Height = 32
    
        X = Map.Events(i).X * 32
        Y = Map.Events(i).Y * 32
        X = ConvertMapX(X)
        Y = ConvertMapY(Y)
    
        ' clipping
        If Y < 0 Then
            With sRECT
                .Top = .Top - Y
            End With
            Y = 0
        End If
    
        If X < 0 Then
            With sRECT
                .Left = .Left - X
            End With
            X = 0
        End If
    
        If Y + Height > DDSD_BackBuffer.lHeight Then
            sRECT.Bottom = sRECT.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
        End If
    
        If X + Width > DDSD_BackBuffer.lWidth Then
            sRECT.Right = sRECT.Right - (X + Width - DDSD_BackBuffer.lWidth)
        End If
        
        If i > Map.EventCount Then Exit Sub
        If 1 > Map.Events(i).pageCount Then Exit Sub
    ' /clipping
        Select Case Map.Events(i).Pages(1).GraphicType
            Case 0
                sRECT.Top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Case 1
                If Map.Events(i).Pages(1).Graphic > 0 And Map.Events(i).Pages(1).Graphic <= NumCharacters Then
                    CharacterTimer(Map.Events(i).Pages(1).Graphic) = GetTickCount + SurfaceTimerMax
                    If DDS_Character(Map.Events(i).Pages(1).Graphic) Is Nothing Then
                        Call InitDDSurf("Characters\" & Map.Events(i).Pages(1).Graphic, DDSD_Character(Map.Events(i).Pages(1).Graphic), DDS_Character(Map.Events(i).Pages(1).Graphic))
                    End If
                    
                    sRECT.Top = (Map.Events(i).Pages(1).GraphicY * (DDSD_Character(Map.Events(i).Pages(1).Graphic).lHeight / 4))
                    sRECT.Left = (Map.Events(i).Pages(1).GraphicX * (DDSD_Character(Map.Events(i).Pages(1).Graphic).lWidth / 4))
                    sRECT.Bottom = sRECT.Top + 32
                    sRECT.Right = sRECT.Left + 32
                    
                    Call Engine_BltFast(X, Y, DDS_Character(Map.Events(i).Pages(1).Graphic), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_SRCCOLORKEY)
                Else
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_SRCCOLORKEY)
                End If
            Case 2
                If Map.Events(i).Pages(1).Graphic > 0 And Map.Events(i).Pages(1).Graphic < NumTileSets Then
                    sRECT.Top = Map.Events(i).Pages(1).GraphicY * 32
                    sRECT.Left = Map.Events(i).Pages(1).GraphicX * 32
                    sRECT.Bottom = sRECT.Top + 32
                    sRECT.Right = sRECT.Left + 32
                    Call Engine_BltFast(X, Y, DDS_Tileset(Map.Events(i).Pages(1).Graphic), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_SRCCOLORKEY)
                End If
        End Select
nextevent:
    Next
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHover", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorEvent_BltGraphic()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Events.picGraphicSel.Visible Then
        Select Case frmEditor_Events.cmbGraphic.ListIndex
            Case 0
                'None
                frmEditor_Events.picGraphicSel.Cls
                Exit Sub
            Case 1
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumCharacters Then
                    CharacterTimer(frmEditor_Events.scrlGraphic.Value) = GetTickCount + SurfaceTimerMax
                    If DDS_Character(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
                        Call InitDDSurf("Characters\" & frmEditor_Events.scrlGraphic.Value, DDSD_Character(frmEditor_Events.scrlGraphic.Value), DDS_Character(frmEditor_Events.scrlGraphic.Value))
                    End If
                    
                    If DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth > 800 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Right = sRECT.Left + (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth - sRECT.Left)
                    Else
                        sRECT.Left = 0
                        sRECT.Right = DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth
                    End If
                    
                    If DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight > 512 Then
                        sRECT.Top = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Bottom = sRECT.Top + (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight - sRECT.Top)
                    Else
                        sRECT.Top = 0
                        sRECT.Bottom = DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight
                    End If
                    
                    With dRECT
                        .Top = 0
                        .Bottom = sRECT.Bottom - sRECT.Top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    Call Engine_BltToDC(DDS_Character(frmEditor_Events.scrlGraphic.Value), sRECT, dRECT, frmEditor_Events.picGraphicSel)
                    
                    If GraphicSelX <= 3 And GraphicSelY <= 3 Then
                        frmEditor_Events.shpLoc.Left = GraphicSelX * (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth / 4)
                        frmEditor_Events.shpLoc.Width = (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth / 4)
                        frmEditor_Events.shpLoc.Top = GraphicSelY * (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight / 4)
                        frmEditor_Events.shpLoc.Height = (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight / 4)
                    Else
                        GraphicSelX = 0
                        GraphicSelY = 0
                    End If
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
            Case 2
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumTileSets Then
                    If DDS_Tileset(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
                        Call InitDDSurf("Tilesets\" & frmEditor_Events.scrlGraphic.Value, DDSD_Tileset(frmEditor_Events.scrlGraphic.Value), DDS_Tileset(frmEditor_Events.scrlGraphic.Value))
                    End If
                    
                    If DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lWidth > 800 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Right = sRECT.Left + 800
                    Else
                        sRECT.Left = 0
                        sRECT.Right = DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lWidth
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value = 0
                    End If
                    
                    If DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lHeight > 512 Then
                        sRECT.Top = frmEditor_Events.vScrlGraphicSel.Value
                        sRECT.Bottom = sRECT.Top + 512
                    Else
                        sRECT.Top = 0
                        sRECT.Bottom = DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lHeight
                        frmEditor_Events.vScrlGraphicSel.Value = 0
                    End If
                    
                    With dRECT
                        .Top = 0
                        .Bottom = sRECT.Bottom - sRECT.Top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    Call Engine_BltToDC(DDS_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, dRECT, frmEditor_Events.picGraphicSel)
         
                    'Now we draw the selection square.. tad bit harder....
                    'Stretched or not....
                    If GraphicSelX2 > 0 Or GraphicSelY2 > 0 Then
                        frmEditor_Events.shpLoc.Top = (GraphicSelY * 32) - frmEditor_Events.vScrlGraphicSel.Value
                        frmEditor_Events.shpLoc.Left = (GraphicSelX * 32) - frmEditor_Events.hScrlGraphicSel.Value
                        frmEditor_Events.shpLoc.Width = (GraphicSelX2 - GraphicSelX) * 32
                        frmEditor_Events.shpLoc.Height = (GraphicSelY2 - GraphicSelY) * 32
                    Else
                        frmEditor_Events.shpLoc.Top = (GraphicSelY * 32) - frmEditor_Events.vScrlGraphicSel.Value
                        frmEditor_Events.shpLoc.Left = (GraphicSelX * 32) - frmEditor_Events.hScrlGraphicSel.Value
                        frmEditor_Events.shpLoc.Width = 32
                        frmEditor_Events.shpLoc.Height = 32
                    End If
                    
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
        End Select
    Else
        Select Case tmpEvent.Pages(curPageNum).GraphicType
            Case 0
                frmEditor_Events.picGraphic.Cls
            Case 1
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumCharacters Then
                    sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * (DDSD_Character(tmpEvent.Pages(curPageNum).Graphic).lHeight / 4)
                    sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * (DDSD_Character(tmpEvent.Pages(curPageNum).Graphic).lWidth / 4)
                    sRECT.Bottom = sRECT.Top + (DDSD_Character(tmpEvent.Pages(curPageNum).Graphic).lHeight / 4)
                    sRECT.Right = sRECT.Left + (DDSD_Character(tmpEvent.Pages(curPageNum).Graphic).lWidth / 4)
                    With dRECT
                        dRECT.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                        dRECT.Bottom = dRECT.Top + (sRECT.Bottom - sRECT.Top)
                        dRECT.Left = (121 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                        dRECT.Right = dRECT.Left + (sRECT.Right - sRECT.Left)
                    End With
                    If DDS_Character(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
                        Call InitDDSurf("Characters\" & frmEditor_Events.scrlGraphic.Value, DDSD_Character(frmEditor_Events.scrlGraphic.Value), DDS_Character(frmEditor_Events.scrlGraphic.Value))
                    End If
                    Call Engine_BltToDC(DDS_Character(frmEditor_Events.scrlGraphic.Value), sRECT, dRECT, frmEditor_Events.picGraphic)
                End If
            Case 2
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumTileSets Then
                    If tmpEvent.Pages(curPageNum).GraphicX2 = 0 Or tmpEvent.Pages(curPageNum).GraphicY2 = 0 Then
                        sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.Top + 32
                        sRECT.Right = sRECT.Left + 32
                        With dRECT
                            dRECT.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                            dRECT.Bottom = dRECT.Top + (sRECT.Bottom - sRECT.Top)
                            dRECT.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            dRECT.Right = dRECT.Left + (sRECT.Right - sRECT.Left)
                        End With
                        If DDS_Tileset(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
                            Call InitDDSurf("Tilesets\" & frmEditor_Events.scrlGraphic.Value, DDSD_Tileset(frmEditor_Events.scrlGraphic.Value), DDS_Tileset(frmEditor_Events.scrlGraphic.Value))
                        End If
                        Call Engine_BltToDC(DDS_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, dRECT, frmEditor_Events.picGraphic)

                    Else
                        sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.Top + ((tmpEvent.Pages(curPageNum).GraphicY2 - tmpEvent.Pages(curPageNum).GraphicY) * 32)
                        sRECT.Right = sRECT.Left + ((tmpEvent.Pages(curPageNum).GraphicX2 - tmpEvent.Pages(curPageNum).GraphicX) * 32)
                        With dRECT
                            dRECT.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                            dRECT.Bottom = dRECT.Top + (sRECT.Bottom - sRECT.Top)
                            dRECT.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            dRECT.Right = dRECT.Left + (sRECT.Right - sRECT.Left)
                        End With
                        If DDS_Tileset(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
                            Call InitDDSurf("Tilesets\" & frmEditor_Events.scrlGraphic.Value, DDSD_Tileset(frmEditor_Events.scrlGraphic.Value), DDS_Tileset(frmEditor_Events.scrlGraphic.Value))
                        End If
                        Call Engine_BltToDC(DDS_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, dRECT, frmEditor_Events.picGraphic)
                    End If
                End If
        End Select
    End If
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltKey", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltEvent(id As Long)
    Dim X As Long, Y As Long, Width As Long, Height As Long, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, anim As Long, spritetop As Long
    If Map.MapEvents(id).Visible = 0 Then Exit Sub
    If InMapEditor Then Exit Sub
    Select Case Map.MapEvents(id).GraphicType
        Case 0
            Exit Sub
            
        Case 1
            If Map.MapEvents(id).GraphicNum <= 0 Or Map.MapEvents(id).GraphicNum > NumCharacters Then Exit Sub
            CharacterTimer(Map.MapEvents(id).GraphicNum) = GetTickCount + SurfaceTimerMax
            If DDS_Character(Map.MapEvents(id).GraphicNum) Is Nothing Then
                Call InitDDSurf("characters\" & Map.MapEvents(id).GraphicNum, DDSD_Character(Map.MapEvents(id).GraphicNum), DDS_Character(Map.MapEvents(id).GraphicNum))
            End If
            Width = DDSD_Character(Map.MapEvents(id).GraphicNum).lWidth / 4
            Height = DDSD_Character(Map.MapEvents(id).GraphicNum).lHeight / 4
            ' Reset frame
            If Map.MapEvents(id).Step = 3 Then
                anim = 0
            ElseIf Map.MapEvents(id).Step = 1 Then
                anim = 2
            End If
            
            Select Case Map.MapEvents(id).Dir
                Case DIR_UP
                    If (Map.MapEvents(id).yOffset > 8) Then anim = Map.MapEvents(id).Step
                Case DIR_DOWN
                    If (Map.MapEvents(id).yOffset < -8) Then anim = Map.MapEvents(id).Step
                Case DIR_LEFT
                    If (Map.MapEvents(id).xOffset > 8) Then anim = Map.MapEvents(id).Step
                Case DIR_RIGHT
                    If (Map.MapEvents(id).xOffset < -8) Then anim = Map.MapEvents(id).Step
            End Select
            
            ' Set the left
            Select Case Map.MapEvents(id).ShowDir
                Case DIR_UP
                    spritetop = 3
                Case DIR_RIGHT
                    spritetop = 2
                Case DIR_DOWN
                    spritetop = 0
                Case DIR_LEFT
                    spritetop = 1
            End Select
            
            If Map.MapEvents(id).WalkAnim = 1 Then anim = 0
            
            If Map.MapEvents(id).Moving = 0 Then anim = Map.MapEvents(id).GraphicX
            
            With sRECT
                .Top = spritetop * Height
                .Bottom = .Top + Height
                .Left = anim * Width
                .Right = .Left + Width
            End With
        
            ' Calculate the X
            X = Map.MapEvents(id).X * PIC_X + Map.MapEvents(id).xOffset - ((Width - 32) / 2)
        
            ' Is the player's height more than 32..?
            If (Height * 4) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                Y = Map.MapEvents(id).Y * PIC_Y + Map.MapEvents(id).yOffset - ((Height) - 32)
            Else
                ' Proceed as normal
                Y = Map.MapEvents(id).Y * PIC_Y + Map.MapEvents(id).yOffset
            End If
        
            ' render the actual sprite
            Call BltSprite(Map.MapEvents(id).GraphicNum, X, Y, sRECT)
            
        Case 2
            If Map.MapEvents(id).GraphicNum < 1 Or Map.MapEvents(id).GraphicNum > NumTileSets Then Exit Sub
            
            If Map.MapEvents(id).GraphicY2 > 0 Or Map.MapEvents(id).GraphicX2 > 0 Then
                With sRECT
                    .Top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .Top + ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) * 32)
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + ((Map.MapEvents(id).GraphicX2 - Map.MapEvents(id).GraphicX) * 32)
                End With
            Else
                With sRECT
                    .Top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .Top + 32
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + 32
                End With
            End If
            
            X = Map.MapEvents(id).X * 32
            Y = Map.MapEvents(id).Y * 32
            
            X = X - ((sRECT.Right - sRECT.Left) / 2)
            Y = Y - (sRECT.Bottom - sRECT.Top) + 32
            
            
            If DDS_Tileset(Map.MapEvents(id).GraphicNum) Is Nothing Then
                Call InitDDSurf("tilesets\" & Map.MapEvents(id).GraphicNum, DDSD_Tileset(Map.MapEvents(id).GraphicNum), DDS_Tileset(Map.MapEvents(id).GraphicNum))
            End If
            If Map.MapEvents(id).GraphicY2 > 0 Then
                Call Engine_BltFast(ConvertMapX(Map.MapEvents(id).X * 32), ConvertMapY((Map.MapEvents(id).Y - ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) - 1)) * 32), DDS_Tileset(Map.MapEvents(id).GraphicNum), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                Call Engine_BltFast(ConvertMapX(Map.MapEvents(id).X * 32), ConvertMapY(Map.MapEvents(id).Y * 32), DDS_Tileset(Map.MapEvents(id).GraphicNum), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
    End Select
End Sub

Sub BltWeather()
    Dim i As Long, sRECT As RECT

    ' rain
    If Map.Weather = WEATHER_RAINING Then
        ' Call DDS_BackBuffer.SetForeColor(RGB(12, 40, 96))
        ' สีของฝนที่ตกลงมา
        Call DDS_BackBuffer.SetForeColor(RGB(82, 139, 185))
        For i = 1 To MAX_RAINDROPS
            With DropRain(i)
                If .Init = True Then
                    ' move o rain
                    .Y = .Y + .ySpeed
                    ' checar a screen
                    If .Y > 480 + 64 Then
                        .Y = Rand(0, 100)
                        .Y = .Y - 100
                        .X = Rand(0, 640 + 64)
                        .ySpeed = Rand(5, 10)
                        .Init = True
                    End If
                    ' draw rain
                    DDS_BackBuffer.DrawLine .X + Camera.Left, .Y + Camera.Top, .X + Camera.Left, .Y + (.ySpeed * 2) + Camera.Top
                Else
                    .Y = Rand(0, 100)
                    .Y = .Y - 100
                    .X = Rand(0, 640 + 64)
                    .ySpeed = Rand(5, 10)
                    .Init = True
                End If
            End With
        Next
    End If
    
    ' snow
    If Map.Weather = WEATHER_SNOWING Then
        Call DDS_BackBuffer.SetForeColor(RGB(255, 255, 255))
        For i = 1 To MAX_SNOWDROPS
            With DropSnow(i)
                If .Init = True Then
                    ' move o snow
                    .Y = .Y + .ySpeed
                    .X = .X + .xSpeed
                    ' checar screen
                    If .Y > 480 + 64 Then
                        .Y = Rand(0, 100) - 100
                        .X = Rand(0, 640 + 64)
                        .ySpeed = Rand(1, 4)
                        .xSpeed = Rand(0, 4) - 2
                    End If
                    ' draw rain
                    With sRECT
                        .Top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With
                    Engine_BltFast .X + Camera.Left, .Y + Camera.Top, DDS_Snow, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    .Y = Rand(0, 480)
                    .X = Rand(0, 640 + 64)
                    .ySpeed = Rand(1, 4)
                    .xSpeed = Rand(0, 4) - 2
                    .Init = True
                End If
            End With
        Next
    End If
    
    ' bird
    If Map.Weather = WEATHER_BIRD Then
        'Call DDS_BackBuffer.SetForeColor(RGB(255, 255, 255))
        For i = 1 To MAX_BIRDDROPS
            With DropBird(i)
                If .Init = True Then
                    ' move o snow
                    .Y = .Y + .ySpeed
                    .X = .X + .xSpeed
                    ' checar a screen
                    If .Y > 480 + 64 Then
                        .Y = Rand(0, 100) - 100
                        .X = Rand(0, 640 + 64)
                        .ySpeed = Rand(1, 4)
                        .xSpeed = Rand(0, 4) - 2
                    End If
                    ' draw rain
                    With sRECT
                        .Top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With
                    Engine_BltFast .X + Camera.Left, .Y + Camera.Top, DDS_Bird, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    .Y = Rand(0, 480)
                    .X = Rand(0, 640 + 64)
                    .ySpeed = Rand(1, 4)
                    .xSpeed = Rand(0, 4) - 2
                    .Init = True
                End If
            End With
        Next
    End If
    
    ' sand
    If Map.Weather = WEATHER_SAND Then 'neve
        'Call DDS_BackBuffer.SetForeColor(RGB(255, 255, 255))
        For i = 1 To MAX_SANDDROPS
            With DropSand(i)
                If .Init = True Then
                    ' move o snow
                    .Y = .Y + .ySpeed
                    .X = .X + .xSpeed
                    ' checkar a screen
                    If .Y > 480 + 64 Then
                        .Y = Rand(0, 100) - 100
                        .X = Rand(0, 640 + 64)
                        .ySpeed = Rand(1, 4)
                        .xSpeed = Rand(0, 4) - 2
                    End If
                    ' draw rain
                    With sRECT
                        .Top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With
                    Engine_BltFast .X + Camera.Left, .Y + Camera.Top, DDS_Sand, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    .Y = Rand(0, 480)
                    .X = Rand(0, 640 + 64)
                    .ySpeed = Rand(1, 4)
                    .xSpeed = Rand(0, 4) - 2
                    .Init = True
                End If
            End With
        Next
    End If
    
    ' fire by allstar
    If Map.Weather = WEATHER_FIRE Then '
        'Call DDS_BackBuffer.SetForeColor(RGB(255, 255, 255))
        For i = 1 To MAX_FIREDROPS
            With DropFire(i)
                If .Init = True Then
                    ' move o snow
                    .Y = .Y + .ySpeed
                    .X = .X + .xSpeed
                    ' checkar a screen
                    If .Y > 480 + 64 Then
                        .Y = Rand(0, 100) - 100
                        .X = Rand(0, 640 + 64)
                        .ySpeed = Rand(1, 4)
                        .xSpeed = Rand(0, 4) - 2
                    End If
                    ' draw rain
                    With sRECT
                        .Top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With
                    Engine_BltFast .X + Camera.Left, .Y + Camera.Top, DDS_Fire, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    .Y = Rand(0, 480)
                    .X = Rand(0, 640 + 64)
                    .ySpeed = Rand(1, 4)
                    .xSpeed = Rand(0, 4) - 2
                    .Init = True
                End If
            End With
        Next
    End If
    
    
End Sub

Sub BltMiniMap()
Dim i As Long
Dim X As Integer, Y As Integer
Dim Direction As Byte
Dim CameraX As Long, CameraY As Long
Dim BlockRect As RECT, WarpRect As RECT, ItemRect As RECT, ShopRect As RECT, NpcOtherRect As RECT, PlayerRect As RECT, PlayerPkRect As RECT, NpcAttackerRect As RECT, NpcShopRect As RECT, NadaRect As RECT
Dim MapX As Long, MapY As Long

        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler

        MapX = Map.maxX
        MapY = Map.maxY

        ' ************
        ' *** Nada ***
        ' ************
        With NadaRect
                .Top = 4
                .Bottom = .Top + 4
                .Left = 0
                .Right = .Left + 4
        End With

        ' Defini-lo no minimap
        For X = 0 To MapX
                For Y = 0 To MapY
                        CameraX = Camera.Left + 25 + (X * 4)
                        CameraY = Camera.Top + 25 + (Y * 4)
                        Engine_BltFast CameraX, CameraY, DDS_MiniMap, NadaRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Next Y
        Next X

        ' *****************
        ' *** Atributos ***
        ' *****************

        ' Bloqueio
        With BlockRect
                .Top = 4
                .Bottom = .Top + 4
                .Left = 4
                .Right = .Left + 4
        End With

        ' Warp
        With WarpRect
                .Top = 4
                .Bottom = .Top + 4
                .Left = 8
                .Right = .Left + 4
        End With

        ' Item
        With ItemRect
                .Top = 4
                .Bottom = .Top + 4
                .Left = 12
                .Right = .Left + 4
        End With

        ' Shop
        With ShopRect
                .Top = 4
                .Bottom = .Top + 4
                .Left = 16
                .Right = .Left + 4
        End With

        ' Defini-los no minimap
        For X = 0 To MapX
                For Y = 0 To MapY
                        Select Case Map.Tile(X, Y).Type
                                Case TILE_TYPE_BLOCKED
                                        CameraX = Camera.Left + 25 + (X * 4)
                                        CameraY = Camera.Top + 25 + (Y * 4)
                                        Engine_BltFast CameraX, CameraY, DDS_MiniMap, BlockRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                                Case TILE_TYPE_WARP
                                        CameraX = Camera.Left + 25 + (X * 4)
                                        CameraY = Camera.Top + 25 + (Y * 4)
                                        Engine_BltFast CameraX, CameraY, DDS_MiniMap, WarpRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                                Case TILE_TYPE_ITEM
                                        CameraX = Camera.Left + 25 + (X * 4)
                                        CameraY = Camera.Top + 25 + (Y * 4)
                                        Engine_BltFast CameraX, CameraY, DDS_MiniMap, ItemRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                                Case TILE_TYPE_SHOP
                                        CameraX = Camera.Left + 25 + (X * 4)
                                        CameraY = Camera.Top + 25 + (Y * 4)
                                        Engine_BltFast CameraX, CameraY, DDS_MiniMap, ShopRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                        End Select
                Next Y
        Next X

        ' **************
        ' *** Player ***
        ' **************

        ' Normal
        With PlayerRect
                .Top = 0
                .Bottom = .Top + 4
                .Left = 4
                .Right = .Left + 4
        End With

        ' Pk
        With PlayerPkRect
                .Top = 0
                .Bottom = .Top + 4
                .Left = 8
                .Right = .Left + 4
        End With

        ' Defini-los no minimap
        For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                        Select Case Player(i).PK
                                Case 0
                                        X = Player(i).X
                                        Y = Player(i).Y
                                        CameraX = Camera.Left + 25 + (X * 4)
                                        CameraY = Camera.Top + 25 + (Y * 4)
                                        Call DDS_BackBuffer.BltFast(CameraX, CameraY, DDS_MiniMap, PlayerRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                                Case 1
                                        X = Player(i).X
                                        Y = Player(i).Y
                                        CameraX = Camera.Left + 25 + (X * 4)
                                        CameraY = Camera.Top + 25 + (Y * 4)
                                        Call DDS_BackBuffer.BltFast(CameraX, CameraY, DDS_MiniMap, PlayerPkRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End Select
                End If
        Next i

        ' ***********
        ' *** NPC ***
        ' ***********

        ' Atacar ao ser atacado e quando for atacado
        With NpcAttackerRect
                .Top = 0
                .Bottom = .Top + 4
                .Left = 12
                .Right = .Left + 4
        End With

        ' Vendendor
        With NpcShopRect
                .Top = 0
                .Bottom = .Top + 4
                .Left = 16
                .Right = .Left + 4
        End With

        ' Outros
        With NpcOtherRect
                .Top = 0
                .Bottom = .Top + 4
                .Left = 20
                .Right = .Left + 4
        End With

        ' Defini-lo no minimap
        For i = 1 To Npc_HighIndex
                If MapNpc(i).Num > 0 Then
                        Select Case NPC(i).Behaviour
                                Case NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC_BEHAVIOUR_ATTACKWHENATTACKED
                                        X = MapNpc(i).X
                                        Y = MapNpc(i).Y
                                        CameraX = Camera.Left + 25 + (X * 4)
                                        CameraY = Camera.Top + 25 + (Y * 4)
                                        Call DDS_BackBuffer.BltFast(CameraX, CameraY, DDS_MiniMap, NpcAttackerRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                                Case NPC_BEHAVIOUR_SHOPKEEPER
                                        X = MapNpc(i).X
                                        Y = MapNpc(i).Y
                                        CameraX = Camera.Left + 25 + (X * 4)
                                        CameraY = Camera.Top + 25 + (Y * 4)
                                        Call DDS_BackBuffer.BltFast(CameraX, CameraY, DDS_MiniMap, NpcShopRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                                Case Else
                                        X = MapNpc(i).X
                                        Y = MapNpc(i).Y
                                        CameraX = Camera.Left + 25 + (X * 4)
                                        CameraY = Camera.Top + 25 + (Y * 4)
                                        Call DDS_BackBuffer.BltFast(CameraX, CameraY, DDS_MiniMap, NpcOtherRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        End Select
                End If
        Next i

        ' Error handler
        Exit Sub
errorhandler:
        HandleError "BltMiniMap", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
