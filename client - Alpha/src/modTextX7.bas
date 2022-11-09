Attribute VB_Name = "modText"
Option Explicit
Public Chat1(0)

' Text declares
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
    frmMain.Font = Font
    frmMain.FontSize = Size - 5
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetFont", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' GDI text drawing onto buffer
Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal text As String, color As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, 0)
    Call TextOut(hDC, X + 1, Y + 1, text, Len(text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, X, Y, text, Len(text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String
Dim Text2X As Long
Dim Text2Y As Long
Dim GuildString As String
Dim X As Integer
Dim Y As Integer
Dim wrapy As Integer
Dim wrap As Integer, Level As String
Dim text As String
Dim SubString As String
Dim CountNum As Long

X = 1
Y = 1

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check access level
    If GetPlayerPK(Index) = NO Then

        Select Case GetPlayerAccess(Index)
            Case 0
                color = QBColor(White)
            Case 1
                color = QBColor(BrightGreen)
            Case 2
                color = QBColor(BrightGreen)
            Case 3
                color = QBColor(Yellow)
            Case 4
                color = QBColor(Yellow)
        End Select

    Else
        color = QBColor(BrightRed)
    End If
    
    If Player(Index).Level <> MAX_LEVELS Then
        Level = Player(Index).Level
    Else
        Level = "Lv.Max"
    End If
    
    If GetPlayerAccess(Index) <= 0 Then
        Name = Trim$(Player(Index).Name) & " [" & Trim(Level) & "]"
    Else
        Name = Trim$(Player(Index).Name) & " [GM]"
    End If
        
    ' calc pos
    TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).xOffset + (PIC_X \ 2) - getWidth(TexthDC, Trim$(Name))
        GuildString = Player(Index).GuildName
    Text2X = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).xOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(GuildString)))
        If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - 16
        'Guild TUT
        Text2Y = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset
    Else
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - (DDSD_Character(GetPlayerSprite(Index)).lHeight / 4) + 16
        'Guild TUT
        Text2Y = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - (DDSD_Character(GetPlayerSprite(Index)).lHeight / 4) + 4
    End If

    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, Name, color)
    
    ' Chat above head
text = Player(Index).Message
Do While Len(text) > 45
If Y = 1 Then
wrap = Len(text) Mod 45
Else
wrap = 45
End If
'We are going to split the string by spaces
Dim ChatArray() As String
ChatArray = Split(text, " ")
'If our Array is larger than 1, we'll word wrap
If UBound(ChatArray) > 0 Then
Do Until Asc(Mid$(text, Len(text) - wrap, 1)) = 32
wrap = wrap + 1
Loop
'If the Array has only 1 value, kill our loop and just skip ahead
ElseIf UBound(ChatArray) <= 0 Then
Exit Do
wrap = 1
End If
SubString = Right$(text, wrap)
TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).xOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(SubString)))
wrapy = TextY - (X * 20)
'Draw the Message

Call DrawText(TexthDC, TextX, wrapy, SubString, QBColor(White))
text = Left$(text, Len(text) - wrap)

X = X + 1
Y = Y + 1
Loop
TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).xOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(text)))
wrapy = TextY - (X * 20)
'Draw the Message
Call DrawText(TexthDC, TextX, wrapy, text, QBColor(White))
' Error handler
Exit Sub
    
    If Not Player(Index).GuildName = vbNullString Then
        Call DrawText(TexthDC, Text2X, Text2Y, GuildString, color)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String
Dim NPCNum As Long
Dim TypeNPC As String
Dim TBoss As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NPCNum = MapNpc(Index).Num
    
    If NPC(NPCNum).BossNum > 0 Then
        
        If NPC(NPCNum).BossNum = 1 Then
            TypeNPC = "บอสปาร์ตี้"
        End If
        
        If NPC(NPCNum).BossNum = 2 Then
            TypeNPC = "มินิบอส"
        End If
        
        If NPC(NPCNum).BossNum = 3 Then
            TypeNPC = "บอส"
        End If
        
        If NPC(NPCNum).BossNum > 3 Then
            TypeNPC = "?"
        End If
    
    Else
        TypeNPC = NPC(NPCNum).Level
    End If
    
    Select Case NPC(NPCNum).Behaviour
    
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            color = QBColor(BrightRed)
            Name = Trim$(NPC(NPCNum).Name) & " [" & TypeNPC & "]"
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            color = QBColor(White)
            Name = Trim$(NPC(NPCNum).Name) & " [" & TypeNPC & "]"
        Case NPC_BEHAVIOUR_GUARD
            color = QBColor(Yellow)
            Name = Trim$(NPC(NPCNum).Name) & " [" & TypeNPC & "]"
        Case NPC_BEHAVIOUR_FRIENDLY
            color = QBColor(Grey)
            Name = Trim$(NPC(NPCNum).Name) & " [สัตว์เลี้ยง]"
        Case Else
            color = QBColor(BrightGreen)
            Name = Trim$(NPC(NPCNum).Name)

    End Select
        

    TextX = ConvertMapX(MapNpc(Index).X * PIC_X) + MapNpc(Index).xOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(Name)))
    If NPC(NPCNum).Sprite < 1 Or NPC(NPCNum).Sprite > NumCharacters Then
        TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).yOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).yOffset - (DDSD_Character(NPC(NPCNum).Sprite).lHeight / 4) + 16
    End If

    ' Draw name
    
    ' Show (Draw) name and NPC Say
    Dim SAY As String
    SAY = Trim$(NPC(NPCNum).AttackSay)
    If Not SAY = vbNullString Then
        If GetTickCount Mod 20000 < 5000 Then
            Call DrawText(TexthDC, TextX, TextY - 15, SAY, QBColor(White))
            Call DrawText(TexthDC, TextX, TextY, Name, color)
        Else
            Call DrawText(TexthDC, TextX, TextY, Name, color)
        End If
    Else
        SAY = vbNullString
        Call DrawText(TexthDC, TextX, TextY, Name, color)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function BltMapAttributes()
    Dim X As Long
    Dim Y As Long
    Dim tX As Long
    Dim tY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optAttribs.Value Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    With Map.Tile(X, Y)
                        tX = ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5)
                        tY = ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                DrawText TexthDC, tX, tY, "B", QBColor(BrightRed)
                            Case TILE_TYPE_WARP
                                DrawText TexthDC, tX, tY, "W", QBColor(BrightBlue)
                            Case TILE_TYPE_ITEM
                                DrawText TexthDC, tX, tY, "I", QBColor(White)
                            Case TILE_TYPE_NPCAVOID
                                DrawText TexthDC, tX, tY, "N", QBColor(White)
                            Case TILE_TYPE_KEY
                                DrawText TexthDC, tX, tY, "K", QBColor(White)
                            Case TILE_TYPE_KEYOPEN
                                DrawText TexthDC, tX, tY, "O", QBColor(White)
                            Case TILE_TYPE_RESOURCE
                                DrawText TexthDC, tX, tY, "O", QBColor(Green)
                            Case TILE_TYPE_DOOR
                                DrawText TexthDC, tX, tY, "D", QBColor(Brown)
                            Case TILE_TYPE_NPCSPAWN
                                DrawText TexthDC, tX, tY, "S", QBColor(Yellow)
                            Case TILE_TYPE_SHOP
                                DrawText TexthDC, tX, tY, "S", QBColor(BrightBlue)
                            Case TILE_TYPE_BANK
                                DrawText TexthDC, tX, tY, "B", QBColor(Blue)
                            Case TILE_TYPE_HEAL
                                DrawText TexthDC, tX, tY, "H", QBColor(BrightGreen)
                            Case TILE_TYPE_TRAP
                                DrawText TexthDC, tX, tY, "T", QBColor(BrightRed)
                            Case TILE_TYPE_SLIDE
                                DrawText TexthDC, tX, tY, "S", QBColor(BrightCyan)
                            Case TILE_TYPE_CHEST
                                DrawText TexthDC, tX, tY, "C", QBColor(Brown)
                            Case TILE_TYPE_SPRITE
                                DrawText TexthDC, tX, tY, "S", QBColor(Pink)
                            Case TILE_TYPE_ANIMATION
                                DrawText TexthDC, tX, tY, "A", QBColor(Magenta)
                            Case TILE_TYPE_CHECKPOINT
                                DrawText TexthDC, tX, tY, "CP", QBColor(BrightGreen)
                            Case TILE_TYPE_CRAFT
                                DrawText TexthDC, tX, tY, "CR", QBColor(Cyan)
                            Case TILE_TYPE_ONCLICK
                                DrawText TexthDC, tX, tY, "SC", QBColor(BrightBlue)

                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "BltMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' แอคชั่นเมสเซจ Ver 2.0
Sub BltActionMsg(ByVal Index As Long)
    Dim X As Long, Y As Long, i As Long, Time As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' does it exist
    If ActionMsg(Index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            Time = 1500

            If ActionMsg(Index).Y > 0 Then
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 4) - 2
            Else
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 4) + 18
            End If

        Case ACTIONMSG_SCROLL
            Time = 1600
        
            If ActionMsg(Index).Y > 0 Then
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 4) - 2 - (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 4) + 18 + (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> Index Then
                        ClearActionMsg Index
                        Index = i
                    End If
                End If
            Next
            X = (frmMain.picScreen.Width \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
            Y = 425

    End Select
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    If GetTickCount < ActionMsg(Index).Created + Time Then
        Call DrawText(TexthDC, X, Y, ActionMsg(Index).Message, QBColor(ActionMsg(Index).color))
    Else
        ClearActionMsg Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function getWidth(ByVal DC As Long, ByVal text As String) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
     
     ' ทำให้ไม่นับสระ
    getWidth = frmMain.TextWidth(text) / 2 ' ทำให้แสดงผลกลางหน้าจอ
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "getWidth", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub AddText(ByVal Msg As String, ByVal color As Integer)
Dim S As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    S = vbNewLine & Msg
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = QBColor(color)
    frmMain.txtChat.SelText = S
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


'Evilbunnie's DrawnChat system
'Evilbunnie's DrawnChat system, Word wrapping by RyokuPublic , fixed by Rob Janes
Sub DrawChat()
Dim i As Integer
Dim X As Integer
Dim Y As Integer
Dim wrap As Integer
Dim text As String
Dim SubString As String
Dim ChatArray() As String


If Chat1(0) = 1 Then
If Chat1(0) <> 1 Then
    Exit Sub
End If

X = 0
    For i = 1 To 5
        text = Chat(i).text
        Y = 1
        ChatArray = Split(text, " ")

        Do While Len(text) > 60
            If Y = 1 Then
                wrap = Len(text) Mod 60
            Else
                wrap = 60
            End If
                    
            'Our Array has more than 1 Word!
            If UBound(ChatArray) > 0 Then
                Do Until Asc(Mid$(text, Len(text) - wrap, 1)) = 32 'break line on spaces only.
                    wrap = wrap + 1
                Loop
            'Our Array only has 1 Word!
            ElseIf UBound(ChatArray) <= 0 Then
                    Exit Do
                    wrap = 1
                End If
              
            SubString = Right$(text, wrap)
       
            Call DrawText(TexthDC, Camera.Left + 10, (Camera.Bottom - 20) - ((i + X) * 20), SubString, Chat(i).Colour)
            text = Left$(text, Len(text) - wrap)
                
            X = X + 1
            Y = Y + 1
        Loop
       
        Call DrawText(TexthDC, Camera.Left + 10, (Camera.Bottom - 20) - ((i + X) * 20), text, Chat(i).Colour)
    Next i
    End If
    
End Sub

'Evilbunnie's DrawChat system
Public Sub ReOrderChat(ByVal nText As String, nColour As Long)
Dim i As Integer
    
    For i = 19 To 1 Step -1
        Chat(i + 1).text = Chat(i).text
        Chat(i + 1).Colour = Chat(i).Colour
    Next
    
    Chat(1).text = nText
    Chat(1).Colour = nColour
    
End Sub

Public Sub DrawEventName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim Name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    If InMapEditor Then Exit Sub

    color = QBColor(White)

    Name = Trim$(Map.MapEvents(Index).Name)
    
    ' calc pos
    TextX = ConvertMapX(Map.MapEvents(Index).X * PIC_X) + Map.MapEvents(Index).xOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(Name)))
    If Map.MapEvents(Index).GraphicType = 0 Then
        TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - 16
    ElseIf Map.MapEvents(Index).GraphicType = 1 Then
        If Map.MapEvents(Index).GraphicNum < 1 Or Map.MapEvents(Index).GraphicNum > NumCharacters Then
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - 16
        Else
            ' Determine location for text
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - (DDSD_Character(Map.MapEvents(Index).GraphicNum).lHeight / 4) + 16
        End If
    ElseIf Map.MapEvents(Index).GraphicType = 2 Then
        If Map.MapEvents(Index).GraphicY2 = 0 Then
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - ((Map.MapEvents(Index).GraphicY2 - Map.MapEvents(Index).GraphicY) * 32) + 16
        Else
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - 32 + 16
        End If
    End If

    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, Name, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEventName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' เช็คสระในภาษาไทย เพื่อไม่ทำการนับสระ
Public Function CheckThai(text As String) As Long

Dim LenStr As Integer, i As Integer
Dim CountStr As Integer

CountStr = 0
If text = "" Then Exit Function
LenStr = Len(text)
 For i = 1 To LenStr
     If Asc(Mid(text, i, 1)) > 160 And Asc(Mid(text, i, 1)) < 207 Then CountStr = CountStr + 1
  Next i
CheckThai = CountStr

End Function
