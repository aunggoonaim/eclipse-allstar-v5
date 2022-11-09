Attribute VB_Name = "modGeneral"
Option Explicit
' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Public Sub Main()
    Call InitServer
End Sub

Public Sub InitServer()
    Dim i As Long
    Dim F As Long
    Dim time1 As Long
    Dim time2 As Long
    Call InitMessages
    time1 = GetTickCount
    frmServer.Show
    ' Initialize the random-number generator
    Randomize ', seed

    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\Data\", "accounts"
    ChkDir App.Path & "\Data\", "animations"
    ChkDir App.Path & "\Data\", "banks"
    ChkDir App.Path & "\Data\", "items"
    ChkDir App.Path & "\Data\", "logs"
    ChkDir App.Path & "\Data\", "maps"
    ChkDir App.Path & "\Data\", "npcs"
    ChkDir App.Path & "\Data\", "resources"
    ChkDir App.Path & "\Data\", "shops"
    ChkDir App.Path & "\Data\", "spells"
    ChkDir App.Path & "\Data\", "quests"
    ChkDir App.Path & "\Data\", "guilds"
    ChkDir App.Path & "\Data\", "doors"
    

    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' load options, set if they dont exist
    If Not FileExist(App.Path & "\data\options.ini", True) Then
        Options.Game_Name = "Eclipse Allstar"
        Options.Port = 401
        Options.MOTD = "Welcome to Eclipse Allstar."
        Options.Website = "http://gamedd.esy.es/"
        SaveOptions
    Else
        LoadOptions
    End If
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    ' frmServer.Socket(0).RemotePort = Options.Port
    
    ' Init all the player sockets
    Call SetStatus("กำลังส่งอาร์เรย์ผู้เล่น...")

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    Call SetStatus("กำลังเช็คไอเทมบนแผนที่...")
    Call SpawnAllMapsItems
    Call SetStatus("กำลังตรวจสอบ npc บนแผนที่...")
    Call SpawnAllMapNpcs
    Call SetStatus("กำลังสร้างอีเว้นท์บนแผนที่...")
    Call SpawnAllMapGlobalEvents
    Call SetStatus("กำลังแคชข้อมูลแผนที่...")
    Call CreateFullMapCache
    Call SetStatus("กำลังโหลดระบบเกม...")
    Call LoadSystemTray
    Call SetStatus("SystemTray...")
    
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If
    
        Call Set_Default_Guild_Ranks

    Call SetStatus("กำลังตั้งค่ารอรับฟังจาก Client...")
    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    time2 = GetTickCount
    Call SetStatus("การตรวจสอบความถูกต้องเซิฟเวอร์เสร็จสิ้น ในเวลา " & time2 - time1 & " มิลลิวินาที.")
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
    Dim i As Long
    ServerOnline = False
    Call SetStatus("กำลังปิดเซิฟเวอร์...")
    Call DestroySystemTray
    Call SetStatus("กำลังเซฟข้อมูลผู้เล่น...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("กำลังยกเลิกการเชื่อมต่อ...")

    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next

    End
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("เคลียไฟล์ขยะ...")
    Call ClearTempTiles
    Call SetStatus("เคลียแผนที่...")
    Call ClearMaps
    Call SetStatus("เคลียไอเทมบนแผนที่...")
    Call ClearMapItems
    Call SetStatus("เคลีย npc บนแผนที่...")
    Call ClearMapNpcs
    Call SetStatus("เคลีย npc...")
    Call ClearNpcs
    Call SetStatus("เคลียการงาน...")
    Call ClearResources
    Call SetStatus("เคลียไอเทม...")
    Call ClearItems
    Call SetStatus("เคลียร้านค้า...")
    Call ClearShops
    Call SetStatus("เคลียสกิล...")
    Call ClearSpells
    Call SetStatus("เคลียอนิเมชั่น...")
    Call ClearAnimations
    Call SetStatus("เคลียเควส...")
    Call ClearQuests
    Call SetStatus("เคลียระบบสมาคม...")
    Call ClearGuilds
    Call SetStatus("เคลียแผนที่...")
    Call ClearDoors
    Call SetStatus("เคลียสัตว์เลี้ยง...")
    Call ClearPets
End Sub

Private Sub LoadGameData()
    Call SetStatus("กำลังโหลดอาชีพ...")
    Call LoadClasses
    Call SetStatus("กำลังโหลดแผนที่...")
    Call LoadMaps
    Call SetStatus("กำลังโหลดไอเทม...")
    Call LoadItems
    Call SetStatus("กำลังโหลด npc...")
    Call LoadNpcs
    Call SetStatus("กำลังโหลดการงาน...")
    Call LoadResources
    Call SetStatus("กำลังโหลดร้านค้า...")
    Call LoadShops
    Call SetStatus("กำลังโหลดสกิล...")
    Call LoadSpells
    Call SetStatus("กำลังโหลดอนิเมชั่น...")
    Call LoadAnimations
    Call SetStatus("กำลังตั้งค่าสวิช...")
    Call LoadSwitches
    Call SetStatus("กำลังตั้งค่าวาเรีย...")
    Call LoadVariables
    Call SetStatus("กำลังโหลดเควส...")
    Call LoadQuests
    Call SetStatus("กำลังโหลดประตู...")
    Call LoadDoors
End Sub

Public Sub TextAdd(Msg As String)
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function
