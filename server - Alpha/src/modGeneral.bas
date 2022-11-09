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
    Call SetStatus("���ѧ���������������...")

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    Call SetStatus("���ѧ��������Ἱ���...")
    Call SpawnAllMapsItems
    Call SetStatus("���ѧ��Ǩ�ͺ npc ��Ἱ���...")
    Call SpawnAllMapNpcs
    Call SetStatus("���ѧ���ҧ����鹷캹Ἱ���...")
    Call SpawnAllMapGlobalEvents
    Call SetStatus("���ѧᤪ������Ἱ���...")
    Call CreateFullMapCache
    Call SetStatus("���ѧ��Ŵ�к���...")
    Call LoadSystemTray
    Call SetStatus("SystemTray...")
    
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If
    
        Call Set_Default_Guild_Ranks

    Call SetStatus("���ѧ��駤�����Ѻ�ѧ�ҡ Client...")
    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    time2 = GetTickCount
    Call SetStatus("��õ�Ǩ�ͺ�����١��ͧ�Կ������������ ����� " & time2 - time1 & " ������Թҷ�.")
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
    Dim i As Long
    ServerOnline = False
    Call SetStatus("���ѧ�Դ�Կ�����...")
    Call DestroySystemTray
    Call SetStatus("���ѧ૿�����ż�����...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("���ѧ¡��ԡ�����������...")

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
    Call SetStatus("���������...")
    Call ClearTempTiles
    Call SetStatus("����Ἱ���...")
    Call ClearMaps
    Call SetStatus("����������Ἱ���...")
    Call ClearMapItems
    Call SetStatus("���� npc ��Ἱ���...")
    Call ClearMapNpcs
    Call SetStatus("���� npc...")
    Call ClearNpcs
    Call SetStatus("���¡�çҹ...")
    Call ClearResources
    Call SetStatus("��������...")
    Call ClearItems
    Call SetStatus("������ҹ���...")
    Call ClearShops
    Call SetStatus("����ʡ��...")
    Call ClearSpells
    Call SetStatus("����͹������...")
    Call ClearAnimations
    Call SetStatus("�������...")
    Call ClearQuests
    Call SetStatus("�����к���Ҥ�...")
    Call ClearGuilds
    Call SetStatus("����Ἱ���...")
    Call ClearDoors
    Call SetStatus("�����ѵ������§...")
    Call ClearPets
End Sub

Private Sub LoadGameData()
    Call SetStatus("���ѧ��Ŵ�Ҫվ...")
    Call LoadClasses
    Call SetStatus("���ѧ��ŴἹ���...")
    Call LoadMaps
    Call SetStatus("���ѧ��Ŵ����...")
    Call LoadItems
    Call SetStatus("���ѧ��Ŵ npc...")
    Call LoadNpcs
    Call SetStatus("���ѧ��Ŵ��çҹ...")
    Call LoadResources
    Call SetStatus("���ѧ��Ŵ��ҹ���...")
    Call LoadShops
    Call SetStatus("���ѧ��Ŵʡ��...")
    Call LoadSpells
    Call SetStatus("���ѧ��Ŵ͹������...")
    Call LoadAnimations
    Call SetStatus("���ѧ��駤����Ԫ...")
    Call LoadSwitches
    Call SetStatus("���ѧ��駤��������...")
    Call LoadVariables
    Call SetStatus("���ѧ��Ŵ���...")
    Call LoadQuests
    Call SetStatus("���ѧ��Ŵ��е�...")
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
