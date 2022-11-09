VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "คอนโซล"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCPS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCpsLock"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtText"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtChat"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "ผู้เล่น"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwInfo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "แผงควบคุม"
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblExpRate"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblDropRate"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraDatabase"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraServer"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "scrlExpRate"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdExp"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdDropRate"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "scrlDropRate"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "optReset"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "ข้อความ"
      TabPicture(3)   =   "frmServer.frx":170DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtSendBy"
      Tab(3).Control(1)=   "txtToSend"
      Tab(3).Control(2)=   "frmColorSelect"
      Tab(3).Control(3)=   "cmdSendMessage"
      Tab(3).Control(4)=   "Label2"
      Tab(3).Control(5)=   "Label1"
      Tab(3).ControlCount=   6
      Begin VB.CheckBox optReset 
         Caption         =   "รีเซ็ตเมื่อออกเกม ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71880
         TabIndex        =   39
         Top             =   2040
         Width           =   1695
      End
      Begin VB.HScrollBar scrlDropRate 
         Height          =   255
         Left            =   -69840
         Max             =   200
         Min             =   1
         TabIndex        =   37
         Top             =   1680
         Value           =   1
         Width           =   1695
      End
      Begin VB.CommandButton cmdDropRate 
         Caption         =   "นำไปใช้"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69480
         TabIndex        =   36
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtSendBy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74280
         TabIndex        =   31
         Text            =   "[GameMaster]"
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmdExp 
         Caption         =   "นำไปใช้"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69480
         TabIndex        =   30
         Top             =   1080
         Width           =   975
      End
      Begin VB.HScrollBar scrlExpRate 
         Height          =   255
         Left            =   -69840
         Max             =   200
         Min             =   1
         TabIndex        =   29
         Top             =   720
         Value           =   1
         Width           =   1695
      End
      Begin VB.TextBox txtToSend 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74280
         TabIndex        =   25
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Frame frmColorSelect 
         Caption         =   "สีตัวอักษร"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -70800
         TabIndex        =   19
         Top             =   360
         Width           =   1935
         Begin VB.OptionButton OptGrey 
            Caption         =   "เทา"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1680
            Width           =   1575
         End
         Begin VB.OptionButton OptPink 
            Caption         =   "ชมพู"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton OptYellow 
            Caption         =   "เหลือง"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   1575
         End
         Begin VB.OptionButton OptRed 
            Caption         =   "แดง"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton OptWhite 
            Caption         =   "ขาว"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptGreen 
            Caption         =   "เขียว"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton OptBlue 
            Caption         =   "ฟ้า"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdSendMessage 
         Caption         =   "ส่ง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73320
         TabIndex        =   18
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Frame fraServer 
         Caption         =   "เซิฟเวอร์"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -71880
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         Begin VB.CheckBox chkServerLog 
            Caption         =   "Server Log"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "ปิดทันที"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "ปิดปรับปรุง(อีก30วิ)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame fraDatabase 
         Caption         =   "โหลดซ้ำ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   2895
         Begin VB.CommandButton cmdRest 
            Caption         =   "รีสตาร์ทเซิฟ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   27
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "อนิเมชั่น"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   15
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "การงาน"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   14
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "ไอเทม"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "ร้านค้า"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "สกิล"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "แผนที่"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "อาชีพ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtChat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Width           =   7695
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   600
         Width           =   7935
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5106
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "User ID"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ตัวละคร"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label lblDropRate 
         Alignment       =   2  'Center
         Caption         =   "Drop Rate : 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69840
         TabIndex        =   38
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblExpRate 
         Alignment       =   2  'Center
         Caption         =   "Exp Rate : 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69840
         TabIndex        =   28
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "ข้อความที่จะส่ง :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74280
         TabIndex        =   26
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "ส่งโดย :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblCpsLock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[Unlock]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblCPS 
         Caption         =   "CPS: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Player(1).Switches(1) = 1
End Sub

Private Sub Command2_Click()
    Player(1).Switches(1) = 0
End Sub

Private Sub cmdDropRate_Click()
Dim DropRate As Long

    ' lblDropRate.Caption = "Exp x : " & frmServer.scrlDropRate.Value
    DropRate = frmServer.scrlDropRate.Value
    
    If frmServer.scrlDropRate.Value >= 1 Then
        Call GlobalMsg("เกมถูกตั้งให้อัตรา Drop เป็น " & frmServer.scrlDropRate.Value * 100 & "% จากปกติ.", Yellow)
    End If
    
End Sub

Private Sub cmdExp_Click()
Dim EXPRATE As Long

    ' lblExpRate.Caption = "Exp x : " & frmServer.scrlExpRate.Value
    EXPRATE = frmServer.scrlExpRate.Value
    
    If frmServer.scrlExpRate.Value >= 1 Then
        Call GlobalMsg("เกมถูกตั้งให้ผู้เล่นได้รับ Exp " & frmServer.scrlExpRate.Value * 100 & "% จากปกติ.", Yellow)
    End If
    
End Sub

Private Sub cmdRest_Click()
Call Shell("ServerRestarter.bat")
Unload Me
End Sub

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If
End Sub

Private Sub OptGrey_Click()
OptWhite = False
OptRed = False
OptBlue = False
OptGreen = False
OptYellow = False
OptPink = False
End Sub

Private Sub OptPink_Click()
OptWhite = False
OptRed = False
OptBlue = False
OptGreen = False
OptYellow = False
OptGrey = False
End Sub

Private Sub OptYellow_Click()
OptWhite = False
OptRed = False
OptBlue = False
OptGreen = False
OptPink = False
OptGrey = False
End Sub

Private Sub scrlDropRate_Change()
    lblDropRate.Caption = "Drop x " & frmServer.scrlDropRate.Value * 100 & "%."
End Sub

Private Sub scrlExpRate_Change()
    lblExpRate.Caption = "Exp x " & frmServer.scrlExpRate.Value * 100 & "%."
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)

    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub

' ********************
Private Sub chkServerLog_Click()

    ' if its not 0, then its true
    If chkServerLog.Value <= 0 Then
        ServerLog = False
    Else
        ServerLog = True
    End If

End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadClasses_Click()
Dim i As Long
    Call LoadClasses
    Call TextAdd("All classes reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendClasses i
        End If
    Next
End Sub

Private Sub cmdReloadItems_Click()
Dim i As Long
    Call LoadItems
    Call TextAdd("All items reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendItems i
        End If
    Next
End Sub

Private Sub cmdReloadMaps_Click()
Dim i As Long
    Call LoadMaps
    Call TextAdd("All maps reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next
End Sub

Private Sub cmdReloadNPCs_Click()
Dim i As Long
    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendNpcs i
        End If
    Next
End Sub

Private Sub cmdReloadShops_Click()
Dim i As Long
    Call LoadShops
    Call TextAdd("All shops reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendShops i
        End If
    Next
End Sub

Private Sub cmdReloadSpells_Click()
Dim i As Long
    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next
End Sub

Private Sub cmdReloadResources_Click()
Dim i As Long
    Call LoadResources
    Call TextAdd("All Resources reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendResources i
        End If
    Next
End Sub

Private Sub cmdReloadAnimations_Click()
Dim i As Long
    Call LoadAnimations
    Call TextAdd("All Animations reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendAnimations i
        End If
    Next
End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
    End If
End Sub

Private Sub Form_Load()
    Call UsersOnline_Start
    chkServerLog.Value = 1
    ServerLog = True
    lblExpRate.Caption = "Exp x " & scrlExpRate.Value * 100 & "%."
    lblDropRate.Caption = "Drop x " & scrlDropRate.Value * 100 & "%."
    
    frmServer.WindowState = vbMinimized
    
End Sub

Private Sub Form_Resize()

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, BrightRed)
            Call TextAdd("เซิฟเวอร์ : " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Sub UsersOnline_Start()
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (i)

        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If

        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server owner!")
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        CloseSocket (FindPlayer(Name))
    End If

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call ServerBanIndex(FindPlayer(Name))
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "คุณได้เลื่อนตำแหน่งเป็น [GM].", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "คุณได้ถูกปลดจากตำแหน่ง [GM].", BrightRed)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long
    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select

End Sub

'/ SEND MESSAGE FROM SERVER BY DOMINO

Private Sub optWhite_Click()
If OptWhite = True Then
'OptWhite = False
OptRed = False
OptBlue = False
OptGreen = False
OptYellow = False
OptPink = False
OptGrey = False
End If
End Sub

Private Sub optRed_Click()
If OptRed = True Then
OptWhite = False
OptBlue = False
OptGreen = False
OptYellow = False
OptPink = False
OptGrey = False
End If
End Sub

Private Sub optBlue_Click()
If OptBlue = True Then
OptWhite = False
OptRed = False
OptGreen = False
OptYellow = False
OptPink = False
OptGrey = False
End If
End Sub

Private Sub optGreen_Click()
If OptGreen = True Then
OptWhite = False
OptRed = False
OptBlue = False
OptYellow = False
OptPink = False
OptGrey = False
End If
End Sub

Private Sub cmdSendMessage_Click()
' Color
If OptWhite = True Then
    Call GlobalMsg(txtSendBy.Text & " : " & txtToSend.Text, White)
ElseIf OptRed = True Then

    Call GlobalMsg(txtSendBy.Text & " : " & txtToSend.Text, BrightRed)
ElseIf OptBlue = True Then

    Call GlobalMsg(txtSendBy.Text & " : " & txtToSend.Text, BrightCyan)
ElseIf OptGreen = True Then

    Call GlobalMsg(txtSendBy.Text & " : " & txtToSend.Text, BrightGreen)
ElseIf OptYellow = True Then
    
    Call GlobalMsg(txtSendBy.Text & " : " & txtToSend.Text, Yellow)
ElseIf OptPink = True Then

   Call GlobalMsg(txtSendBy.Text & " : " & txtToSend.Text, Pink)
ElseIf OptGrey = True Then

   Call GlobalMsg(txtSendBy.Text & " : " & txtToSend.Text, Grey)

Else
    
Call GlobalMsg(txtSendBy.Text & " : " & txtToSend.Text, White)

End If
End Sub
