VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest System"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9030
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   602
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraTasks 
      Caption         =   "Tasks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   3600
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Frame Frame2 
         Height          =   5415
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   2775
         Begin VB.HScrollBar scrlNPC 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   54
            Top             =   1680
            Width           =   2535
         End
         Begin VB.HScrollBar scrlItem 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   53
            Top             =   2280
            Width           =   2535
         End
         Begin VB.HScrollBar scrlAmount 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   52
            Top             =   4680
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   51
            Top             =   2880
            Width           =   2535
         End
         Begin VB.TextBox txtSpeech 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            MaxLength       =   200
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtTaskLog 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   1080
            Width           =   2535
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   48
            Top             =   3480
            Width           =   2535
         End
         Begin VB.CheckBox chkEnd 
            Caption         =   "End Quest Now?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Left            =   120
            TabIndex        =   47
            Top             =   5040
            Width           =   1935
         End
         Begin VB.Label lblNPC 
            AutoSize        =   -1  'True
            Caption         =   "NPC: 0"
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
            Left            =   120
            TabIndex        =   61
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "Item: 0"
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
            Left            =   120
            TabIndex        =   60
            Top             =   2040
            Width           =   480
         End
         Begin VB.Label lblAmount 
            AutoSize        =   -1  'True
            Caption         =   "Amount: 0"
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
            Left            =   120
            TabIndex        =   59
            Top             =   4440
            Width           =   720
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            X1              =   120
            X2              =   2640
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Label lblMap 
            AutoSize        =   -1  'True
            Caption         =   "Map: 0"
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
            Left            =   120
            TabIndex        =   58
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label lblSpeech 
            AutoSize        =   -1  'True
            Caption         =   "Task Speech:"
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
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lblLog 
            AutoSize        =   -1  'True
            Caption         =   "Task Log:"
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
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblResource 
            AutoSize        =   -1  'True
            Caption         =   "Resource: 0"
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
            Left            =   120
            TabIndex        =   55
            Top             =   3240
            Width           =   870
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5415
         Left            =   3000
         TabIndex        =   36
         Top             =   960
         Width           =   2175
         Begin VB.OptionButton optTask 
            Caption         =   "Nothing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Slay NPC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Gather Items"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Talk to NPC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Reach Map"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   41
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Give Item to NPC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   40
            Top             =   1560
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Kill Player"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   39
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Train with Resource"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   7
            Left            =   120
            TabIndex        =   38
            Top             =   2040
            Width           =   1815
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Get from NPC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   8
            Left            =   120
            TabIndex        =   37
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000000&
            X1              =   120
            X2              =   2040
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.HScrollBar scrlTotalTasks 
         Height          =   255
         Left            =   1680
         Max             =   10
         Min             =   1
         TabIndex        =   34
         Top             =   600
         Value           =   1
         Width           =   3495
      End
      Begin VB.TextBox txtQuestLog 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   32
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         Caption         =   "Selected Task: 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Starting Quest Log:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame fraRewards 
      Caption         =   "Rewards"
      Height          =   6495
      Left            =   3600
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.HScrollBar scrlExp 
         Height          =   255
         Left            =   2280
         Max             =   256
         TabIndex        =   68
         Top             =   360
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemRew 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   30
         Top             =   1200
         Value           =   1
         Width           =   2535
      End
      Begin VB.HScrollBar scrlItemRewValue 
         Height          =   255
         Left            =   2880
         Max             =   10
         Min             =   1
         TabIndex        =   29
         Top             =   1200
         Value           =   1
         Width           =   2295
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         Caption         =   "Experience Reward: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   1635
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   5160
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblItemReward 
         AutoSize        =   -1  'True
         Caption         =   "Item Reward: 0 (1)"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1425
      End
   End
   Begin VB.Frame fraRequirements 
      Caption         =   "Requirements"
      Height          =   6495
      Left            =   3600
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.HScrollBar scrlItemNum 
         Height          =   255
         Index           =   0
         Left            =   120
         Max             =   255
         TabIndex        =   65
         Top             =   1920
         Value           =   1
         Width           =   2535
      End
      Begin VB.HScrollBar scrlItemValue 
         Height          =   255
         Index           =   0
         Left            =   2880
         Max             =   10
         Min             =   1
         TabIndex        =   64
         Top             =   1920
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlItemValue 
         Height          =   255
         Index           =   1
         Left            =   2880
         Max             =   10
         Min             =   1
         TabIndex        =   63
         Top             =   2520
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlItemNum 
         Height          =   255
         Index           =   1
         Left            =   120
         Max             =   255
         TabIndex        =   62
         Top             =   2520
         Value           =   1
         Width           =   2535
      End
      Begin VB.CheckBox chkRepeat 
         Caption         =   "Repeatitive Quest?"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   1815
      End
      Begin VB.HScrollBar scrlReq 
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   22
         Top             =   360
         Width           =   2895
      End
      Begin VB.HScrollBar scrlReq 
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   21
         Top             =   720
         Width           =   2895
      End
      Begin VB.HScrollBar scrlReq 
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   20
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   5160
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lblGiveItem 
         AutoSize        =   -1  'True
         Caption         =   "Give Item on Start: 0 (1)"
         Height          =   180
         Left            =   120
         TabIndex        =   67
         Top             =   1680
         Width           =   1875
      End
      Begin VB.Label lblTakeItem 
         AutoSize        =   -1  'True
         Caption         =   "Take Item on the End: 0 (1)"
         Height          =   180
         Left            =   120
         TabIndex        =   66
         Top             =   2280
         Width           =   2100
      End
      Begin VB.Label lblReq 
         AutoSize        =   -1  'True
         Caption         =   "Level Requirement: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label lblReq 
         AutoSize        =   -1  'True
         Caption         =   "Item Requirement: 0"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label lblReq 
         AutoSize        =   -1  'True
         Caption         =   "Quest Requirement: 0"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1665
      End
   End
   Begin VB.Frame fraSpeechs 
      Caption         =   "Quest Speechs"
      Height          =   6495
      Left            =   3600
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox txtChat 
         Height          =   270
         Index           =   1
         Left            =   120
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox txtChat 
         Height          =   270
         Index           =   2
         Left            =   120
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox txtChat 
         Height          =   270
         Index           =   3
         Left            =   120
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label lblQ1 
         AutoSize        =   -1  'True
         Caption         =   "Request Speech:"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblQ2 
         AutoSize        =   -1  'True
         Caption         =   "Meanwhile Speech:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label lblQ3 
         AutoSize        =   -1  'True
         Caption         =   "Finished Speech:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1290
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Quest Title"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Rewards"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   3000
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Tasks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   4320
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Requirements"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Speechs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         MaxLength       =   30
         TabIndex        =   7
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
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
      Left            =   360
      TabIndex        =   5
      Top             =   7800
      Width           =   2895
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   7200
      TabIndex        =   3
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quest List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7080
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////

Option Explicit
Private TempTask As Long

Private Sub Form_Load()
    scrlTotalTasks.Max = MAX_TASKS
    scrlNPC.Max = MAX_NPCS
    scrlItem.Max = MAX_ITEMS
    scrlMap.Max = MAX_MAPS
    scrlResource.Max = MAX_RESOURCES
    scrlAmount.Max = MAX_BYTE
    scrlReq(1).Max = MAX_LEVELS
    scrlReq(2).Max = MAX_ITEMS
    scrlReq(3).Max = MAX_QUESTS
    scrlItemNum(0).Max = MAX_ITEMS
    scrlItemNum(1).Max = MAX_BYTE
    scrlItemValue(0).Max = MAX_ITEMS
    scrlItemValue(1).Max = MAX_BYTE
    scrlExp.Max = MAX_BYTE
    scrlItemRew.Max = MAX_ITEMS
    scrlItemRewValue.Max = MAX_BYTE
End Sub

Private Sub cmdSave_Click()
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        QuestEditorOk
    End If
End Sub

Private Sub cmdCancel_Click()
    QuestEditorCancel
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    ClearQuest EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    QuestEditorInit
End Sub

Private Sub lstIndex_Click()
    QuestEditorInit
End Sub

Private Sub scrlTotalTasks_Change()
    Dim i As Long
    
    lblSelected = "Selected Task: " & scrlTotalTasks.Value
    
    LoadTask EditorIndex, scrlTotalTasks.Value
End Sub

Private Sub optTask_Click(Index As Integer)
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Order = Index
    LoadTask EditorIndex, scrlTotalTasks.Value
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long
    tmpIndex = lstIndex.ListIndex
    Quest(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtQuestLog_Change()
    Quest(EditorIndex).QuestLog = Trim$(txtQuestLog.text)
End Sub

Private Sub txtSpeech_Change()
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Speech = Trim$(txtSpeech.text)
End Sub

Private Sub txtTaskLog_Change()
    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskLog = Trim$(txtTaskLog.text)
End Sub

Private Sub chkRepeat_Click()
    If chkRepeat.Value = 1 Then
        Quest(EditorIndex).Repeat = 1
    Else
        Quest(EditorIndex).Repeat = 0
    End If
End Sub

Private Sub scrlReq_Change(Index As Integer)
    If Index = 1 Then
        lblReq(Index).Caption = "Level Requirement: " & scrlReq(Index).Value
    ElseIf Index = 2 Then
        lblReq(Index).Caption = "Item Requirement: " & scrlReq(Index).Value
    Else
        lblReq(Index).Caption = "Quest Requirement: " & scrlReq(Index).Value
    End If
    Quest(EditorIndex).Requirement(Index) = scrlReq(Index).Value
End Sub

Private Sub scrlExp_Change()
    lblEXP = "Experience Reward: " & scrlExp.Value
    Quest(EditorIndex).RewardExp = scrlExp.Value
End Sub

Private Sub scrlItemRew_Change()
    lblItemReward.Caption = "Item Reward: " & scrlItemRew.Value & " (" & scrlItemRewValue.Value & ")"
    Quest(EditorIndex).RewardItem = scrlItemRew.Value
End Sub

Private Sub scrlItemRewValue_Change()
    lblItemReward.Caption = "Item Reward: " & scrlItemRew.Value & " (" & scrlItemRewValue.Value & ")"
    Quest(EditorIndex).RewardItemAmount = scrlItemRewValue.Value
End Sub

Private Sub txtChat_Change(Index As Integer)
    Quest(EditorIndex).Chat(Index) = Trim$(txtChat(Index).text)
End Sub

Private Sub scrlItemNum_Change(Index As Integer)
    If Index = 0 Then
        lblGiveItem = "Give Item on Start: " & scrlItemNum(Index).Value & " (" & scrlItemValue(Index).Value & ")"
        Quest(EditorIndex).QuestGiveItem = scrlItemNum(Index).Value
    Else
        lblTakeItem = "Take Item on the End: " & scrlItemNum(Index).Value & " (" & scrlItemValue(Index).Value & ")"
        Quest(EditorIndex).QuestRemoveItem = scrlItemNum(Index).Value
    End If
End Sub

Private Sub scrlItemValue_Change(Index As Integer)
    If Index = 0 Then
        lblGiveItem = "Give Item on Start: " & scrlItemNum(Index).Value & " (" & scrlItemValue(Index).Value & ")"
        Quest(EditorIndex).QuestGiveItemValue = scrlItemValue(Index).Value
    Else
        lblTakeItem = "Take Item on the End: " & scrlItemNum(Index).Value & " (" & scrlItemValue(Index).Value & ")"
        Quest(EditorIndex).QuestRemoveItemValue = scrlItemValue(Index).Value
    End If
End Sub

Private Sub scrlAmount_Change()
    lblAmount.Caption = "Amount: " & scrlAmount.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Amount = scrlAmount.Value
End Sub

Private Sub scrlNPC_Change()
    lblNPC.Caption = "NPC: " & scrlNPC.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).NPC = scrlNPC.Value
End Sub

Private Sub scrlItem_Change()
    lblItem.Caption = "Item: " & scrlItem.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Item = scrlItem.Value
End Sub

Private Sub scrlMap_Change()
    lblMap.Caption = "Map: " & scrlMap.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Map = scrlMap.Value
End Sub

Private Sub scrlResource_Change()
    lblResource.Caption = "Resource: " & scrlResource.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Resource = scrlResource.Value
End Sub

Private Sub chkEnd_Click()
    If chkEnd.Value = 1 Then
        Quest(EditorIndex).Task(scrlTotalTasks.Value).QuestEnd = True
    Else
        Quest(EditorIndex).Task(scrlTotalTasks.Value).QuestEnd = False
    End If
End Sub

Private Sub optShowFrame_Click(Index As Integer)
    fraSpeechs.Visible = False
    fraRequirements.Visible = False
    fraRewards.Visible = False
    fraTasks.Visible = False
    
    If optShowFrame(Index).Value = True Then
        Select Case Index
            Case 0
                fraSpeechs.Visible = True
            Case 1
                fraRequirements.Visible = True
            Case 2
                fraRewards.Visible = True
            Case 3
                fraTasks.Visible = True
        End Select
    End If
End Sub
