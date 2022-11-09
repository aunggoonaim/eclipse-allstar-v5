VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   8625
   ClientLeft      =   4995
   ClientTop       =   1380
   ClientWidth     =   11115
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
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   741
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "วิธีคิดอัตราดรอป"
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
      Left            =   9000
      TabIndex        =   91
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aspd มอนส์เตอร์"
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
      Left            =   9000
      TabIndex        =   90
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "บันทึก"
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
      Left            =   3480
      TabIndex        =   32
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
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
      Left            =   6840
      TabIndex        =   31
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ลบทิ้ง"
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
      Left            =   5160
      TabIndex        =   30
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "NPC (ข้อมูล)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   7695
      Begin VB.HScrollBar scrlAlpha 
         Height          =   255
         Left            =   3000
         Max             =   255
         TabIndex        =   97
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   94
         Top             =   7320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   92
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox txtAbsorbMagic 
         Alignment       =   2  'Center
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
         Left            =   6000
         TabIndex        =   87
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtReflectDmg 
         Alignment       =   2  'Center
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
         Left            =   6000
         TabIndex        =   86
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtMATK 
         Alignment       =   2  'Center
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
         Left            =   4320
         TabIndex        =   84
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtRegenMp 
         Alignment       =   2  'Center
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
         Left            =   360
         TabIndex        =   82
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtRegenHp 
         Alignment       =   2  'Center
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
         Left            =   360
         TabIndex        =   80
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtBlock 
         Alignment       =   2  'Center
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
         Left            =   4440
         TabIndex        =   76
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox txtDodge 
         Alignment       =   2  'Center
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
         Left            =   3000
         TabIndex        =   74
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtDEF 
         Alignment       =   2  'Center
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
         Left            =   1800
         TabIndex        =   72
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox scrlCritChange 
         Alignment       =   2  'Center
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
         Left            =   5880
         TabIndex        =   70
         Text            =   "10"
         Top             =   6120
         Width           =   735
      End
      Begin VB.TextBox txtCrit 
         Alignment       =   2  'Center
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
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   66
         Text            =   "0"
         Top             =   5520
         Width           =   615
      End
      Begin VB.TextBox txtEXP_max 
         Alignment       =   2  'Center
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
         Left            =   2760
         TabIndex        =   65
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox scrlAttackSpeed 
         Alignment       =   1  'Right Justify
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
         Left            =   5640
         TabIndex        =   63
         Text            =   "3000"
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Frame fraBoss 
         Caption         =   "บอสโหมด"
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
         Left            =   5640
         TabIndex        =   58
         Top             =   1800
         Width           =   1935
         Begin VB.TextBox txtBossNum 
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
            TabIndex        =   59
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblBoss 
            Caption         =   "หมายเลขสคริปบอส :"
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
            TabIndex        =   60
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame fraSpell 
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
         Height          =   1575
         Left            =   5640
         TabIndex        =   53
         Top             =   120
         Width           =   1935
         Begin VB.HScrollBar scrlSpellNum 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   55
            Top             =   360
            Value           =   1
            Width           =   1695
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   54
            Top             =   1200
            Value           =   1
            Width           =   1695
         End
         Begin VB.Label lblSpellName 
            AutoSize        =   -1  'True
            Caption         =   "ชื่อสกิล : ไม่มี"
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
            Top             =   720
            Width           =   915
         End
         Begin VB.Label lblSpellNum 
            AutoSize        =   -1  'True
            Caption         =   "หมายเลข : 0"
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
            Top             =   960
            Width           =   870
         End
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
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
         Left            =   5640
         TabIndex        =   51
         Text            =   "0"
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Frame fraDrop 
         Caption         =   "ดรอปไอเทม"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   42
         Top             =   5040
         Width           =   5175
         Begin VB.TextBox txtChance 
            Alignment       =   2  'Center
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
            Left            =   3480
            TabIndex        =   46
            Text            =   "0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   45
            Top             =   1560
            Width           =   3495
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            TabIndex        =   44
            Top             =   1920
            Width           =   3495
         End
         Begin VB.HScrollBar scrlDrop 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   43
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.Label lblPer 
            Alignment       =   2  'Center
            Caption         =   "%"
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
            Left            =   2640
            TabIndex        =   61
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "อัตราดรอป : ( ใส่ 0.0001 ถึง 1 )"
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
            TabIndex        =   50
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   2115
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "หมายเลข : 0"
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
            TabIndex        =   49
            Top             =   1560
            Width           =   870
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "ชื่อไอเทม : ไม่มี"
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
            TabIndex        =   48
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "จำนวน : 0"
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
            TabIndex        =   47
            Top             =   1920
            UseMnemonic     =   0   'False
            Width           =   720
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   4680
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.CheckBox chkQuest 
         Caption         =   "เป็นผู้ให้เควส?"
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
         TabIndex        =   41
         Top             =   7440
         Width           =   1815
      End
      Begin VB.HScrollBar scrlQuest 
         Height          =   255
         Left            =   1680
         Max             =   255
         TabIndex        =   39
         Top             =   7440
         Width           =   2895
      End
      Begin VB.ComboBox cmbSound 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtDamage 
         Alignment       =   2  'Center
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
         Left            =   600
         TabIndex        =   35
         Top             =   3240
         Width           =   615
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   3580
         Width           =   1575
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   4080
         ScaleHeight     =   96
         ScaleMode       =   0  'User
         ScaleWidth      =   96
         TabIndex        =   22
         Top             =   120
         Width           =   1440
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
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
         Left            =   600
         TabIndex        =   20
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox cmbBehaviour 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   1440
         List            =   "frmEditor_NPC.frx":3348
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2040
         Width           =   3255
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1680
         Max             =   255
         TabIndex        =   18
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtAttackSay 
         Alignment       =   2  'Center
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
         TabIndex        =   17
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status ของ Npc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1200
         TabIndex        =   6
         Top             =   3840
         Width           =   4095
         Begin VB.TextBox txtLevel 
            Alignment       =   2  'Center
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
            Left            =   3360
            TabIndex        =   78
            Top             =   780
            Width           =   615
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   1440
            Max             =   255
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   2760
            Max             =   255
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   120
            Max             =   255
            TabIndex        =   8
            Top             =   720
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   1440
            Max             =   255
            TabIndex        =   7
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Level :"
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
            Left            =   2760
            TabIndex        =   79
            Top             =   780
            Width           =   585
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "SLv : 0"
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
            Index           =   1
            Left            =   135
            TabIndex        =   16
            Top             =   480
            Width           =   1230
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "ตาบอด : 0%"
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
            Index           =   2
            Left            =   1440
            TabIndex        =   15
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Int : 0"
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
            Index           =   3
            Left            =   2760
            TabIndex        =   14
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Job : 0"
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
            Index           =   4
            Left            =   105
            TabIndex        =   13
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "มึน : 0 %"
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
            Index           =   5
            Left            =   1485
            TabIndex        =   12
            Top             =   960
            Width           =   1110
         End
      End
      Begin VB.TextBox txtHP 
         Alignment       =   2  'Center
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
         Left            =   480
         TabIndex        =   5
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtEXP 
         Alignment       =   2  'Center
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
         Left            =   1920
         TabIndex        =   4
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label lblAlpha 
         Caption         =   "Alpha :"
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
         Left            =   2160
         TabIndex        =   96
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ป้องกันเวทย์ : %"
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
         Left            =   6000
         TabIndex        =   95
         Top             =   7080
         Width           =   1140
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ป้องกันกายภาพ : %"
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
         Left            =   5895
         TabIndex        =   93
         Top             =   6480
         Width           =   1350
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ดูดซับเวทย์มนต์ : %"
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
         Left            =   5820
         TabIndex        =   89
         Top             =   3480
         Width           =   1380
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ความแรงสะท้อน : %"
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
         Left            =   5805
         TabIndex        =   88
         Top             =   2880
         Width           =   1410
      End
      Begin VB.Label lblMATK 
         AutoSize        =   -1  'True
         Caption         =   "MATK :"
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
         Left            =   3720
         TabIndex        =   85
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label lblRegenMp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Regen Mp"
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
         Left            =   240
         TabIndex        =   83
         Top             =   4440
         Width           =   765
      End
      Begin VB.Label lblRegenHp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Regen Hp"
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
         Left            =   240
         TabIndex        =   81
         Top             =   3840
         Width           =   765
      End
      Begin VB.Label lblBlock 
         AutoSize        =   -1  'True
         Caption         =   "สะท้อน (%)"
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
         Left            =   3600
         TabIndex        =   77
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label lblDodge 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "หลบ (%)"
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
         Left            =   2280
         TabIndex        =   75
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label lblDEF 
         AutoSize        =   -1  'True
         Caption         =   "DEF :"
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
         Left            =   1320
         TabIndex        =   73
         Top             =   3240
         Width           =   405
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "/ 10"
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
         Left            =   6720
         TabIndex        =   71
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lblCritChange 
         Alignment       =   2  'Center
         Caption         =   "ความแรงโป๊ะเชะ [x เท่า]"
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
         Left            =   5640
         TabIndex        =   69
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "อัตราโป๊ะเชะ (%)"
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
         Left            =   5640
         TabIndex        =   68
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "%"
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
         Left            =   6840
         TabIndex        =   67
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "-"
         Height          =   255
         Left            =   2640
         TabIndex        =   64
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label lblAttackSpeed 
         Alignment       =   2  'Center
         Caption         =   "ASPD : 3000"
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
         Left            =   5640
         TabIndex        =   62
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "เวลาในการเกิด (วินาที)"
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
         Left            =   5640
         TabIndex        =   52
         Top             =   4080
         UseMnemonic     =   0   'False
         Width           =   1860
      End
      Begin VB.Label lblQuest 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   4575
         TabIndex        =   40
         Top             =   7440
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "เสียง (เมื่อโจมตี) :"
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
         TabIndex        =   37
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ATK :"
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
         TabIndex        =   36
         Top             =   3240
         Width           =   405
      End
      Begin VB.Label lblAnimation 
         Alignment       =   2  'Center
         Caption         =   "อนิเมชั่นโจมตี : None"
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
         Left            =   840
         TabIndex        =   33
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "กราฟฟิค : 0"
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
         TabIndex        =   29
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "ชื่อ :"
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
         Left            =   240
         TabIndex        =   28
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ประเภท :"
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
         Left            =   360
         TabIndex        =   27
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "ระยะมองเห็น : 0"
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
         TabIndex        =   26
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "คำพูด (เฉพาะ Sign) :"
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
         TabIndex        =   25
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Exp :"
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
         Left            =   1440
         TabIndex        =   24
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "HP :"
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
         TabIndex        =   23
         Top             =   2880
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "รายชื่อ NPC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
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
         Height          =   7275
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "เปลี่ยน Array Size"
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
      Left            =   240
      TabIndex        =   0
      Top             =   8040
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DropIndex As Byte
Private SpellIndex As Long

Private Sub cmbBehaviour_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NPC(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Command1_Click()
    frmHelp.Show
End Sub

Private Sub Command2_Click()

MsgBox "ใส่ตัวเลข 100 - 10000 เท่านั้น (มิลลิ วินาที)."

End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite.Max = NumCharacters
    scrlAnimation.Max = MAX_ANIMATIONS
    'ALATAR
    scrlQuest.Max = MAX_QUESTS
    scrlNum.Max = MAX_ITEMS
    scrlDrop.Max = MAX_NPC_DROPS
    '/ALATAR
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Frame4_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub FraConv_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAlpha_Change()
    NPC(EditorIndex).alpha = scrlAlpha.Value
    lblAlpha.Caption = "Alpha : " & scrlAlpha.Value
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnimation.Caption = "อนิเมชั่นโจมตี : " & sString
    NPC(EditorIndex).Animation = scrlAnimation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAttackSpeed_Change()
Dim intSpeed As Integer
Dim dblValue As Double
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If scrlAttackSpeed.text > 10000 Or scrlAttackSpeed.text < 200 Then
    scrlAttackSpeed.text = 3000
    'MsgBox "กรุณาใส่ตัวเลข 200 - 10000 (milli sec)"
End If

intSpeed = scrlAttackSpeed.text
NPC(EditorIndex).AttackSpeed = intSpeed

If intSpeed >= 200 And intSpeed <= 1000 Then
dblValue = Round(1000 / intSpeed, 3)
lblAttackSpeed.Caption = "Aspd : " & dblValue & " ครั้ง/1 s."
ElseIf intSpeed > 1000 Then
dblValue = intSpeed / 1000
lblAttackSpeed.Caption = "Aspd : 1 ครั้ง/ " & dblValue & " s."
Else
' lblAttackSpeed.Caption = "Attack speed: " & intSpeed
End If

' Error handler
Exit Sub
errorhandler:
HandleError "scrlAttackSpeed_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub

End Sub

Private Sub scrlCritChange_Change()

NPC(EditorIndex).CritChange = scrlCritChange.text
If scrlCritChange.text >= 10 Then
    lblCritChange.Caption = "ความแรงโป๊ะเชะ " & scrlCritChange.text / 10 & "x"
Else
    lblCritChange.Caption = "ความแรงโป๊ะเชะ 0" & scrlCritChange.text / 10 & "x"
End If

End Sub

Private Sub scrlDrop_Change()

DropIndex = scrlDrop.Value
    fraDrop.Caption = "ดรอปไอเทม - " & DropIndex
    txtChance.text = NPC(EditorIndex).DropChance(DropIndex)
    scrlNum.Value = NPC(EditorIndex).DropItem(DropIndex)
    scrlValue.Value = NPC(EditorIndex).DropItemValue(DropIndex)
    
    ' Check Item
    If scrlNum.Value > 0 Then
        lblItemName.Caption = "ชื่อไอเทม : " & Trim$(Item(scrlNum.Value).Name)
    Else
        lblItemName.Caption = "ชื่อไอเทม : ไม่มี"
    End If
    
End Sub

Private Sub scrlSpell_Change()
 lblSpellNum.Caption = "หมายเลข : " & scrlSpell.Value
    If scrlSpell.Value > 0 Then
        lblSpellName.Caption = "ชื่อสกิล : " & Trim$(Spell(scrlSpell.Value).Name)
    Else
        lblSpellName.Caption = "ชื่อสกิล : ไม่มี"
    End If
    NPC(EditorIndex).Spell(SpellIndex) = scrlSpell.Value
End Sub

Private Sub scrlSpellNum_Change()
SpellIndex = scrlSpellNum.Value
    fraSpell.Caption = "สกิล - " & SpellIndex
    scrlSpell.Value = NPC(EditorIndex).Spell(SpellIndex)
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite.Caption = "กราฟฟิค : " & scrlSprite.Value
    Call EditorNpc_BltSprite
    NPC(EditorIndex).Sprite = scrlSprite.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRange.Caption = "ระยะมองเห็น : " & scrlRange.Value
    NPC(EditorIndex).Range = scrlRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNum.Caption = "หมายเลข : " & scrlNum.Value
    DropIndex = scrlDrop.Value

    ' Check Item
    If scrlNum.Value > 0 Then
        lblItemName.Caption = "ชื่อไอเทม : " & Trim$(Item(scrlNum.Value).Name)
    Else
        lblItemName.Caption = "ชื่อไอเทม : ไม่มี"
    End If
    
    NPC(EditorIndex).DropItem(DropIndex) = scrlNum.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStat_Change(Index As Integer)
Dim prefix As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            prefix = "SLv. : "
        Case 2
            prefix = "ตาบอด : "
        Case 3
            prefix = "Int : "
        Case 4
            prefix = "Job : "
        Case 5
            prefix = "มึน : "
    End Select
    
    If Index <> 5 And Index <> 2 Then
        lblStat(Index).Caption = prefix & scrlStat(Index).Value
    Else
        lblStat(Index).Caption = prefix & scrlStat(Index).Value & "%"
    End If
    
    NPC(EditorIndex).Stat(Index) = scrlStat(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DropIndex = scrlDrop.Value
    
    lblValue.Caption = "จำนวน : " & scrlValue.Value
    NPC(EditorIndex).DropItemValue(DropIndex) = scrlValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAbsorbMagic_Change()
    If Not Len(txtAbsorbMagic.text) > 0 Then Exit Sub
    If IsNumeric(txtAbsorbMagic.text) Then NPC(EditorIndex).AbsorbMagic = Val(txtAbsorbMagic.text)
End Sub

Private Sub txtAttackSay_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If txtAttackSay.text <> vbNullString Then
        NPC(EditorIndex).AttackSay = txtAttackSay.text
    Else
        NPC(EditorIndex).AttackSay = " "
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtBlock_Change()

' แก้ไข เผื่อบัค
If Not Len(txtBlock.text) > 0 Then Exit Sub
If IsNumeric(txtBlock.text) Then NPC(EditorIndex).Block = Val(txtBlock.text)

End Sub

Private Sub txtBossNum_Change()
If Not IsNumeric(txtBossNum.text) Then Exit Sub

NPC(EditorIndex).BossNum = txtBossNum.text
End Sub

Private Sub txtChance_Change()
       On Error GoTo chanceErr
    
    DropIndex = scrlDrop.Value
    
    If Not IsNumeric(txtChance.text) And Not Right$(txtChance.text, 1) = "%" And Not InStr(1, txtChance.text, "/") > 0 And Not InStr(1, txtChance.text, ".") Then
        txtChance.text = "0"
        NPC(EditorIndex).DropChance(DropIndex) = 0
        Exit Sub
    End If
    
    If Right$(txtChance.text, 1) = "%" Then
        txtChance.text = Left(txtChance.text, Len(txtChance.text) - 1) / 100
    ElseIf InStr(1, txtChance.text, "/") > 0 Then
        Dim i() As String
        i = Split(txtChance.text, "/")
        txtChance.text = Int(i(0) / i(1) * 1000) / 1000
    End If
    
    If txtChance.text > 1 Or txtChance.text < 0 Then
        'Err.Description = "Value must be between 0 and 1!"
        'GoTo chanceErr
    End If
    
    lblPer.Caption = txtChance.text * 100 & "%"
    NPC(EditorIndex).DropChance(DropIndex) = txtChance.text
    Exit Sub
    
chanceErr:
    MsgBox "Invalid entry for chance! " & Err.Description
    txtChance.text = "0"
    NPC(EditorIndex).DropChance(DropIndex) = 0
End Sub

Private Sub txtCrit_Change()
    NPC(EditorIndex).CritRate = txtCrit.text
End Sub

Private Sub txtDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtDamage.text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.text) Then NPC(EditorIndex).Damage = Val(txtDamage.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDEF_Change()

' แก้ไข เผื่อบัค
If Not Len(txtDEF.text) > 0 Then Exit Sub
If IsNumeric(txtDEF.text) Then NPC(EditorIndex).Def = Val(txtDEF.text)

End Sub

Private Sub txtDodge_Change()

' แก้ไข เผื่อบัค
If Not Len(txtDodge.text) > 0 Then Exit Sub
If IsNumeric(txtDodge.text) Then NPC(EditorIndex).Dodge = Val(txtDodge.text)

End Sub

Private Sub txtEXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtEXP.text) > 0 Then Exit Sub
    If IsNumeric(txtEXP.text) Then NPC(EditorIndex).EXP = Val(txtEXP.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEXP_max_Change()

If Not Len(txtEXP_max.text) > 0 Then Exit Sub
If IsNumeric(txtEXP_max.text) Then NPC(EditorIndex).EXP_max = Val(txtEXP_max.text)

End Sub

Private Sub txtHP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtHP.text) > 0 Then Exit Sub
    If IsNumeric(txtHP.text) Then NPC(EditorIndex).HP = Val(txtHP.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtLevel.text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.text) Then NPC(EditorIndex).Level = Val(txtLevel.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtlevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMATK_Change()

If Not Len(txtMATK.text) > 0 Then Exit Sub
If IsNumeric(txtMATK.text) Then NPC(EditorIndex).MATK = Val(txtMATK.text)

End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    NPC(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtReflectDmg_Change()
    If Not Len(txtReflectDmg.text) > 0 Then Exit Sub
    If IsNumeric(txtReflectDmg.text) Then NPC(EditorIndex).ReflectDmg = Val(txtReflectDmg.text)
End Sub

Private Sub txtRegenHp_Change()

' แก้ไข เผื่อบัค
If Not Len(txtRegenHp.text) > 0 Then Exit Sub
If IsNumeric(txtRegenHp.text) Then NPC(EditorIndex).RegenHp = Val(txtRegenHp.text)

End Sub

Private Sub txtRegenMp_Change()

' แก้ไข เผื่อบัค
If Not Len(txtRegenMp.text) > 0 Then Exit Sub
If IsNumeric(txtRegenMp.text) Then NPC(EditorIndex).RegenMp = Val(txtRegenMp.text)

End Sub

Private Sub txtSpawnSecs_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtSpawnSecs.text) > 0 Then Exit Sub
    NPC(EditorIndex).SpawnSecs = Val(txtSpawnSecs.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        NPC(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        NPC(EditorIndex).Sound = "None."
    End If
    
    ' sound when select
    PlaySound cmbSound.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'ALATAR

Private Sub chkQuest_Click()
    NPC(EditorIndex).Quest = chkQuest.Value
End Sub

Private Sub scrlQuest_Change()
    lblQuest = scrlQuest.Value
    NPC(EditorIndex).QuestNum = scrlQuest.Value
End Sub

'/ALATAR
