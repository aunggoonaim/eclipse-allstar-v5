VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13170
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
   Icon            =   "frmEditor_Spell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   878
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4320
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
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
      Left            =   7680
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
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
      Left            =   6000
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "ข้อมูลสกิล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   9735
      Begin VB.Frame fraSupport 
         Caption         =   "อื่นๆ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   6840
         TabIndex        =   68
         Top             =   1920
         Width           =   2775
         Begin VB.TextBox txtS3 
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
            Left            =   1560
            TabIndex        =   88
            Text            =   "0"
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtS2 
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
            Left            =   1560
            TabIndex        =   86
            Text            =   "0"
            Top             =   1560
            Width           =   495
         End
         Begin VB.CheckBox chkCanCancle 
            Caption         =   "ถูกขัดขวางได้?"
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
            Left            =   120
            TabIndex        =   85
            Top             =   720
            Width           =   1335
         End
         Begin VB.Frame fraPassive 
            Caption         =   "รูปแบบสกิลติดตัว"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   0
            TabIndex        =   79
            Top             =   3960
            Width           =   2775
            Begin VB.TextBox txtS4 
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
               Left            =   1680
               TabIndex        =   91
               Text            =   "0"
               Top             =   960
               Width           =   495
            End
            Begin VB.CheckBox chkPATK 
               Caption         =   "ทำงานเมื่อโจมตี"
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
               TabIndex        =   82
               Top             =   360
               Width           =   1455
            End
            Begin VB.CheckBox chkPDEF 
               Caption         =   "ทำงานถูกโจมตี"
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
               Left            =   240
               TabIndex        =   81
               Top             =   600
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.TextBox txtPerSkill 
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
               Left            =   840
               TabIndex        =   80
               Text            =   "0"
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label13 
               Caption         =   "% : LV"
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
               TabIndex        =   92
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               Caption         =   "โอกาศ"
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
               Left            =   240
               TabIndex        =   84
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label10 
               Caption         =   "% +"
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
               Left            =   1320
               TabIndex        =   83
               Top             =   960
               Width           =   375
            End
         End
         Begin VB.CheckBox chkPhysicalDmg 
            Caption         =   "กายภาพ?"
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
            TabIndex        =   74
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtATKPer 
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
            Left            =   240
            TabIndex        =   73
            Text            =   "0"
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtMagicPer 
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
            Left            =   240
            TabIndex        =   72
            Text            =   "0"
            Top             =   2160
            Width           =   735
         End
         Begin VB.CheckBox chkPassive 
            Caption         =   "สกิลติดตัว?"
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
            TabIndex        =   71
            Top             =   3480
            Width           =   1335
         End
         Begin VB.CheckBox chkMagicDmg 
            Caption         =   "เวทย์มนต์ ?"
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
            TabIndex        =   70
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox chkCanMove 
            Caption         =   "เป็นกายภาพ?"
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
            Left            =   120
            TabIndex        =   69
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label12 
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
            Left            =   2280
            TabIndex        =   90
            Top             =   4920
            Width           =   135
         End
         Begin VB.Label Label11 
            Caption         =   "% : LV"
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
            TabIndex        =   89
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "% : LV"
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
            TabIndex        =   87
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblATKPer 
            Alignment       =   2  'Center
            Caption         =   "+ พลังโจมตี"
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
            TabIndex        =   78
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "% +"
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
            Left            =   1080
            TabIndex        =   77
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label lblMagicPer 
            Alignment       =   2  'Center
            Caption         =   "+ โจมตีเวทย์"
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
            TabIndex        =   76
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "% +"
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
            Left            =   1080
            TabIndex        =   75
            Top             =   2160
            Width           =   375
         End
      End
      Begin VB.Frame frmSpellProjectiles 
         Caption         =   "สกิลแบบธนู"
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
         Left            =   6840
         TabIndex        =   59
         Top             =   240
         Width           =   2775
         Begin VB.HScrollBar scrlProjectilePic 
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   480
            Width           =   1215
         End
         Begin VB.HScrollBar scrlProjectileRange 
            Height          =   255
            Left            =   1440
            TabIndex        =   62
            Top             =   480
            Width           =   1215
         End
         Begin VB.HScrollBar scrlProjectileSpeed 
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1080
            Width           =   1215
         End
         Begin VB.HScrollBar scrlProjectileDamage 
            Height          =   255
            Left            =   1440
            TabIndex        =   60
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblProjectilePiC 
            Alignment       =   2  'Center
            Caption         =   "รูป : 0"
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
            TabIndex        =   67
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblProjectileRange 
            Alignment       =   2  'Center
            Caption         =   "ระยะโจมตี : 0"
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
            Left            =   1440
            TabIndex        =   66
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblProjectilesSpeed 
            Alignment       =   2  'Center
            Caption         =   "ความเร็ว : 0"
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
            TabIndex        =   65
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblProjectileDamage 
            Alignment       =   2  'Center
            Caption         =   "พลังโจมตี : 0"
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
            Left            =   1440
            TabIndex        =   64
            Top             =   840
            Width           =   1215
         End
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   6840
         Width           =   1215
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   1440
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   6240
         Width           =   5295
      End
      Begin VB.Frame Frame6 
         Caption         =   "ข้อมูล"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   3255
         Begin VB.TextBox txtS1 
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
            Left            =   1920
            TabIndex        =   93
            Text            =   "0"
            Top             =   1560
            Width           =   495
         End
         Begin VB.HScrollBar scrlStun 
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   5400
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnim 
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   4800
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnimCast 
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   4200
            Width           =   2895
         End
         Begin VB.CheckBox chkAOE 
            Caption         =   "Area of Effect spell?"
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
            TabIndex        =   43
            Top             =   3120
            Width           =   1935
         End
         Begin VB.HScrollBar scrlAOE 
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   3600
            Width           =   3015
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2760
            Width           =   3015
         End
         Begin VB.HScrollBar scrlInterval 
            Height          =   255
            Left            =   1680
            Max             =   60
            TabIndex        =   38
            Top             =   2160
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDuration 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   36
            Top             =   2160
            Width           =   1455
         End
         Begin VB.HScrollBar scrlVital 
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1560
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   1680
            Max             =   3
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   1680
            Max             =   100
            TabIndex        =   20
            Top             =   960
            Width           =   1455
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   18
            Top             =   960
            Width           =   1455
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "+"
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
            Left            =   1680
            TabIndex        =   95
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label14 
            Caption         =   "% : LV"
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
            Left            =   2520
            TabIndex        =   94
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblStun 
            Caption         =   "Stun Duration: None"
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
            TabIndex        =   52
            Top             =   5160
            Width           =   2895
         End
         Begin VB.Label lblAnim 
            Caption         =   "Animation: None"
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
            TabIndex        =   48
            Top             =   4560
            Width           =   2895
         End
         Begin VB.Label lblAnimCast 
            Caption         =   "Cast Anim: None"
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
            TabIndex        =   46
            Top             =   3960
            Width           =   2895
         End
         Begin VB.Label lblAOE 
            Caption         =   "AoE: Self-cast"
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
            Top             =   3360
            Width           =   3015
         End
         Begin VB.Label lblRange 
            Caption         =   "Range: Self-cast"
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
            TabIndex        =   39
            Top             =   2520
            Width           =   3015
         End
         Begin VB.Label lblInterval 
            Caption         =   "Interval: 0s"
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
            Left            =   1680
            TabIndex        =   37
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label lblDuration 
            Caption         =   "Duration: 0s"
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
            TabIndex        =   35
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label lblVital 
            Caption         =   "Vital: 0"
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
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblDir 
            Caption         =   "Dir: Down"
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
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblY 
            Caption         =   "Y: 0"
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
            Left            =   1680
            TabIndex        =   19
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblX 
            Caption         =   "X: 0"
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
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblMap 
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
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ข้อมูลพื้นฐาน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3255
         Begin VB.HScrollBar scrlHP 
            Height          =   255
            Left            =   1800
            Max             =   9999
            TabIndex        =   57
            Top             =   1680
            Width           =   1335
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
            Height          =   480
            Left            =   2640
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   50
            Top             =   5160
            Width           =   480
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   5400
            Width           =   2415
         End
         Begin VB.HScrollBar scrlCool 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   32
            Top             =   4680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   120
            Max             =   200
            TabIndex        =   30
            Top             =   4080
            Width           =   3015
         End
         Begin VB.ComboBox cmbClass 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   3480
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   26
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   24
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMP 
            Height          =   255
            Left            =   120
            Max             =   9999
            TabIndex        =   13
            Top             =   1680
            Width           =   1335
         End
         Begin VB.ComboBox cmbType 
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
            ItemData        =   "frmEditor_Spell.frx":08CA
            Left            =   120
            List            =   "frmEditor_Spell.frx":08E6
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1080
            Width           =   3015
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
            TabIndex        =   9
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblHP 
            Caption         =   "ใช้ Hp : None"
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
            Left            =   1800
            TabIndex        =   58
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblIcon 
            Caption         =   "ไอคอน : None"
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
            TabIndex        =   44
            Top             =   5160
            Width           =   2415
         End
         Begin VB.Label lblCool 
            Caption         =   "คูลดาวน์ : 0 วินาที"
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
            TabIndex        =   31
            Top             =   4440
            Width           =   2535
         End
         Begin VB.Label lblCast 
            Caption         =   "ใช้เวลาร่าย : 0 วินาที"
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
            TabIndex        =   29
            Top             =   3840
            Width           =   2895
         End
         Begin VB.Label Label5 
            Caption         =   "ต้องการอาชีพ :"
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
            TabIndex        =   27
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label lblAccess 
            Caption         =   "ต้องการระดับ : None"
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
            TabIndex        =   25
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label lblLevel 
            Caption         =   "ต้องการเลเวล : None"
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
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblMP 
            Caption         =   "ใช้ MP : None"
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
            TabIndex        =   12
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ชื่อสกิล :"
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
            TabIndex        =   8
            Top             =   240
            Width           =   570
         End
      End
      Begin VB.Label Label4 
         Caption         =   "เสียง :"
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
         TabIndex        =   55
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "คำอธิบายสกิล :"
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
         TabIndex        =   53
         Top             =   6240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "รายชื่อสกิล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
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
         Height          =   6885
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "เปลี่ยนขนาดอาเรย์"
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
      Top             =   7560
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAOE_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkAOE.Value = 0 Then
        Spell(EditorIndex).IsAoE = False
    Else
        Spell(EditorIndex).IsAoE = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkAOE_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkCanCancle_Click()
    Spell(EditorIndex).CanCancle = chkCanCancle.Value
End Sub

Private Sub chkCanMove_Click()
    Spell(EditorIndex).CanMove = chkCanMove.Value
End Sub

Private Sub chkMagicDmg_Click()
    Spell(EditorIndex).MagicDmg = chkMagicDmg.Value
End Sub

Private Sub chkPassive_Click()

    Spell(EditorIndex).Passive = chkPassive.Value
    
    If Spell(EditorIndex).Passive = 1 Then
            fraPassive.Visible = True
        Else
            fraPassive.Visible = False
    End If

End Sub

Private Sub chkPATK_Click()

    Spell(EditorIndex).PATK = chkPATK.Value
    
    If chkPATK.Value = 1 Then
        chkPDEF.Value = 0
    End If

End Sub

Private Sub chkPDEF_Click()

    Spell(EditorIndex).PDEF = chkPDEF.Value
    
    If chkPDEF.Value = 1 Then
        chkPATK.Value = 0
    End If

End Sub

Private Sub chkPhysicalDmg_Click()
    Spell(EditorIndex).PhysicalDmg = chkPhysicalDmg.Value
End Sub

Private Sub cmbClass_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).ClassReq = cmbClass.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClass_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

Spell(EditorIndex).Type = cmbType.ListIndex

If cmbType.ListIndex = 6 Then
     frmSpellProjectiles.Visible = True
Else
     frmSpellProjectiles.Visible = False
End If

' Error handler
Exit Sub
errorhandler:
HandleError "cmbType_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearSpell EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccess_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAccess.Value > 0 Then
        lblAccess.Caption = "ต้องการระดับ : [GM] ขั้น " & scrlAccess.Value
    Else
        lblAccess.Caption = "ต้องการระดับ : None"
    End If
    Spell(EditorIndex).AccessReq = scrlAccess.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccess_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnim.Value > 0 Then
        lblAnim.Caption = "Animation: " & Trim$(Animation(scrlAnim.Value).Name)
    Else
        lblAnim.Caption = "Animation: None"
    End If
    Spell(EditorIndex).SpellAnim = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnimCast.Value > 0 Then
        lblAnimCast.Caption = "Cast Anim: " & Trim$(Animation(scrlAnimCast.Value).Name)
    Else
        lblAnimCast.Caption = "Cast Anim: None"
    End If
    Spell(EditorIndex).CastAnim = scrlAnimCast.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAOE_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAOE.Value > 0 Then
        lblAOE.Caption = "AoE: " & scrlAOE.Value & " tiles."
    Else
        lblAOE.Caption = "AoE: Self-cast"
    End If
    Spell(EditorIndex).AoE = scrlAOE.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAOE_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlCast.Value < 10 Then
        lblCast.Caption = "ใช้เวลาร่าย : " & scrlCast.Value * 100 & " มิลลิวินาที"
        Spell(EditorIndex).CastTime = scrlCast.Value
    Else
        lblCast.Caption = "ใช้เวลาร่าย : " & (scrlCast.Value * 100) / 1000 & " วินาที"
        Spell(EditorIndex).CastTime = scrlCast.Value
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCool_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCool.Caption = "คูลดาวน์ : " & scrlCool.Value & " วินาที"
    Spell(EditorIndex).CDTime = scrlCool.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCool_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDir_Change()
Dim sDir As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlDir.Value
        Case DIR_UP
            sDir = "Up"
        Case DIR_DOWN
            sDir = "Down"
        Case DIR_RIGHT
            sDir = "Right"
        Case DIR_LEFT
            sDir = "Left"
    End Select
    lblDir.Caption = "Dir: " & sDir
    Spell(EditorIndex).Dir = scrlDir.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDir_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDuration_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblDuration.Caption = "Duration: " & scrlDuration.Value & "s"
    Spell(EditorIndex).Duration = scrlDuration.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDuration_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlHP_Change()
If Options.Debug = 1 Then On Error GoTo errorhandler

If scrlHP.Value > 0 Then
         lblHP.Caption = "ใช้ HP : " & scrlHP.Value
Else
         lblHP.Caption = "ใช้ HP : None"
End If
Spell(EditorIndex).HPCost = scrlHP.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlHP_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub scrlIcon_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlIcon.Value > 0 Then
        lblIcon.Caption = "ไอคอน : " & scrlIcon.Value
    Else
        lblIcon.Caption = "ไอคอน : None"
    End If
    Spell(EditorIndex).Icon = scrlIcon.Value
    EditorSpell_BltIcon
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlIcon_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlInterval_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblInterval.Caption = "Interval: " & scrlInterval.Value & "s"
    Spell(EditorIndex).Interval = scrlInterval.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlInterval_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlLevel.Value > 0 Then
        lblLevel.Caption = "ต้องการเลเวล : " & scrlLevel.Value
    Else
        lblLevel.Caption = "ต้องการเลเวล : None"
    End If
    Spell(EditorIndex).LevelReq = scrlLevel.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevel_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMap_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblMap.Caption = "Map: " & scrlMap.Value
    Spell(EditorIndex).Map = scrlMap.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMap_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlMP.Value > 0 Then
        lblMP.Caption = "ใช้ MP : " & scrlMP.Value
    Else
        lblMP.Caption = "ใช้ MP : None"
    End If
    Spell(EditorIndex).MPCost = scrlMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMP_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlRange.Value > 0 Then
        lblRange.Caption = "Range: " & scrlRange.Value & " tiles."
    Else
        lblRange.Caption = "Range: Self-cast"
    End If
    Spell(EditorIndex).Range = scrlRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStun_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlStun.Value > 0 Then
        lblStun.Caption = "Stun Duration: " & scrlStun.Value & "s"
    Else
        lblStun.Caption = "Stun Duration: None"
    End If
    Spell(EditorIndex).StunDuration = scrlStun.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStun_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlVital_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblVital.Caption = "Vital: " & scrlVital.Value
    Spell(EditorIndex).Vital = scrlVital.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlVital_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblX.Caption = "X: " & scrlX.Value
    Spell(EditorIndex).X = scrlX.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlX_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblY.Caption = "Y: " & scrlY.Value
    Spell(EditorIndex).Y = scrlY.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlY_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtATKPer_Change()

    If txtATKPer.text > -1 Then
        Spell(EditorIndex).ATKPer = txtATKPer.text
    Else
        txtATKPer.text = 0
    End If

End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMagicPer_Change()

    If txtMagicPer.text > -1 Then
        Spell(EditorIndex).MagicPer = txtMagicPer.text
    Else
        txtMagicPer.text = 0
    End If

End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Spell(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Spell(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Spell(EditorIndex).Sound = "None."
    End If
    
    PlaySound cmbSound.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileDamage_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
lblProjectileDamage.Caption = "Damage : " & scrlProjectileDamage.Value
Spell(EditorIndex).ProjecTile.Damage = scrlProjectileDamage.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlProjectilePic_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

' projectile
Private Sub scrlProjectilePic_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
lblProjectilePiC.Caption = "Pic : " & scrlProjectilePic.Value
Spell(EditorIndex).ProjecTile.Pic = scrlProjectilePic.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlProjectilePic_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

' ProjecTile
Private Sub scrlProjectileRange_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
lblProjectileRange.Caption = "Range : " & scrlProjectileRange.Value
Spell(EditorIndex).ProjecTile.Range = scrlProjectileRange.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlProjectileRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

' projectile
Private Sub scrlProjectileSpeed_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
lblProjectilesSpeed.Caption = "Speed : " & scrlProjectileSpeed.Value
Spell(EditorIndex).ProjecTile.Speed = scrlProjectileSpeed.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlRarity_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub txtPerSkill_Change()

    If txtPerSkill.text > 0 Then
        Spell(EditorIndex).PerSkill = txtPerSkill.text
    Else
        txtPerSkill.text = 0
    End If

End Sub

Private Sub txtS1_Change()
    Spell(EditorIndex).S1 = txtS1.text
End Sub

Private Sub txtS2_Change()
    Spell(EditorIndex).S2 = txtS2.text
End Sub

Private Sub txtS3_Change()
    Spell(EditorIndex).S3 = txtS3.text
End Sub

Private Sub txtS4_Change()
    Spell(EditorIndex).S4 = txtS4.text
End Sub
