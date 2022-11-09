VERSION 5.00
Begin VB.Form frmEditor_Pet 
   Caption         =   "แก้ไขสัตวเลี้ยง"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdArray 
      Caption         =   "เปลี่ยน Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   95
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "รายชื่อ PET"
      Height          =   7815
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7275
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PET (ข้อมูล)"
      Height          =   7815
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtEXP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   64
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtHP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   480
         TabIndex        =   63
         Top             =   2880
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status ของ Npc"
         Height          =   1215
         Left            =   840
         TabIndex        =   50
         Top             =   3840
         Width           =   4095
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   1440
            Max             =   255
            TabIndex        =   56
            Top             =   720
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   120
            Max             =   255
            TabIndex        =   55
            Top             =   720
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   2760
            Max             =   255
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   1440
            Max             =   255
            TabIndex        =   53
            Top             =   240
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   52
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtLevel 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   51
            Top             =   780
            Width           =   615
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "มึน : 0 %"
            Height          =   195
            Index           =   5
            Left            =   1485
            TabIndex        =   62
            Top             =   960
            Width           =   1110
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Job : 0"
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   61
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Int : 0"
            Height          =   195
            Index           =   3
            Left            =   2760
            TabIndex        =   60
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "ตาบอด : 0%"
            Height          =   195
            Index           =   2
            Left            =   1440
            TabIndex        =   59
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "SLv : 0"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   58
            Top             =   480
            Width           =   1230
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Level :"
            Height          =   195
            Left            =   2760
            TabIndex        =   57
            Top             =   780
            Width           =   585
         End
      End
      Begin VB.TextBox txtAttackSay 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   3495
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1680
         Max             =   255
         TabIndex        =   48
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   315
         ItemData        =   "frmEditor_Pet.frx":0000
         Left            =   1440
         List            =   "frmEditor_Pet.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   480
         TabIndex        =   46
         Top             =   720
         Width           =   3015
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1680
         Max             =   255
         TabIndex        =   45
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1440
         Left            =   3600
         ScaleHeight     =   96
         ScaleMode       =   0  'User
         ScaleWidth      =   96
         TabIndex        =   44
         Top             =   120
         Width           =   1440
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   3360
         TabIndex        =   43
         Top             =   3580
         Width           =   1575
      End
      Begin VB.TextBox txtDamage 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   42
         Top             =   3240
         Width           =   615
      End
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2400
         Width           =   3495
      End
      Begin VB.HScrollBar scrlQuest 
         Height          =   255
         Left            =   1680
         Max             =   255
         TabIndex        =   40
         Top             =   7440
         Width           =   2895
      End
      Begin VB.CheckBox chkQuest 
         Caption         =   "เป็นผู้ให้เควส?"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   7440
         Width           =   1815
      End
      Begin VB.Frame fraDrop 
         Caption         =   "ดรอปไอเทม"
         Height          =   2295
         Left            =   120
         TabIndex        =   29
         Top             =   5040
         Width           =   4815
         Begin VB.HScrollBar scrlDrop 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   33
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            TabIndex        =   32
            Top             =   1920
            Width           =   3495
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   31
            Top             =   1560
            Width           =   3495
         End
         Begin VB.TextBox txtChance 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            TabIndex        =   30
            Text            =   "0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   4680
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "จำนวน : 0"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1920
            UseMnemonic     =   0   'False
            Width           =   720
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "ชื่อไอเทม : ไม่มี"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "หมายเลข : 0"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   1560
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "อัตราดรอป : ( ใส่ 0.0001 ถึง 1 )"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   2115
         End
         Begin VB.Label lblPer 
            Alignment       =   2  'Center
            Caption         =   "%"
            Height          =   255
            Left            =   2640
            TabIndex        =   34
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5040
         TabIndex        =   28
         Text            =   "0"
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Frame fraSpell 
         Caption         =   "สกิล"
         Height          =   1575
         Left            =   5040
         TabIndex        =   23
         Top             =   120
         Width           =   1935
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   25
            Top             =   1200
            Value           =   1
            Width           =   1695
         End
         Begin VB.HScrollBar scrlSpellNum 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   24
            Top             =   360
            Value           =   1
            Width           =   1695
         End
         Begin VB.Label lblSpellNum 
            AutoSize        =   -1  'True
            Caption         =   "หมายเลข : 0"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   870
         End
         Begin VB.Label lblSpellName 
            AutoSize        =   -1  'True
            Caption         =   "ชื่อสกิล : ไม่มี"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   915
         End
      End
      Begin VB.Frame fraBoss 
         Caption         =   "บอสโหมด"
         Height          =   975
         Left            =   5040
         TabIndex        =   20
         Top             =   1800
         Width           =   1935
         Begin VB.TextBox txtBossNum 
            Height          =   270
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblBoss 
            Caption         =   "หมายเลขสคริปบอส :"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox scrlAttackSpeed 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   5040
         TabIndex        =   19
         Text            =   "3000"
         Top             =   4920
         Width           =   1815
      End
      Begin VB.TextBox txtEXP_max 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   2760
         TabIndex        =   18
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtCrit 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   5520
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "0"
         Top             =   5520
         Width           =   615
      End
      Begin VB.TextBox scrlCritChange 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   5280
         TabIndex        =   16
         Text            =   "10"
         Top             =   6120
         Width           =   735
      End
      Begin VB.TextBox txtDEF 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtDodge 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtBlock 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4440
         TabIndex        =   13
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox txtRegenHp 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtRegenMp 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtMATK 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtReflectDmg 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5400
         TabIndex        =   9
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtAbsorbMagic 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5400
         TabIndex        =   8
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   7
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   6
         Top             =   7320
         Width           =   975
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "HP :"
         Height          =   195
         Left            =   120
         TabIndex        =   92
         Top             =   2880
         Width           =   315
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Exp :"
         Height          =   195
         Left            =   1440
         TabIndex        =   91
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "คำพูด (เฉพาะ Sign) :"
         Height          =   195
         Left            =   120
         TabIndex        =   90
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "ระยะมองเห็น : 0"
         Height          =   195
         Left            =   120
         TabIndex        =   89
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ประเภท :"
         Height          =   195
         Left            =   360
         TabIndex        =   88
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "ชื่อ :"
         Height          =   195
         Left            =   120
         TabIndex        =   87
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   285
      End
      Begin VB.Label lblSprite 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ภาพตัวละคร : 0"
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label lblAnimation 
         Alignment       =   2  'Center
         Caption         =   "อนิเมชั่นโจมตี : None"
         Height          =   255
         Left            =   840
         TabIndex        =   85
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ATK :"
         Height          =   195
         Left            =   120
         TabIndex        =   84
         Top             =   3240
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "เสียง (เมื่อโจมตี) :"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblQuest 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   4575
         TabIndex        =   82
         Top             =   7440
         Width           =   330
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "เวลาในการเกิด (วินาที)"
         Height          =   195
         Left            =   5040
         TabIndex        =   81
         Top             =   4080
         UseMnemonic     =   0   'False
         Width           =   1860
      End
      Begin VB.Label lblAttackSpeed 
         Alignment       =   2  'Center
         Caption         =   "ASPD : 3000"
         Height          =   255
         Left            =   5040
         TabIndex        =   80
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "-"
         Height          =   255
         Left            =   2640
         TabIndex        =   79
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "%"
         Height          =   255
         Left            =   6240
         TabIndex        =   78
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "อัตราโป๊ะเชะ (%)"
         Height          =   255
         Left            =   5040
         TabIndex        =   77
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label lblCritChange 
         Alignment       =   2  'Center
         Caption         =   "ความแรงโป๊ะเชะ [x เท่า]"
         Height          =   255
         Left            =   5040
         TabIndex        =   76
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "/ 10"
         Height          =   255
         Left            =   6120
         TabIndex        =   75
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lblDEF 
         AutoSize        =   -1  'True
         Caption         =   "DEF :"
         Height          =   195
         Left            =   1320
         TabIndex        =   74
         Top             =   3240
         Width           =   405
      End
      Begin VB.Label lblDodge 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "หลบ (%)"
         Height          =   195
         Left            =   2280
         TabIndex        =   73
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label lblBlock 
         AutoSize        =   -1  'True
         Caption         =   "สะท้อน (%)"
         Height          =   195
         Left            =   3600
         TabIndex        =   72
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label lblRegenHp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Regen Hp"
         Height          =   195
         Left            =   30
         TabIndex        =   71
         Top             =   3840
         Width           =   765
      End
      Begin VB.Label lblRegenMp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Regen Mp"
         Height          =   195
         Left            =   30
         TabIndex        =   70
         Top             =   4440
         Width           =   765
      End
      Begin VB.Label lblMATK 
         AutoSize        =   -1  'True
         Caption         =   "MATK :"
         Height          =   195
         Left            =   3720
         TabIndex        =   69
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ความแรงสะท้อน : %"
         Height          =   195
         Left            =   5205
         TabIndex        =   68
         Top             =   2880
         Width           =   1410
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ดูดซับเวทย์มนต์ : %"
         Height          =   195
         Left            =   5220
         TabIndex        =   67
         Top             =   3480
         Width           =   1380
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ป้องกันกายภาพ : %"
         Height          =   195
         Left            =   5295
         TabIndex        =   66
         Top             =   6480
         Width           =   1350
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ป้องกันเวทย์ : %"
         Height          =   195
         Left            =   5400
         TabIndex        =   65
         Top             =   7080
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ลบทิ้ง"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "บันทึก"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aspd มอนส์เตอร์"
      Height          =   255
      Left            =   8520
      TabIndex        =   1
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "วิธีคิดอัตราดรอป"
      Height          =   255
      Left            =   8520
      TabIndex        =   0
      Top             =   8040
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditor_Pet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
