VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "เครื่องมือแก้ไขไอเทม"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14925
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   640
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame frmWeapon 
      Caption         =   "สถานะอาวุธ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   13200
      TabIndex        =   175
      Top             =   0
      Width           =   1695
      Begin VB.TextBox txtBuffTime 
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
         Index           =   7
         Left            =   1080
         TabIndex        =   206
         Text            =   "0"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtBuff 
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
         Index           =   7
         Left            =   120
         TabIndex        =   204
         Text            =   "0"
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtBuffTime 
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
         Index           =   5
         Left            =   1080
         TabIndex        =   202
         Text            =   "0"
         Top             =   3000
         Width           =   375
      End
      Begin VB.TextBox txtBuff 
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
         Index           =   5
         Left            =   120
         TabIndex        =   200
         Text            =   "0"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtBuffTime 
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
         Index           =   4
         Left            =   1080
         TabIndex        =   198
         Text            =   "0"
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtBuff 
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
         Index           =   4
         Left            =   120
         TabIndex        =   196
         Text            =   "0"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text12 
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
         Left            =   1080
         TabIndex        =   194
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtKick 
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
         Left            =   120
         TabIndex        =   192
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtBuffTime 
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
         Index           =   2
         Left            =   1080
         TabIndex        =   190
         Text            =   "0"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtBuff 
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
         Index           =   2
         Left            =   120
         TabIndex        =   188
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtBuffTime 
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
         Index           =   8
         Left            =   1080
         TabIndex        =   186
         Text            =   "0"
         Top             =   4800
         Width           =   375
      End
      Begin VB.TextBox txtBuff 
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
         Index           =   8
         Left            =   120
         TabIndex        =   184
         Text            =   "0"
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox txtBuffTime 
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
         Index           =   6
         Left            =   1080
         TabIndex        =   182
         Text            =   "0"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txtBuff 
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
         Index           =   6
         Left            =   120
         TabIndex        =   180
         Text            =   "0"
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtBuffTime 
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
         Index           =   3
         Left            =   1080
         TabIndex        =   178
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtBuff 
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
         Index           =   3
         Left            =   120
         TabIndex        =   176
         Text            =   "0"
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "s"
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
         TabIndex        =   207
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "หายตัว (%)"
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
         TabIndex        =   205
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "s"
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
         TabIndex        =   203
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "หวาดกลัว (%)"
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
         TabIndex        =   201
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "s"
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
         TabIndex        =   199
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "ห้ามฟื้นฟู (%)"
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
         TabIndex        =   197
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "s"
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
         TabIndex        =   195
         Top             =   600
         Width           =   135
      End
      Begin VB.Label lblKick 
         Alignment       =   2  'Center
         Caption         =   "มึน (%)"
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
         TabIndex        =   193
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "s"
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
         TabIndex        =   191
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label lblBuff 
         Alignment       =   2  'Center
         Caption         =   "แช่แข็ง (%)"
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
         TabIndex        =   189
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "s"
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
         TabIndex        =   187
         Top             =   4800
         Width           =   135
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "ใบ้ (%)"
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
         TabIndex        =   185
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "s"
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
         TabIndex        =   183
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "ห้ามโจมตี (%)"
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
         TabIndex        =   181
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "s"
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
         TabIndex        =   179
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "ตาบอด (%)"
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
         TabIndex        =   177
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.Frame fraCoin 
      Caption         =   "ประเภทเงิน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      TabIndex        =   148
      Top             =   3480
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkGold 
         Caption         =   "ทอง"
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
         TabIndex        =   150
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkSilver 
         Caption         =   "เงิน"
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
         TabIndex        =   149
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "ข้อมูลไอเทมสวมใส่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   3240
      TabIndex        =   104
      Top             =   4560
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox scrlSpeedLow 
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
         Left            =   4800
         TabIndex        =   153
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkLHand 
         Caption         =   "อาวุธมือรอง"
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
         Left            =   1200
         TabIndex        =   147
         Top             =   1920
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   118
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cmbTool 
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
         ItemData        =   "frmEditor_Item.frx":3332
         Left            =   1320
         List            =   "frmEditor_Item.frx":3342
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   240
         Width           =   4815
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1560
         Max             =   9999
         TabIndex        =   116
         Top             =   840
         Width           =   1575
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   115
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5280
         Max             =   255
         TabIndex        =   114
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   113
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   112
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   5160
         TabIndex        =   111
         Top             =   1560
         Width           =   975
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   120
         ScaleHeight     =   72
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   400
         TabIndex        =   110
         Top             =   2280
         Width           =   6000
      End
      Begin VB.HScrollBar scrlToolpower 
         Height          =   255
         Left            =   1560
         Max             =   255
         TabIndex        =   109
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkTwoh 
         Caption         =   "ถือ 2 มือ ?"
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
         TabIndex        =   108
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox ChkDagger 
         Caption         =   "ใส่กับมือรอง ?"
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
         Left            =   2400
         TabIndex        =   107
         Top             =   1920
         Width           =   1335
      End
      Begin VB.HScrollBar ScrlDagPdoll 
         Height          =   255
         Left            =   5040
         TabIndex        =   106
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox scrlSpeed 
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
         Left            =   4800
         TabIndex        =   105
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
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
         Left            =   120
         TabIndex        =   130
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "เป็นเครื่องมือ :"
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
         TabIndex        =   129
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   128
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   885
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
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
         Left            =   2160
         TabIndex        =   127
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
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
         Left            =   4440
         TabIndex        =   126
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
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
         Left            =   120
         TabIndex        =   125
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   780
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
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
         Left            =   2160
         TabIndex        =   124
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "+ Aspd : 0 ms."
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
         Left            =   3240
         TabIndex        =   123
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "ภาพที่แสดง : 0"
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
         Left            =   3960
         TabIndex        =   122
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label lblToolpwr 
         Caption         =   "พลังเครื่องมือ : 0"
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
         TabIndex        =   121
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblDagPdoll 
         Caption         =   "Paperdoll : "
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
         Left            =   3960
         TabIndex        =   120
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblSpeedLow 
         AutoSize        =   -1  'True
         Caption         =   "- Aspd : 0 ms."
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
         Left            =   3240
         TabIndex        =   119
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1560
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ความต้องการ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      TabIndex        =   89
      Top             =   3480
      Width           =   6255
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   96
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   840
         Max             =   255
         TabIndex        =   95
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   4080
         Max             =   255
         TabIndex        =   94
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   93
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   840
         Max             =   255
         TabIndex        =   92
         Top             =   240
         Width           =   855
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
         Height          =   270
         Left            =   3840
         TabIndex        =   91
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtMP 
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
         Left            =   5280
         TabIndex        =   90
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
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
         Left            =   1800
         TabIndex        =   103
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   555
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
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
         Left            =   120
         TabIndex        =   102
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
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
         Left            =   3480
         TabIndex        =   101
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
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
         Left            =   1800
         TabIndex        =   100
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
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
         Left            =   120
         TabIndex        =   99
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblHP 
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
         Height          =   255
         Left            =   3480
         TabIndex        =   98
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblMP 
         Caption         =   "MP :"
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
         Left            =   4800
         TabIndex        =   97
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   3495
      Left            =   3240
      TabIndex        =   58
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   74
         Top             =   600
         Width           =   480
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   73
         Top             =   600
         Width           =   1335
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
         Height          =   255
         Left            =   720
         TabIndex        =   72
         Top             =   240
         Width           =   2055
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
         ItemData        =   "frmEditor_Item.frx":3363
         Left            =   120
         List            =   "frmEditor_Item.frx":338B
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   1200
         Width           =   2655
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   70
         Top             =   1320
         Width           =   1095
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   4200
         Max             =   30000
         TabIndex        =   69
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
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
         ItemData        =   "frmEditor_Item.frx":33EB
         Left            =   4200
         List            =   "frmEditor_Item.frx":33F8
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   600
         Width           =   1935
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   3
         TabIndex        =   67
         Top             =   960
         Width           =   1935
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
         Height          =   1215
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   66
         Top             =   1800
         Width           =   2655
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox cmbClassReq 
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
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   2040
         Width           =   2175
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4320
         Max             =   5
         TabIndex        =   63
         Top             =   2400
         Width           =   1815
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4320
         Max             =   99
         TabIndex        =   62
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CheckBox chkReUse 
         Caption         =   "ใช้ซ้ำได้?"
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
         TabIndex        =   61
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtDelayUse 
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
         TabIndex        =   60
         Text            =   "0"
         Top             =   3120
         Width           =   615
      End
      Begin VB.HScrollBar scrlStackLimit 
         Height          =   255
         Left            =   4320
         Max             =   99
         TabIndex        =   59
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   88
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   87
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   285
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "อนิเมชั่น : ไม่มี"
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
         Left            =   2880
         TabIndex        =   86
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "ราคา : 0"
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
         Left            =   2880
         TabIndex        =   85
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "ผูกประเภท :"
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
         Left            =   2880
         TabIndex        =   84
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity : 0"
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
         Left            =   2880
         TabIndex        =   83
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label3 
         Caption         =   "คำอธิบาย :"
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
         TabIndex        =   82
         Top             =   1560
         Width           =   975
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
         Left            =   2880
         TabIndex        =   81
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   2880
         TabIndex        =   80
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "ต้องการระดับ : 0"
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
         Left            =   2880
         TabIndex        =   79
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "ต้องการเลเวล : 0"
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
         Left            =   2880
         TabIndex        =   78
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label lblDelayUse 
         Alignment       =   2  'Center
         Caption         =   "ดีเลย์ :"
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
         TabIndex        =   77
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "s"
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
         TabIndex        =   76
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label lblStackLimit 
         Caption         =   "เก็บได้ : 0 ชิ้น"
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
         Left            =   2880
         TabIndex        =   75
         Top             =   3120
         Width           =   1215
      End
   End
   Begin VB.Frame fraOther 
      Caption         =   "Other Data"
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
      Left            =   9480
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CheckBox chkClassR7 
         Caption         =   "ซามูไร"
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
         TabIndex        =   174
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox chkClassR8 
         Caption         =   "ฮันเตอร์"
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
         TabIndex        =   173
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox chkClassR9 
         Caption         =   "สไนเปอร์"
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
         TabIndex        =   172
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox chkClassR10 
         Caption         =   "แอสแซสซิน"
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
         TabIndex        =   171
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CheckBox chkClassR11 
         Caption         =   "ดาร์คลอร์ด"
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
         TabIndex        =   170
         Top             =   3240
         Width           =   1095
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
         Left            =   2160
         TabIndex        =   169
         Text            =   "0"
         Top             =   6840
         Width           =   975
      End
      Begin VB.TextBox txtDmgReflect 
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
         Left            =   2040
         TabIndex        =   165
         Text            =   "0"
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txtReflect 
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
         TabIndex        =   164
         Text            =   "0"
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txtDMGHigh 
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
         TabIndex        =   163
         Text            =   "100"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtMagicHigh 
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
         TabIndex        =   162
         Text            =   "100"
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtMagicLow 
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
         TabIndex        =   157
         Text            =   "50"
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtDMGLow 
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
         TabIndex        =   154
         Text            =   "50"
         Top             =   3720
         Width           =   735
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
         Height          =   270
         Left            =   2040
         TabIndex        =   152
         Text            =   "0"
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox chkClassR6 
         Caption         =   "วิซาร์ด"
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
         TabIndex        =   146
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CheckBox chkClassR5 
         Caption         =   "พาลาดิน"
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
         TabIndex        =   145
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox chkClassR4 
         Caption         =   "เบอเซิร์ก"
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
         TabIndex        =   144
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CheckBox chkClassR3 
         Caption         =   "การ์เดี้ยน"
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
         TabIndex        =   143
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chkClassR2 
         Caption         =   "เอลฟ์"
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
         TabIndex        =   142
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chkClassR1 
         Caption         =   "มนุษย์"
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
         TabIndex        =   141
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtRegenHP 
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
         Left            =   480
         TabIndex        =   137
         Text            =   "0"
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtRegenMP 
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
         Left            =   2040
         TabIndex        =   136
         Text            =   "0"
         Top             =   4680
         Width           =   975
      End
      Begin VB.CheckBox chkPer2 
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
         Left            =   2880
         TabIndex        =   135
         Top             =   4965
         Width           =   495
      End
      Begin VB.CheckBox chkPer1 
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
         Left            =   1200
         TabIndex        =   134
         Top             =   4965
         Width           =   495
      End
      Begin VB.CheckBox chkDropOnDeath 
         Caption         =   "ไอเทมตกเมื่อตาย ?"
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
         TabIndex        =   133
         Top             =   6480
         Width           =   1695
      End
      Begin VB.TextBox txtVampire 
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
         Left            =   360
         TabIndex        =   132
         Text            =   "0"
         Top             =   1920
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
         Height          =   270
         Left            =   360
         TabIndex        =   49
         Text            =   "0"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtCritRate 
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
         Left            =   2040
         TabIndex        =   48
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtDelayDown 
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
         TabIndex        =   47
         Text            =   "1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNDEF 
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
         Left            =   360
         TabIndex        =   46
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtCritATK 
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
         Left            =   2040
         TabIndex        =   45
         Text            =   "0"
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkAdd1 
         Caption         =   "เพิ่ม"
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
         Top             =   4965
         Width           =   615
      End
      Begin VB.CheckBox chkSub1 
         Caption         =   "ลด"
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
         Left            =   720
         TabIndex        =   43
         Top             =   4965
         Width           =   495
      End
      Begin VB.CheckBox chkAdd2 
         Caption         =   "เพิ่ม"
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
         TabIndex        =   42
         Top             =   4965
         Width           =   615
      End
      Begin VB.CheckBox chkSub2 
         Caption         =   "ลด"
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
         Left            =   2400
         TabIndex        =   41
         Top             =   4965
         Width           =   495
      End
      Begin VB.TextBox txtHPCase 
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
         Left            =   720
         TabIndex        =   40
         Text            =   "0"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtMPCase 
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
         Left            =   2280
         TabIndex        =   39
         Text            =   "0"
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "อัตราดูดซับเวทย์มนต์ % : >"
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
         TabIndex        =   168
         Top             =   6840
         Width           =   1935
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "ความแรงสะท้อน % :"
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
         TabIndex        =   167
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "อัตราสะท้อน % :"
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
         TabIndex        =   166
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "-"
         Height          =   255
         Left            =   1680
         TabIndex        =   161
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "-"
         Height          =   255
         Left            =   1680
         TabIndex        =   160
         Top             =   3720
         Width           =   135
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
         Left            =   2880
         TabIndex        =   159
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblMagicLow 
         Alignment       =   2  'Center
         Caption         =   "MATK"
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
         TabIndex        =   158
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label5 
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
         Left            =   2880
         TabIndex        =   156
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label lblDMGLow 
         Alignment       =   2  'Center
         Caption         =   "ATK"
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
         TabIndex        =   155
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label lblDodge 
         Alignment       =   2  'Center
         Caption         =   "เพิ่มหลบหลีก (%)"
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
         Left            =   1920
         TabIndex        =   151
         Top             =   1730
         Width           =   1335
      End
      Begin VB.Label lblCanEquip 
         Alignment       =   2  'Center
         Caption         =   "อาชีพที่สามารถใส่ได้ :"
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
         Left            =   600
         TabIndex        =   140
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label lblRegenMP 
         Alignment       =   2  'Center
         Caption         =   "เพิ่ม Regen Mp :"
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
         TabIndex        =   139
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label lblRegenHP 
         Alignment       =   2  'Center
         Caption         =   "เพิ่ม Regen Hp :"
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
         TabIndex        =   138
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label lblVampire 
         Alignment       =   2  'Center
         Caption         =   "ดูดเลือด (%)"
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
         TabIndex        =   131
         Top             =   1700
         Width           =   1335
      End
      Begin VB.Label lblMATK 
         Alignment       =   2  'Center
         Caption         =   "เพิ่มโจมตีเวทย์ :"
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
         TabIndex        =   57
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblCritRate 
         Alignment       =   2  'Center
         Caption         =   "เพิ่มอัตราคริ (%)"
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
         Left            =   1920
         TabIndex        =   56
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblDelayDown 
         Alignment       =   2  'Center
         Caption         =   "ลดเวลาร่ายสกิล : "
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
         Top             =   1245
         Width           =   1575
      End
      Begin VB.Label lblDelayDownP 
         Alignment       =   1  'Right Justify
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
         Left            =   960
         TabIndex        =   54
         Top             =   1480
         Width           =   615
      End
      Begin VB.Label lblNDEF 
         Alignment       =   2  'Center
         Caption         =   "โจมตีทะลุเกราะ (%) :"
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
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblCritATK 
         Alignment       =   2  'Center
         Caption         =   "เพิ่มความแรงคริ (%)"
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
         TabIndex        =   52
         Top             =   1245
         Width           =   1575
      End
      Begin VB.Label lblHPCase 
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
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   5280
         Width           =   375
      End
      Begin VB.Label lblMPCase 
         Caption         =   "MP :"
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
         TabIndex        =   50
         Top             =   5280
         Width           =   375
      End
   End
   Begin VB.ComboBox cmbCTool 
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
      ItemData        =   "frmEditor_Item.frx":3421
      Left            =   9600
      List            =   "frmEditor_Item.frx":342E
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame fraRecipe 
      Caption         =   "ใบผสมไอเทม"
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
      Left            =   9480
      TabIndex        =   29
      Top             =   600
      Width           =   3735
      Begin VB.ComboBox cmbCToolReq 
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
         ItemData        =   "frmEditor_Item.frx":3453
         Left            =   1800
         List            =   "frmEditor_Item.frx":3460
         TabIndex        =   37
         Top             =   1080
         Width           =   1575
      End
      Begin VB.HScrollBar scrlResult 
         Height          =   255
         Left            =   1800
         TabIndex        =   36
         Top             =   480
         Width           =   1575
      End
      Begin VB.HScrollBar scrlItem2 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   1575
      End
      Begin VB.HScrollBar scrlItem1 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblResult 
         Caption         =   "ผลลัพธ์ :"
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
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblItem1 
         Caption         =   "ไอเทม1 :"
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
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblItem2 
         Caption         =   "ไอเทม2 :"
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
         Top             =   840
         Width           =   1815
      End
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
      Left            =   3960
      TabIndex        =   4
      Top             =   9135
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "ระบบธนู"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3240
      TabIndex        =   20
      Top             =   8040
      Visible         =   0   'False
      Width           =   6240
      Begin VB.HScrollBar scrlProjectileDamage 
         Height          =   255
         Left            =   4440
         Max             =   9999
         TabIndex        =   28
         Top             =   525
         Width           =   1470
      End
      Begin VB.HScrollBar scrlProjectileSpeed 
         Height          =   255
         Left            =   4440
         Max             =   200
         TabIndex        =   26
         Top             =   180
         Value           =   1
         Width           =   1470
      End
      Begin VB.HScrollBar scrlProjectileRange 
         Height          =   255
         Left            =   1440
         Max             =   100
         TabIndex        =   24
         Top             =   525
         Width           =   1110
      End
      Begin VB.HScrollBar scrlProjectilePic 
         Height          =   255
         Left            =   1080
         Max             =   30
         TabIndex        =   22
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label lblProjectileDamage 
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
         Height          =   225
         Left            =   2880
         TabIndex        =   27
         Top             =   525
         Width           =   1425
      End
      Begin VB.Label lblProjectilesSpeed 
         Caption         =   "ความเร็วธนู : 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2880
         TabIndex        =   25
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label lblProjectileRange 
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
         Height          =   180
         Left            =   150
         TabIndex        =   23
         Top             =   540
         Width           =   1245
      End
      Begin VB.Label lblProjectilePiC 
         BackStyle       =   0  'Transparent
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
         Height          =   270
         Left            =   150
         TabIndex        =   21
         Top             =   240
         Width           =   795
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
      Left            =   225
      TabIndex        =   5
      Top             =   9135
      Width           =   2895
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
      Left            =   7080
      TabIndex        =   3
      Top             =   9135
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
      Left            =   5520
      TabIndex        =   2
      Top             =   9135
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
         Height          =   8445
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "ข้อมูลของไอเทมแบบ Consume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3240
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CheckBox chkInstant 
         Caption         =   "Instant Cast?"
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
         TabIndex        =   19
         Top             =   2760
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   17
         Top             =   2400
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         Caption         =   "Cast Spell: None"
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
         TabIndex        =   18
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   3225
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
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
         TabIndex        =   16
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   3300
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
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
         TabIndex        =   14
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   3375
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
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
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   3375
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "ข้อมูลของสกิล"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1320
         Max             =   255
         Min             =   1
         TabIndex        =   10
         Top             =   720
         Value           =   1
         Width           =   2175
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "ชื่อ : ไม่มี"
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
         TabIndex        =   12
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "เลขสกิล :"
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
         TabIndex        =   11
         Top             =   720
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub chkAdd1_Click()
    Item(EditorIndex).Add1 = chkAdd1.Value
End Sub

Private Sub chkAdd2_Click()
    Item(EditorIndex).Add2 = chkAdd2.Value
End Sub

Private Sub chkClassR1_Click()

Item(EditorIndex).ClassR1 = chkClassR1.Value

End Sub

Private Sub chkClassR10_Click()
    Item(EditorIndex).ClassR10 = chkClassR10.Value
End Sub

Private Sub chkClassR11_Click()
    Item(EditorIndex).ClassR11 = chkClassR11.Value
End Sub

Private Sub chkClassR2_Click()

Item(EditorIndex).ClassR2 = chkClassR2.Value

End Sub

Private Sub chkClassR3_Click()

Item(EditorIndex).ClassR3 = chkClassR3.Value

End Sub

Private Sub chkClassR4_Click()

Item(EditorIndex).ClassR4 = chkClassR4.Value

End Sub

Private Sub chkClassR5_Click()

Item(EditorIndex).ClassR5 = chkClassR5.Value

End Sub

Private Sub chkClassR6_Click()

Item(EditorIndex).ClassR6 = chkClassR6.Value

End Sub

Private Sub chkClassR7_Click()
    Item(EditorIndex).ClassR7 = chkClassR7.Value
End Sub

Private Sub chkClassR8_Click()
    Item(EditorIndex).ClassR8 = chkClassR8.Value
End Sub

Private Sub chkClassR9_Click()
    Item(EditorIndex).ClassR9 = chkClassR9.Value
End Sub

Private Sub ChkDagger_Click()
If Options.Debug = 1 Then On Error GoTo errorhandler

If ChkDagger.Value = 0 Then
Item(EditorIndex).isDagger = False
Item(EditorIndex).LHand = 0
chkLHand.Value = 0
Else
Item(EditorIndex).isDagger = True
Item(EditorIndex).LHand = 1
chkLHand.Value = 1
Item(EditorIndex).isTwoHanded = False
chkTwoh.Value = 0
End If

' Error handler
Exit Sub
errorhandler:
HandleError "chkTwoh", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub chkLHand_Click()

If chkLHand.Value = 0 Then
    Item(EditorIndex).LHand = chkLHand.Value
    ChkDagger.Value = 0
    Item(EditorIndex).isDagger = False
Else
    Item(EditorIndex).LHand = chkLHand.Value
    Item(EditorIndex).isDagger = True
    ChkDagger.Value = 1
    Item(EditorIndex).isTwoHanded = False
    chkTwoh.Value = 0
End If

If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
    If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
        lblDamage.Caption = "Damage : " & scrlDamage.Value
        lblMATK.Caption = "เพิ่มโจมตีเวทย์ :"
    Else
        lblDamage.Caption = "Defense : " & scrlDamage.Value
        lblMATK.Caption = "เพิ่มป้องกันเวทย์เวทย์ :"
            
        If Item(EditorIndex).LHand > 0 And cmbType.ListIndex = ITEM_TYPE_SHIELD Then
            lblDamage.Caption = "Damage : " & scrlDamage.Value
            lblMATK.Caption = "เพิ่มโจมตีเวทย์ :"
        End If
    End If
End If

End Sub

Private Sub chkPer1_Click()
    Item(EditorIndex).Per1 = chkPer1.Value
End Sub

Private Sub chkPer2_Click()
    Item(EditorIndex).Per2 = chkPer2.Value
End Sub

Private Sub chkSub1_Click()
    Item(EditorIndex).Sub1 = chkSub1.Value
End Sub

Private Sub chkSub2_Click()
    Item(EditorIndex).Sub2 = chkSub2.Value
End Sub

Private Sub chkTwoh_Click()
'If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

If chkTwoh.Value = 0 Then
Item(EditorIndex).isTwoHanded = False
Else
Item(EditorIndex).isTwoHanded = True
Item(EditorIndex).isDagger = False
ChkDagger.Value = 0
End If

' Error handler
Exit Sub
errorhandler:
HandleError "chkTwoh", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub cmbBind_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbCTool_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Tool = cmbCTool.ListIndex
End Sub

Private Sub cmbCToolReq_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ToolReq = cmbCToolReq.ListIndex
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).Sound = "None."
    End If
    
    PlaySound cmbSound.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPic.Max = NumItems
    scrlAnim.Max = MAX_ANIMATIONS
    scrlPaperdoll.Max = NumPaperdolls
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            lblDamage.Caption = "Damage : " & scrlDamage.Value
            lblMATK.Caption = "เพิ่มโจมตีเวทย์ :"
            
            Frame4.Visible = True
        Else
            lblDamage.Caption = "Defense : " & scrlDamage.Value
            lblMATK.Caption = "เพิ่มป้องกันเวทย์เวทย์ :"
            
            If Item(EditorIndex).LHand > 0 Then
                lblDamage.Caption = "Damage : " & scrlDamage.Value
                lblMATK.Caption = "เพิ่มโจมตีเวทย์ :"
            End If
        End If
            fraEquipment.Visible = True
        'scrlDamage_Change
    Else
        fraEquipment.Visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_RECIPE) Then
        fraRecipe.Visible = True
    Else
        fraRecipe.Visible = False
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkDropOnDeath_Click()
    Item(EditorIndex).DropOnDeath = chkDropOnDeath.Value
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccessReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "ต้องการระดับ : " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHp.Value
    Item(EditorIndex).AddHP = scrlAddHp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddEXP = scrlAddExp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.Value = 0 Then
        sString = "ไม่มี"
    Else
        sString = Trim$(Animation(scrlAnim.Value).Name)
    End If
    lblAnim.Caption = "อนิเมชั่น : " & sString
    Item(EditorIndex).Animation = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlDagPdoll_Change()
If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
lblDagPdoll.Caption = "ภาพมือรอง : " & ScrlDagPdoll.Value
Item(EditorIndex).Daggerpdoll = ScrlDagPdoll.Value
End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        If (cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
            lblDamage.Caption = "พลังโจมตี : " & scrlDamage.Value
        Else
            If (cmbType.ListIndex = ITEM_TYPE_SHIELD) And Item(EditorIndex).LHand > 0 Then
                lblDamage.Caption = "พลังโจมตี : " & scrlDamage.Value
            Else
                lblDamage.Caption = "พลังป้องกัน : " & scrlDamage.Value
            End If
        End If
    End If
    Item(EditorIndex).Data2 = scrlDamage.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlItem1_Change()
If scrlItem1.Value > 0 Then
        lblItem1.Caption = "ไอเทม1: " & Trim$(Item(scrlItem1.Value).Name)
    Else
        lblItem1.Caption = "ไอเทม1: ไม่มี"
    End If
    
    Item(EditorIndex).Data1 = scrlItem1.Value
End Sub

Private Sub scrlItem2_Change()
If scrlItem2.Value > 0 Then
        lblItem2.Caption = "ไอเทม2: " & Trim$(Item(scrlItem2.Value).Name)
    Else
        lblItem2.Caption = "ไอเทม2: ไม่มี"
    End If
    
    Item(EditorIndex).Data2 = scrlItem2.Value
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "ต้องการเลเวล : " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "ภาพที่แสดง : " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.Value
    Call EditorItem_BltPaperdoll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "รูป : " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value
    Call EditorItem_BltItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPrice_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = "ราคา : " & scrlPrice.Value
    Item(EditorIndex).Price = scrlPrice.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileDamage.Caption = "พลังโจมตี : " & scrlProjectileDamage.Value
    Item(EditorIndex).ProjecTile.Damage = scrlProjectileDamage.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectilePic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilePiC.Caption = "รูป : " & scrlProjectilePic.Value
    Item(EditorIndex).ProjecTile.Pic = scrlProjectilePic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ProjecTile
Private Sub scrlProjectileRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRange.Caption = "ระยะโจมตี : " & scrlProjectileRange.Value
    Item(EditorIndex).ProjecTile.Range = scrlProjectileRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    lblProjectilesSpeed.Caption = "ความเร็วธนู : " & scrlProjectileSpeed.Value
    Item(EditorIndex).ProjecTile.Speed = scrlProjectileSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Select Case scrlRarity.Value
    'Add Cases as levels of rarity you want. For me, there are only 3:
    Case 0 'Item normal
    lblRarity.Caption = "Rarity : ธรรมดา"
    Case 1
    lblRarity.Caption = "Rarity : ธิดา"
    Case 2
    lblRarity.Caption = "Rarity :ตำนาน"
    Case 3
    lblRarity.Caption = "Rarity : เทพ"
    Case Else
    lblRarity.Caption = "Rarity : ไม่มีอยู่จริง"
    End Select
    Item(EditorIndex).Rarity = scrlRarity.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlResult_Change()
If scrlResult.Value > 0 Then
        lblResult.Caption = "ผลลัพธ์: " & Trim$(Item(scrlResult.Value).Name)
    Else
        lblResult.Caption = "ผลลัพธ์: ไม่มี"
    End If
    
    Item(EditorIndex).Data3 = scrlResult.Value
End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "+ Aspd : " & scrlSpeed.text & " ms."
    Item(EditorIndex).Speed = scrlSpeed.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeedLow_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeedLow.Caption = "- Aspd : " & scrlSpeedLow.text & " ms."
    Item(EditorIndex).SpeedLow = scrlSpeedLow.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
Dim text As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "+ Str: "
        Case 2
            text = "+ End: "
        Case 3
            text = "+ Int: "
        Case 4
            text = "+ Agi: "
        Case 5
            text = "+ Will: "
    End Select
            
    lblStatBonus(Index).Caption = text & scrlStatBonus(Index).Value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim text As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "Str: "
        Case 2
            text = "End: "
        Case 3
            text = "Int: "
        Case 4
            text = "Agi: "
        Case 5
            text = "Will: "
    End Select
    
    lblStatReq(Index).Caption = text & scrlStatReq(Index).Value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If Len(Trim$(Spell(scrlSpell.Value).Name)) > 0 Then
        lblSpellName.Caption = "ชื่อ : " & Trim$(Spell(scrlSpell.Value).Name)
    Else
        lblSpellName.Caption = "ชื่อ : ไม่มี"
    End If
    
    lblSpell.Caption = "เลขสกิล : " & scrlSpell.Value
    
    Item(EditorIndex).Data1 = scrlSpell.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlToolpower_Change()
If scrlToolpower.Value > 0 Then
        lblToolpwr.Caption = "พลังเครื่องมือ : " & scrlToolpower.Value
    Else
        lblToolpwr.Caption = "พลังเครื่องมือ : 0"
    End If
    
    Item(EditorIndex).Toolpower = scrlToolpower.Value
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtAbsorbMagic_Change()
        Item(EditorIndex).AbsorbMagic = txtAbsorbMagic.text
End Sub

Private Sub txtBuff_Change(Index As Integer)
    Item(EditorIndex).Buff(Index) = txtBuff(Index).text
End Sub

Private Sub txtBuffTime_Change(Index As Integer)
    Item(EditorIndex).BuffTime(Index) = txtBuffTime(Index).text
End Sub

Private Sub txtCritATK_Change()
        Item(EditorIndex).CritATK = txtCritATK.text
End Sub

Private Sub txtCritRate_Change()
        Item(EditorIndex).CritRate = txtCritRate.text
End Sub

Private Sub txtDelayDown_Change()

    If txtDelayDown.text > 0 Then
        Item(EditorIndex).DelayDown = txtDelayDown.text
        lblDelayDownP.Caption = (1 - txtDelayDown.text) * 100 & " %"
    Else
        Item(EditorIndex).DelayDown = txtDelayDown.text
        lblDelayDownP.Caption = (1 - txtDelayDown.text) * 100 & " %"
    End If
    
    If txtDelayDown.text > 1 Or txtDelayDown.text <= 0 Then
        txtDelayDown.text = 1
        Item(EditorIndex).DelayDown = txtDelayDown.text
    '    MsgBox "กรุณาใส่ตัวเลขระหว่าง 0.01 - 1 เท่านั้น.", vbCritical, "คำเตือน"
    End If

End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDMGHigh_Change()
    Item(EditorIndex).DmgHigh = txtDMGHigh.text
End Sub

Private Sub txtDMGLow_Change()
    Item(EditorIndex).DmgLow = txtDMGLow.text
End Sub

Private Sub txtDmgReflect_Change()
        Item(EditorIndex).DmgReflect = txtDmgReflect.text
End Sub

Private Sub txtDodge_Change()
    Item(EditorIndex).Dodge = txtDodge.text
End Sub

Private Sub txtHP_Change()

    If txtHP.text > 0 Then
        Item(EditorIndex).HP = txtHP.text
    Else
        txtHP.text = 0
        Item(EditorIndex).HP = 0
    End If

End Sub

Private Sub txtHPCase_Change()
        Item(EditorIndex).HPCase = txtHPCase.text
End Sub

Private Sub txtKick_Change()
    Item(EditorIndex).Kick = txtKick.text
End Sub

Private Sub txtMagicHigh_Change()
    Item(EditorIndex).MagicHigh = txtMagicHigh.text
End Sub

Private Sub txtMagicLow_Change()
    Item(EditorIndex).MagicLow = txtMagicLow.text
End Sub

Private Sub txtMATK_Change()
        Item(EditorIndex).MATK = txtMATK.text
End Sub

Private Sub txtMP_Change()

    If txtMP.text > 0 Then
        Item(EditorIndex).MP = txtMP.text
    Else
        txtMP.text = 0
        Item(EditorIndex).MP = 0
    End If

End Sub

Private Sub txtMPCase_Change()
        Item(EditorIndex).MPCase = txtMPCase.text
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtNDEF_Change()
        Item(EditorIndex).NDEF = txtNDEF.text
End Sub

Private Sub txtReflect_Change()
        Item(EditorIndex).Reflect = txtReflect.text
End Sub

Private Sub txtRegenHp_Change()
        Item(EditorIndex).RegenHp = txtRegenHP.text
End Sub

Private Sub txtRegenMp_Change()
        Item(EditorIndex).RegenMp = txtRegenMP.text
End Sub

Private Sub txtVampire_Change()
        Item(EditorIndex).Vampire = txtVampire.text
End Sub
