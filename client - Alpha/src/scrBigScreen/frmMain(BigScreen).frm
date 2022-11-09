VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8370
   ClientLeft      =   4515
   ClientTop       =   1455
   ClientWidth     =   11310
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   222
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain(BigScreen).frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   558
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   754
   Visible         =   0   'False
   Begin VB.PictureBox picInventory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   209
      Top             =   2760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "วิธีเล่นเกม"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      MaskColor       =   &H0080FF80&
      TabIndex        =   208
      Top             =   7440
      Width           =   495
   End
   Begin VB.CommandButton cmdWho 
      BackColor       =   &H008080FF&
      Caption         =   "เติมเงิน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      MaskColor       =   &H0080FF80&
      TabIndex        =   207
      Top             =   6600
      Width           =   615
   End
   Begin VB.PictureBox picParty3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   192
      Top             =   2760
      Visible         =   0   'False
      Width           =   2910
      Begin VB.CommandButton lblBefore3 
         Caption         =   "<"
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
         Left            =   15
         TabIndex        =   194
         Top             =   3840
         Width           =   255
      End
      Begin VB.CommandButton lblNext3 
         Caption         =   ">"
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
         TabIndex        =   193
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   198
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   197
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   196
         Top             =   1935
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   195
         Top             =   2670
         Width           =   2415
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   9
         Left            =   90
         Top             =   735
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   9
         Left            =   90
         Top             =   840
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   10
         Left            =   90
         Top             =   1485
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   10
         Left            =   90
         Top             =   1620
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   11
         Left            =   90
         Top             =   2205
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   11
         Left            =   90
         Top             =   2340
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   12
         Left            =   90
         Top             =   2940
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   12
         Left            =   90
         Top             =   3075
         Visible         =   0   'False
         Width           =   2730
      End
   End
   Begin VB.PictureBox picParty2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   185
      Top             =   2760
      Visible         =   0   'False
      Width           =   2910
      Begin VB.CommandButton lblNext2 
         Caption         =   ">"
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
         TabIndex        =   187
         Top             =   3840
         Width           =   255
      End
      Begin VB.CommandButton lblBefore2 
         Caption         =   "<"
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
         Left            =   15
         TabIndex        =   186
         Top             =   3840
         Width           =   255
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   8
         Left            =   90
         Top             =   3075
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   8
         Left            =   90
         Top             =   2940
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   7
         Left            =   90
         Top             =   2340
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   7
         Left            =   90
         Top             =   2205
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   6
         Left            =   90
         Top             =   1620
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   6
         Left            =   90
         Top             =   1485
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   5
         Left            =   90
         Top             =   840
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   5
         Left            =   90
         Top             =   735
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   191
         Top             =   2670
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   190
         Top             =   1935
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   189
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   188
         Top             =   465
         Width           =   2415
      End
   End
   Begin VB.PictureBox picSpellDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   3240
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   178
      Top             =   9120
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picSpellDescPic 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   179
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblEXPSKILL 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Exp / Exp Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   205
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Image imgEXPSKILL 
         Height          =   240
         Left            =   120
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label lblSpellDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1170
         Left            =   240
         TabIndex        =   181
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label lblSpellName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   180
         Top             =   210
         Width           =   2805
      End
   End
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00B5B5B5&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8010
      Left            =   11280
      ScaleHeight     =   532
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   103
      Top             =   120
      Visible         =   0   'False
      Width           =   2865
      Begin VB.CommandButton cmdAPet 
         Caption         =   "สัตว์เลี้ยง"
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
         TabIndex        =   206
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtAName 
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
         Left            =   240
         TabIndex        =   132
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "เตะ"
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
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "แบน"
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
         TabIndex        =   130
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "วาร์ปเขามาหา"
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
         TabIndex        =   129
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "วาร์ปไปหาเขา"
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
         TabIndex        =   128
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtAMap 
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
         Left            =   960
         TabIndex        =   127
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "วาร์ปไปยัง"
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
         TabIndex        =   126
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "พิกัด"
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
         TabIndex        =   125
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "แจ้งแผนที่บัค"
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
         TabIndex        =   124
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdADestroy 
         Caption         =   "ล้างการแบน"
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
         TabIndex        =   123
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdAItem 
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
         Height          =   255
         Left            =   240
         TabIndex        =   122
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMap 
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
         Height          =   255
         Left            =   240
         TabIndex        =   121
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdANpc 
         Caption         =   "NPC"
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
         TabIndex        =   120
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAResource 
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
         Height          =   255
         Left            =   1440
         TabIndex        =   119
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdAShop 
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
         Height          =   255
         Left            =   240
         TabIndex        =   118
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdASpell 
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
         Height          =   255
         Left            =   1440
         TabIndex        =   117
         Top             =   3720
         Width           =   1095
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   116
         Top             =   5760
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   115
         Top             =   6360
         Value           =   1
         Width           =   2295
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "เสกไอเทมนี้"
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
         TabIndex        =   114
         Top             =   6720
         Width           =   2295
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "ตั้งค่าตัวละคร"
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
         TabIndex        =   113
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "ให้กำเนิด"
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
         TabIndex        =   112
         Top             =   5040
         Width           =   1095
      End
      Begin VB.TextBox txtASprite 
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
         TabIndex        =   111
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtAAccess 
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
         Left            =   1440
         TabIndex        =   110
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "ตั้งค่า Access"
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
         TabIndex        =   109
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton cmdAAnim 
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
         Height          =   255
         Left            =   1440
         TabIndex        =   108
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdLevel 
         Caption         =   "เลเวลอัพ"
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
         TabIndex        =   107
         Top             =   7200
         Width           =   2295
      End
      Begin VB.CommandButton cmdSSMap 
         Caption         =   "ถ่ายรูป"
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
         TabIndex        =   106
         Top             =   7560
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ประตู"
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
         TabIndex        =   105
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "เควส"
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
         TabIndex        =   104
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "หน้าต่างแก้ไขเกม"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   140
         Top             =   120
         Width           =   2865
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   240
         TabIndex        =   139
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   16
         X2              =   168
         Y1              =   144
         Y2              =   144
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "แผนที่#:"
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
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Line Line2 
         X1              =   16
         X2              =   168
         Y1              =   200
         Y2              =   200
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "เครื่องมือช่วยแก้ไขเกม :"
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
         TabIndex        =   137
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Line Line3 
         X1              =   16
         X2              =   168
         Y1              =   304
         Y2              =   304
      End
      Begin VB.Line Line4 
         X1              =   16
         X2              =   168
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lblAItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "เสกไอเทม : ไม่มี"
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
         Left            =   0
         TabIndex        =   136
         Top             =   5520
         Width           =   2775
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "จำนวน : 1"
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
         TabIndex        =   135
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "ตัวละคร#:"
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
         TabIndex        =   134
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access :"
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
         TabIndex        =   133
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line5 
         X1              =   16
         X2              =   168
         Y1              =   472
         Y2              =   472
      End
   End
   Begin VB.PictureBox picCurrency 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
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
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   360
      ScaleHeight     =   1605
      ScaleWidth      =   7200
      TabIndex        =   91
      Top             =   4560
      Visible         =   0   'False
      Width           =   7200
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   92
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblCurrencyCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   95
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label lblCurrencyOk 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Okay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2880
         TabIndex        =   94
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label lblCurrency 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "How many do you want to drop?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   93
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   8025
      ScaleHeight     =   1395
      ScaleWidth      =   3195
      TabIndex        =   90
      Top             =   6840
      Width           =   3225
      Begin VB.Image imgButton 
         Height          =   435
         Index           =   9
         Left            =   1080
         Picture         =   "frmMain(BigScreen).frx":08CA
         Top             =   960
         Width           =   1035
      End
      Begin VB.Image imgButton 
         Height          =   435
         Index           =   4
         Left            =   1080
         Picture         =   "frmMain(BigScreen).frx":209C
         Top             =   480
         Width           =   1035
      End
      Begin VB.Image imgButton 
         Height          =   435
         Index           =   5
         Left            =   2160
         Picture         =   "frmMain(BigScreen).frx":5B8C
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imgButton 
         Height          =   435
         Index           =   6
         Left            =   2160
         Picture         =   "frmMain(BigScreen).frx":96A6
         Top             =   480
         Width           =   1035
      End
      Begin VB.Image imgButton 
         Height          =   435
         Index           =   7
         Left            =   1080
         Picture         =   "frmMain(BigScreen).frx":D3BF
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imgButton 
         Height          =   435
         Index           =   8
         Left            =   2160
         Picture         =   "frmMain(BigScreen).frx":ED6D
         Top             =   960
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Image imgButton 
         Height          =   435
         Index           =   3
         Left            =   0
         Picture         =   "frmMain(BigScreen).frx":1070F
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imgButton 
         Height          =   435
         Index           =   2
         Left            =   0
         Picture         =   "frmMain(BigScreen).frx":1438E
         Top             =   960
         Width           =   1035
      End
      Begin VB.Image imgButton 
         Height          =   435
         Index           =   1
         Left            =   0
         Picture         =   "frmMain(BigScreen).frx":1812A
         Top             =   480
         Width           =   1035
      End
   End
   Begin VB.ComboBox cbMAP 
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
      ItemData        =   "frmMain(BigScreen).frx":1BBBE
      Left            =   6120
      List            =   "frmMain(BigScreen).frx":1BBCE
      Style           =   2  'Dropdown List
      TabIndex        =   89
      Top             =   8010
      Width           =   1185
   End
   Begin VB.PictureBox picSpells 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   84
      Top             =   2760
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.PictureBox picParty 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   77
      Top             =   2760
      Visible         =   0   'False
      Width           =   2910
      Begin VB.CommandButton lblBefore1 
         Caption         =   "<"
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
         Left            =   15
         TabIndex        =   184
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton lblNext1 
         Caption         =   ">"
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
         TabIndex        =   183
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   83
         Top             =   465
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   82
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   81
         Top             =   1935
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   80
         Top             =   2670
         Width           =   2415
      End
      Begin VB.Label lblPartyInvite 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   79
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyLeave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   78
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   735
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   840
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1485
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1620
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2205
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2340
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   4
         Left            =   120
         Top             =   2940
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   4
         Left            =   120
         Top             =   3075
         Visible         =   0   'False
         Width           =   2730
      End
   End
   Begin VB.PictureBox picQuestLog 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   68
      Top             =   2760
      Visible         =   0   'False
      Width           =   2910
      Begin VB.TextBox txtQuestTaskLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   70
         Top             =   1440
         Width           =   2625
      End
      Begin VB.ListBox lstQuestLog 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2205
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   2655
      End
      Begin VB.Image imgQuestButton 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   2
         Left            =   1440
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Image imgQuestButton 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   6
         Left            =   1440
         Top             =   3480
         Width           =   1395
      End
      Begin VB.Image imgQuestButton 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   5
         Left            =   120
         Top             =   3480
         Width           =   1395
      End
      Begin VB.Image imgQuestButton 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   4
         Left            =   1440
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Image imgQuestButton 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   3
         Left            =   120
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "งานที่เกิดขึ้นจริง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "คำพูดสุดท้าย"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   75
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Image imgQuestButton 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   1
         Left            =   120
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "สถานะเควส"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "งานหลัก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   73
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ผลตอบแทน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   71
         Top             =   3600
         Width           =   1215
      End
   End
   Begin VB.PictureBox picCharacter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      Picture         =   "frmMain(BigScreen).frx":1BBF2
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   59
      Top             =   2760
      Visible         =   0   'False
      Width           =   2910
      Begin VB.CommandButton lblTrainStat 
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
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   102
         Top             =   2070
         Width           =   255
      End
      Begin VB.CommandButton lblTrainStat 
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
         Height          =   195
         Index           =   4
         Left            =   2640
         TabIndex        =   101
         Top             =   1830
         Width           =   255
      End
      Begin VB.CommandButton lblTrainStat 
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
         Height          =   195
         Index           =   3
         Left            =   1260
         TabIndex        =   100
         Top             =   2310
         Width           =   255
      End
      Begin VB.CommandButton lblTrainStat 
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
         Height          =   195
         Index           =   2
         Left            =   1260
         TabIndex        =   99
         Top             =   2070
         Width           =   255
      End
      Begin VB.CommandButton lblTrainStat 
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
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   98
         Top             =   1830
         Width           =   255
      End
      Begin VB.PictureBox picFace 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   1140
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   60
         Top             =   900
         Width           =   720
      End
      Begin VB.Label lblNameCls 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "อาชีพ : ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   182
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblCharName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   15
         TabIndex        =   67
         Top             =   495
         Width           =   2850
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   66
         ToolTipText     =   "ค่าพลังชีวิตและพลังโจมตีแบบประชิด"
         Top             =   1830
         Width           =   570
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   2040
         TabIndex        =   65
         ToolTipText     =   "ค่าการหลบหลีกและความเร็วการโจมตี"
         Top             =   1830
         Width           =   585
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   64
         ToolTipText     =   "ค่าพลังป้องกัน อัตราฟื้นฟูพลังชีวิต และอัตราสะท้อน"
         Top             =   2070
         Width           =   570
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   2040
         TabIndex        =   63
         ToolTipText     =   "ค่าการโจมตีระยะไกล(ธนู/ปืน) อัตราคริติคอล และลดเวลาร่ายเวทย์"
         Top             =   2070
         Width           =   585
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   62
         ToolTipText     =   "ค่าพลังโจมตีทางเวทย์ ป้องกันทางเวทย์ และอัตราฟื้นฟูมานา"
         Top             =   2310
         Width           =   570
      End
      Begin VB.Label lblPoints 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2040
         TabIndex        =   61
         Top             =   2310
         Width           =   570
      End
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   50
      Top             =   2760
      Visible         =   0   'False
      Width           =   2910
      Begin VB.CommandButton cmdFps 
         Caption         =   "คอมสเป็คสูง"
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
         TabIndex        =   203
         Top             =   1920
         Width           =   1095
      End
      Begin VB.HScrollBar scrlVolume 
         Height          =   255
         Left            =   360
         Max             =   100
         Min             =   1
         TabIndex        =   97
         Top             =   3480
         Value           =   100
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox Picture3 
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
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   54
         Top             =   840
         Width           =   1935
         Begin VB.OptionButton optMOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optMOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   55
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture4 
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
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   51
         Top             =   1440
         Width           =   1935
         Begin VB.OptionButton optSOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   53
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optSOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Label lblVolume 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ความดังเสียง : 100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   96
         Top             =   3120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   1200
         Width           =   540
      End
   End
   Begin VB.Timer tmrChat 
      Interval        =   5000
      Left            =   15240
      Top             =   360
   End
   Begin VB.PictureBox picEventChat 
      BackColor       =   &H000C0E10&
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
      Height          =   1920
      Left            =   360
      ScaleHeight     =   1920
      ScaleWidth      =   7200
      TabIndex        =   43
      Top             =   4200
      Visible         =   0   'False
      Width           =   7200
      Begin VB.Label lblEventChatContinue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< ต่อไป >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   49
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 1 >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 2 >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 3 >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   46
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 4 >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   45
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblEventChat 
         BackColor       =   &H000C0E10&
         Caption         =   "This is text that appears for an event."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   120
         TabIndex        =   44
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.PictureBox picPet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   36
      Top             =   2760
      Visible         =   0   'False
      Width           =   2910
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   240
         ScaleHeight     =   3435
         ScaleWidth      =   2355
         TabIndex        =   37
         Top             =   240
         Width           =   2415
         Begin VB.Label lblPetAttack 
            Alignment       =   2  'Center
            BackColor       =   &H80000012&
            Caption         =   "โจมตีเป้าหมาย !"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   42
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblPetFollow 
            Alignment       =   2  'Center
            BackColor       =   &H80000012&
            Caption         =   "ตามฉันมา"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   600
            TabIndex        =   41
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lblPetWander 
            Alignment       =   2  'Center
            BackColor       =   &H80000012&
            Caption         =   "ไปเดินเล่น"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   600
            TabIndex        =   40
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblPetDisband 
            Alignment       =   2  'Center
            BackColor       =   &H80000012&
            Caption         =   "เก็บสัตว์เลี้ยง"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   39
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "สัตว์เลี้ยง ;"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.PictureBox picQuestDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   3480
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   30
      Top             =   1800
      Visible         =   0   'False
      Width           =   4425
      Begin VB.Label lblQuestExtra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblQuestName 
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อเควส"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   4215
      End
      Begin VB.Label lblQuestClose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดหน้าต่าง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   2760
         TabIndex        =   33
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblQuestAccept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยอมรับเควส"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   960
         TabIndex        =   32
         Top             =   1680
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblQuestSay 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   4140
      End
   End
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   0
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   4
      Top             =   9120
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picItemDescPic 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   19
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblItemDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1650
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   5
         Top             =   210
         Width           =   2805
      End
   End
   Begin VB.PictureBox picDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
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
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   360
      ScaleHeight     =   2085
      ScaleWidth      =   7200
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   7200
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ตกลง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   29
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label lblDialogue_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Robin has requested a trade. Would you like to accept?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   6735
      End
      Begin VB.Label lblDialogue_Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Request"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   6975
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ยืนยัน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   26
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   4800
         TabIndex        =   25
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.PictureBox picTempInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   6480
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   3
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   7080
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   9
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   7680
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   22
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picSSMap 
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
      Height          =   255
      Left            =   15000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   23
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H80000008&
      Height          =   5745
      Left            =   720
      ScaleHeight     =   383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   7200
      Begin VB.PictureBox picTheirTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   3855
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   12
         Top             =   465
         Width           =   2895
      End
      Begin VB.PictureBox picYourTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   435
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   11
         Top             =   480
         Width           =   2895
      End
      Begin VB.Image imgDeclineTrade 
         Height          =   435
         Left            =   3675
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Image imgAcceptTrade 
         Height          =   435
         Left            =   2475
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label lblTradeStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   5520
         Width           =   5895
      End
      Begin VB.Label lblTheirWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   4500
         Width           =   1815
      End
      Begin VB.Label lblYourWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   4500
         Width           =   1815
      End
   End
   Begin VB.PictureBox picCover 
      Appearance      =   0  'Flat
      BackColor       =   &H00181C21&
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
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   15360
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picHotbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   476
      TabIndex        =   20
      Top             =   7440
      Width           =   7140
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1050
      Left            =   120
      TabIndex        =   1
      Top             =   6375
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   1852
      _Version        =   393217
      BackColor       =   790032
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain(BigScreen).frx":2251D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtMyChat 
      BackColor       =   &H00404040&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1560
      TabIndex        =   2
      Top             =   8040
      Width           =   4425
   End
   Begin VB.PictureBox picBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   5385
      Left            =   720
      ScaleHeight     =   359
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.PictureBox picShop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   3600
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   4125
      Begin VB.PictureBox picShopItems 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   615
         ScaleHeight     =   211
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   7
         Top             =   630
         Width           =   2895
      End
      Begin VB.Image imgLeaveShop 
         Height          =   435
         Left            =   2715
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Image imgShopSell 
         Height          =   435
         Left            =   1545
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Image imgShopBuy 
         Height          =   435
         Left            =   375
         Top             =   4350
         Width           =   1035
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00181C21&
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
      ForeColor       =   &H80000008&
      Height          =   6240
      Left            =   120
      ScaleHeight     =   416
      ScaleMode       =   0  'User
      ScaleWidth      =   736
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   11040
      Begin VB.PictureBox picMe 
         BackColor       =   &H00000000&
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
         Left            =   3480
         ScaleHeight     =   5235
         ScaleWidth      =   4155
         TabIndex        =   176
         Top             =   240
         Visible         =   0   'False
         Width           =   4215
         Begin VB.Label lblDefInt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   1560
            TabIndex        =   202
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "พลังป้องกันทางเวทย์ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   120
            TabIndex        =   201
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblStrLHand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1560
            TabIndex        =   200
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "พลังโจมตีมือรอง :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   199
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblExpShow 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Exp / Exp Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1200
            TabIndex        =   177
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label lblStr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1560
            TabIndex        =   175
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label lblStrg 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "พลังโจมตีกายภาพ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   174
            Top             =   720
            Width           =   1335
         End
         Begin VB.Image imgEXPBar2 
            Height          =   240
            Left            =   1200
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label lblExpNow 
            BackStyle       =   0  'Transparent
            Caption         =   "Exp ปัจจุบัน :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   120
            TabIndex        =   173
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblHeader 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ข้อมูลส่วนตัว"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   255
            Left            =   1440
            TabIndex        =   172
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3960
            TabIndex        =   171
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "พลังโจมตีทางเวทย์ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   120
            TabIndex        =   170
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "อัตราโป๊ะเชะ (คริติคอล) % :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   0
            TabIndex        =   169
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "อัตราหลบหลีก % :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   255
            Left            =   120
            TabIndex        =   168
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   120
            TabIndex        =   167
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label lblInt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   1560
            TabIndex        =   166
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label lblCrit 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2040
            TabIndex        =   165
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblDodge 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   255
            Left            =   1440
            TabIndex        =   164
            Top             =   3120
            Width           =   2655
         End
         Begin VB.Label lblBlock 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   1560
            TabIndex        =   163
            Top             =   2160
            Width           =   2415
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "ความเร็วในการเดิน : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C0C0&
            Height          =   255
            Left            =   120
            TabIndex        =   162
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label lblWalk 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   1800
            TabIndex        =   161
            Top             =   2400
            Width           =   1935
         End
         Begin VB.Label lblAttackspeed 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   1800
            TabIndex        =   160
            Top             =   2640
            Width           =   1935
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "ความเร็วการโจมตี : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "พลังโจมตีระยะไกล :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label lblLongAttack 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1560
            TabIndex        =   157
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ความแรงโป๊ะ (คริติคอล) % :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "โจมตีเจาะเกราะ (%) :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "โอกาศทำให้มึน (%) :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ลดเวลาร่ายเวทย์ (%) :"
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
            Height          =   255
            Left            =   120
            TabIndex        =   153
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label lblCritATK 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2040
            TabIndex        =   152
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label lblNDEF 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   1560
            TabIndex        =   151
            Top             =   3360
            Width           =   2415
         End
         Begin VB.Label lblKick 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   1560
            TabIndex        =   150
            Top             =   3600
            Width           =   2415
         End
         Begin VB.Label lblCastTime 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
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
            Height          =   255
            Left            =   1680
            TabIndex        =   149
            Top             =   3840
            Width           =   2175
         End
         Begin VB.Label lblVampire 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   1800
            TabIndex        =   148
            Top             =   4080
            Width           =   1935
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "ดูดเลือดเมื่อโจมตี (%) : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   147
            Top             =   4080
            Width           =   1575
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Regen HP :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   120
            TabIndex        =   146
            Top             =   4320
            Width           =   1455
         End
         Begin VB.Label lblRegenHP 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   1560
            TabIndex        =   145
            Top             =   4320
            Width           =   2415
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Regen MP :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   4560
            Width           =   1455
         End
         Begin VB.Label lblRegenMP 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   1560
            TabIndex        =   143
            Top             =   4560
            Width           =   2415
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "พลังป้องกัน :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   4800
            Width           =   1455
         End
         Begin VB.Label lblDEF 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "xxx / xxx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   255
            Left            =   1560
            TabIndex        =   141
            Top             =   4800
            Width           =   2415
         End
      End
      Begin MSWinsockLib.Winsock Socket 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblMapLoad 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "กำลังโหลดแผนที่..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   204
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Label lblChatter 
      BackColor       =   &H000000FF&
      Caption         =   "กด Enter เพื่อแชท"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   120
      TabIndex        =   88
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label lblEXP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4320
      TabIndex        =   87
      Top             =   7140
      Width           =   2865
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4320
      TabIndex        =   86
      Top             =   6840
      Width           =   2865
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4320
      TabIndex        =   85
      Top             =   6510
      Width           =   2865
   End
   Begin VB.Image imgHPBar 
      Height          =   240
      Left            =   4320
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Image imgMPBar 
      Height          =   240
      Left            =   4320
      Top             =   6810
      Width           =   2895
   End
   Begin VB.Image imgEXPBar 
      Height          =   240
      Left            =   4320
      Top             =   7140
      Width           =   2895
   End
   Begin VB.Label lblPing 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local ----------------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   6360
      TabIndex        =   17
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblGold 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0g"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8280
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   2265
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' Cursor
'Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
'Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
'Private oldCursor As Long
'Private newCursor As Long

' Private Const GCL_HCURSOR = (-12)

' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long

Private Sub cmdAAnim_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditAnimation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAnim_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAPet_Click()

' Edit Pet

End Sub

Private Sub cmdFps_Click()
    Fps_Max = Not Fps_Max
    
    If Fps_Max = True Then
        cmdFps.Caption = "คอมสเป็คต่ำ"
    Else
        cmdFps.Caption = "คอมสเป็คสูง"
    End If
    
End Sub

Private Sub cmdLevel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestLevelUp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' render the map temp
    ScreenshotMap
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub Command1_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        AddText "You need to be a high enough staff member to do this!", AlertColor
        Exit Sub
    End If
End Sub

Private Sub cmdWho_Click()
    ShellExecute 0, "open", "https://www.tmtopup.com/topup/?uid=139654", vbNullString, vbNullString, 1
    'SendWhosOnline
    'frmOnline.Visible = True
    'frmOnline.Show
End Sub

Private Sub Command2_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If
    
    SendRequestEditdoors
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Command3_Click()
If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
    Exit Sub
End If

    ' Send Quest Edit
    SendRequestEditQuest

End Sub

Private Sub Command4_Click()
    frmDialog.Show
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' move GUI
    'picAdmin.Left = 544
    'picCurrency.Left = txtChat.Left
    'picCurrency.top = txtChat.top
    'picDialogue.top = txtChat.top
    'picDialogue.Left = txtChat.Left
    'picCover.top = picScreen.top - 1
    'picCover.Left = picScreen.Left - 1
    'picCover.height = picScreen.height + 2
    'picCover.width = picScreen.width + 2
    
    BFPS = True
    ' Not lock it work more than
    ' FPS_Lock = True
    cbMAP.ListIndex = 0
    SetFocusOnGame
    
    'load a new cursor and assign to the control you wish
    'to have the custom cursor
    'change the path to your cursor icon
       
       ' cursor ภาพเมาส์
   '    newCursor = LoadCursorFromFile("data files\graphics\cursor\Default.ani")
       
   '    If newCursor > 0 Then
   '         oldCursor = SetClassLong(frmMain.hwnd, GCL_HCURSOR, newCursor)
   '         oldCursor = SetClassLong(picScreen.hwnd, GCL_HCURSOR, newCursor)
   '    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Cancel = True
    logoutGame
    
    ' ย้อนกลับเมาส์เป็นปกติ
    
  '  If oldCursor > 0 Then
  '      Call SetClassLong(frmMain.hwnd, GCL_HCURSOR, oldCursor)
  '      Call SetClassLong(picScreen.hwnd, GCL_HCURSOR, oldCursor)
  '      DestroyCursor newCursor
    '  End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' ย้อนกลับเมาส์เป็นปกติ
    
 '   If oldCursor > 0 Then
 '       Call SetClassLong(frmMain.hwnd, GCL_HCURSOR, oldCursor)
 '       Call SetClassLong(picScreen.hwnd, GCL_HCURSOR, oldCursor)
 '       DestroyCursor newCursor
'    End If

    ' cursor ภาพเมาส์
  '     newCursor = LoadCursorFromFile("data files\graphics\cursor\Default.ani")
       
   '    If newCursor > 0 Then
   '         oldCursor = SetClassLong(frmMain.hwnd, GCL_HCURSOR, newCursor)
   '         oldCursor = SetClassLong(picScreen.hwnd, GCL_HCURSOR, newCursor)
   '    End If

    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Image1_Click()

End Sub

Private Sub imgAcceptTrade_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    AcceptTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgAcceptTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_Click(Index As Integer)
Dim Buffer As clsBuffer
Dim i As Long
imgButton(7).Visible = True
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            If Not picInventory.Visible Then
                ' show the window
                picInventory.Visible = True
                picCharacter.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = False
                picPet.Visible = False
                BltInventory
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            Else
                picInventory.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            End If
        Case 2
            If Not picSpells.Visible Then
                ' send packet
                Set Buffer = New clsBuffer
                Buffer.WriteLong CSpells
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                ' show the window
                picSpells.Visible = True
                picInventory.Visible = False
                picCharacter.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = False
                picPet.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            Else
                picSpells.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            End If
        Case 3
            If Not picCharacter.Visible Then
                ' send packet
                SendRequestPlayerData
                ' show the window
                picCharacter.Visible = True
                picInventory.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = False
                picPet.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                ' Render
                BltEquipment
                BltFace
                SetFocusOnGame
            Else
                picCharacter.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            End If
        Case 4
            If Not picOptions.Visible Then
                ' show the window
                picCharacter.Visible = False
                picInventory.Visible = False
                picSpells.Visible = False
                picOptions.Visible = True
                picParty.Visible = False
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = False
                picPet.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            Else
                picOptions.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            End If
        Case 5
            If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                SendTradeRequest
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            Else
                AddText "ไม่พบเป้าหมายที่ต้องการค้าขายด้วย.", BrightRed
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            End If
        Case 6
            ' show the window
            If Not picParty.Visible Then
                picCharacter.Visible = False
                picInventory.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = True
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = False
                picPet.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            Else
                picParty.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            End If
        Case 7
            If Not picQuestLog.Visible Then
                UpdateQuestLog
                ' show the window
                picCharacter.Visible = False
                picInventory.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = True
                picPet.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            Else
                picQuestLog.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            End If
        Case 8
            Call AddText("ยังไม่เปิดให้บริการในส่วนนี้ค่ะ.", BrightRed)
            Exit Sub
        
            If Not picPet.Visible Then
                ' show the window
                picCharacter.Visible = False
                picInventory.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = False
                picPet.Visible = True
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            Else
                picPet.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
            End If
        Case 9
        
Dim Speed As Long, Def As Long, Mdef As Long, RegenHp As Long, RegenMp As Long
Dim ATKL, numWeapon As Long, CT As Double
Dim Dodge As Long, Reflect As Long
Dim LeftHand As Boolean
    ' แสดงข้อมูลส่วนตัว
    picMe.Visible = True

    Def = 0
    Mdef = 0
    RegenHp = 0
    RegenMp = 0
    Dodge = 0
    Reflect = 0
    LeftHand = False
    
    ' lblStrLHand
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        If Item(GetPlayerEquipment(MyIndex, Shield)).LHand > 0 Then
            lblStrLHand.Caption = ((((Item(GetPlayerEquipment(MyIndex, Shield)).Data2) + (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Shield)).DmgLow) / 100) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Shield)).Data2) + (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Shield)).DmgHigh) / 100)
            LeftHand = True
        Else
            lblStrLHand.Caption = "0 - 0"
        End If
    Else
        lblStrLHand.Caption = "0 - 0"
    End If
    
    ' lblCrit
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            If (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate > 0 And (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate < 1 Then
                lblCrit.Caption = "0" & (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate & " %"
            Else
                If (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate < 80 Then
                    lblCrit.Caption = (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate & " %"
                Else
                    lblCrit.Caption = " 80 %"
                End If
            End If
        Else
            If (GetPlayerStat(MyIndex, willpower) / 2.5) > 0 And (GetPlayerStat(MyIndex, willpower) / 2.5) < 1 Then
                lblCrit.Caption = "0" & (GetPlayerStat(MyIndex, willpower) / 2.5) & " %"
            Else
                If (GetPlayerStat(MyIndex, willpower) / 2.5) < 80 Then
                    lblCrit.Caption = (GetPlayerStat(MyIndex, willpower) / 2.5) & " %"
                Else
                    lblCrit.Caption = "80 %"
                End If
            End If
        End If
        
        If LeftHand = True Then
            If (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Shield)).CritRate < 80 Then
                lblCrit.Caption = lblCrit.Caption & " / " & (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Shield)).CritRate & " % (ขวา/ซ้าย)"
            Else
                lblCrit.Caption = lblCrit.Caption & " /  80 % (ขวา/ซ้าย)"
            End If
        End If
    
    ' Dodge & Reflect
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Weapon)).Dodge
        Reflect = Reflect + Item(GetPlayerEquipment(MyIndex, Weapon)).Reflect
    End If
    
    If GetPlayerEquipment(MyIndex, Armor) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Armor)).Dodge
        Reflect = Reflect + Item(GetPlayerEquipment(MyIndex, Armor)).Reflect
    End If
    
    If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Helmet)).Dodge
        Reflect = Reflect + Item(GetPlayerEquipment(MyIndex, Helmet)).Reflect
    End If
    
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Shield)).Dodge
        Reflect = Reflect + Item(GetPlayerEquipment(MyIndex, Shield)).Reflect
    End If
    
    ' Mdef
    If GetPlayerEquipment(MyIndex, Armor) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Armor)).MATK
    End If
    
    If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Helmet)).MATK
    End If
    
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Shield)).MATK
    End If
    
    Mdef = Mdef + (GetPlayerLevel(MyIndex) * 2)
    
    ' lblMdef
    lblDefInt.Caption = GetPlayerStat(MyIndex, Intelligence) + Mdef & " -  " & (GetPlayerStat(MyIndex, Intelligence) * 2) + Mdef
    
    ' lblDodge
        If Dodge + (GetPlayerStat(MyIndex, Agility) / 2.5) > 0 And Dodge + (GetPlayerStat(MyIndex, Agility) / 2.5) < 1 Then
            lblDodge.Caption = "0" & Dodge + (GetPlayerStat(MyIndex, Agility) / 4) & " %"
        Else
            If Dodge + (GetPlayerStat(MyIndex, Agility) / 4) < 80 Then
                lblDodge.Caption = Dodge + (GetPlayerStat(MyIndex, Agility) / 4) & " %"
            Else
                lblDodge.Caption = "80 %"
            End If
        End If
        
    ' lblReflect
        If Reflect + (GetPlayerStat(MyIndex, Endurance) / 10) > 0 And Reflect + (GetPlayerStat(MyIndex, Endurance) / 10) < 1 Then
            lblBlock.Caption = Reflect + (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
        Else
            If Reflect + (GetPlayerStat(MyIndex, Endurance) / 10) < 80 Then
                lblBlock.Caption = Reflect + (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
            Else
                lblBlock.Caption = "80 %"
            End If
        End If

    ' Fixed bug <For now>
    lblStr.Caption = (((GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) / 2) & " - " & (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)
    frmMain.lblLongAttack.Caption = "0 - 0"
    
    ' lblInt
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        lblInt.Caption = (((Item(GetPlayerEquipment(MyIndex, Weapon)).MATK + (GetPlayerStat(MyIndex, Intelligence) * 4) + (GetPlayerLevel(MyIndex) * 2))) * Item(GetPlayerEquipment(MyIndex, Weapon)).MagicLow) / 100 & " - " & ((Item(GetPlayerEquipment(MyIndex, Weapon)).MATK + (GetPlayerStat(MyIndex, Intelligence) * 4) + (GetPlayerLevel(MyIndex) * 2))) * (Item(GetPlayerEquipment(MyIndex, Weapon)).MagicHigh / 100)
    Else
        lblInt.Caption = ((GetPlayerLevel(MyIndex) * 2) + (GetPlayerStat(MyIndex, Intelligence) * 4)) / 2 & " - " & (GetPlayerLevel(MyIndex) * 2) + (GetPlayerStat(MyIndex, Intelligence) * 4)
    End If
    
    ATKL = (GetPlayerStat(MyIndex, willpower) * 3.5) + (GetPlayerLevel(MyIndex) * 2)

    ' Fixed lblStr And lblLong
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        lblStr.Caption = ((((Item(GetPlayerEquipment(MyIndex, Weapon)).Data2) + (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgLow) / 100) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Weapon)).Data2) + (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgHigh) / 100)
        ' Fixed ProjecTile Damage
        If (Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic) > 0 Then
            frmMain.lblLongAttack.Caption = ((((Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Damage + ATKL) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgLow) / 100)) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Damage + ATKL) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgHigh) / 100))
        Else
            frmMain.lblLongAttack.Caption = "0 - 0"
        End If
    Else
        lblStr.Caption = (((GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) / 2) & " - " & (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)
        frmMain.lblLongAttack.Caption = "0 - 0"
    End If
    
    ' lblCritATK
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            If (Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic) > 0 Then
                lblCritATK.Caption = (120 + (GetPlayerStat(MyIndex, willpower))) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritATK & " %"
            Else
                lblCritATK.Caption = (120 + (GetPlayerStat(MyIndex, willpower))) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritATK & " %"
            End If
        Else ' ถ้าไม่ใส่อาวุธ
            lblCritATK.Caption = 120 + (GetPlayerStat(MyIndex, willpower))
        End If
        
        If LeftHand = True Then
            lblCritATK.Caption = lblCritATK.Caption & " / " & (120 + (GetPlayerStat(MyIndex, willpower))) + Item(GetPlayerEquipment(MyIndex, Shield)).CritATK & " % (ขวา/ซ้าย)"
        End If
        
    ' lblWalk
    If Int(WALK_SPEED + Int(GetPlayerStat(MyIndex, Stats.Agility) / 50)) < 6 Then
        lblWalk.Caption = Int(WALK_SPEED + Int(GetPlayerStat(MyIndex, Stats.Agility) / 50)) & " หน่วย."
    Else
        lblWalk.Caption = "6 หน่วย (สูงสุด)."
    End If
    
    ' lblAttackspeed
    ' ความเร็วในการโจมตี
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            Speed = ((2000 + Item(GetPlayerEquipment(MyIndex, Weapon)).Speed) - Item(GetPlayerEquipment(MyIndex, Weapon)).SpeedLow) - ((GetPlayerStat(MyIndex, Stats.Agility) * 5))
            If Speed > 1000 Then
                lblAttackspeed.Caption = FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            Else
                If Speed > 200 Then
                    lblAttackspeed.Caption = "0" & FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
                Else
                    lblAttackspeed.Caption = "0.100 ครั้ง/วินาที"
                End If
            End If
        Else
            Speed = 2000 - ((GetPlayerStat(MyIndex, Stats.Agility) * 5))
            If Speed > 1000 Then
                lblAttackspeed.Caption = FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            Else
                lblAttackspeed.Caption = "0" & FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            End If
        End If
        
        ' lblNDEF
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblNDEF.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).NDEF & " %"
        Else
            lblNDEF.Caption = "0 %"
        End If
        
        ' lblKick
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblKick.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).Kick & " %"
        Else
            lblKick.Caption = "0 %"
        End If
        
        ' lblCastTime ร่ายเวทย์ v2
        
        Select Case (1 + (GetPlayerStat(MyIndex, willpower) / 50))
            Case 7: CT = 100
            Case 6: CT = 50
            Case 5: CT = 33
            Case 4: CT = 100 - 25
            Case 3: CT = 100 - 33
            Case 2: CT = 100 - 50
            Case 1: CT = 0
            Case Else: CT = 0
        End Select
        
        CT = 100 - (100 / (1 + (GetPlayerStat(MyIndex, willpower) / 50)))
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblCastTime.Caption = (100 - (Item(GetPlayerEquipment(MyIndex, Weapon)).DelayDown * 100)) & " + " & FormatNumber(CT, 2) & " %"
        Else
            lblCastTime.Caption = FormatNumber(CT, 2) & " %"
        End If
                
        ' lblVampire
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblVampire.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).Vampire & " %"
        Else
            lblVampire.Caption = "0 %"
        End If
    
        ' lblRegenHP
        
            RegenHp = (GetPlayerStat(MyIndex, Stats.Endurance) * 2) + (GetPlayerLevel(MyIndex) / 2) + (GetPlayerMaxVital(MyIndex, HP) * 0.01) + 2
        
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Weapon)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Weapon)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Armor) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Armor)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Armor)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Helmet)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Helmet)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Shield) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Shield)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Shield)).RegenHp)
                End If
            End If
    
            lblRegenHP.Caption = RegenHp & " หน่วย."
    
        ' lblRegenMP
        
            RegenMp = (GetPlayerStat(MyIndex, Stats.Intelligence)) + (GetPlayerLevel(MyIndex) / 4) + (GetPlayerMaxVital(MyIndex, MP) * 0.01) + 1
            
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Weapon)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Weapon)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Armor) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Armor)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Armor)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Helmet)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Helmet)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Shield) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Shield)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Shield)).RegenMp)
                End If
            End If
    
            lblRegenMP.Caption = RegenMp & " หน่วย."
    
        ' lblDEF
        
        If GetPlayerEquipment(MyIndex, Armor) > 0 Then
            Def = Def + Item(GetPlayerEquipment(MyIndex, Armor)).Data2
        End If
    
        If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
            Def = Def + Item(GetPlayerEquipment(MyIndex, Helmet)).Data2
        End If
    
        Def = Def + (GetPlayerStat(MyIndex, Endurance) * 2) + (GetPlayerLevel(MyIndex) * 2)
        
        ' Check berserker
        If GetPlayerClass(MyIndex) = 4 Then ' Berserker Class None def
            Def = 0
        End If
    
        lblDEF.Caption = Def & " หน่วย."

    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Main Index
    
    ' change the button we're hovering on
    If Not MainButton(Index).state = 2 Then ' make sure we're not clicking
        changeButtonState_Main Index, 1 ' hover
    End If
    
    ' play sound
    If Not LastButtonSound_Main = Index Then
        PlaySound Sound_ButtonHover
        LastButtonSound_Main = Index
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    ' reset all buttons
    resetButtons_Main -1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Main Index
    
    ' change the button we're hovering on
    changeButtonState_Main Index, 2 ' clicked
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub Label10_Click()
'ShellExecute 0, "open", "http://monsterwaronline.blogspot.com/", vbNullString, vbNullString, 1
ShellExecute 0, "open", "https://www.facebook.com/pages/Monster-War-Online/139536579518271?ref=hl/", vbNullString, vbNullString, 1
End Sub

Private Sub Label11_Click()
picMe.Visible = False
End Sub

Private Sub Label19_Click()
frmDialog.Show
End Sub

Private Sub Label2_Click()
PetAttack MyIndex
End Sub

Private Sub Label3_Click()
PetFollow MyIndex
End Sub

Private Sub Label4_Click()
PetWander MyIndex
End Sub

Private Sub Label5_Click()
PetDisband MyIndex
End Sub

Private Sub lblBefore2_Click()
    picParty2.Visible = False
    picParty.Visible = True
    picParty3.Visible = False
End Sub

Private Sub lblBefore3_Click()
    picParty2.Visible = True
    picParty.Visible = False
    picParty3.Visible = False
End Sub

Private Sub lblChoices_Click(Index As Integer)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CEventChatReply
Buffer.WriteLong EventReplyID
Buffer.WriteLong EventReplyPage
Buffer.WriteLong Index
SendData Buffer.ToArray
Set Buffer = Nothing
ClearEventChat
InEvent = False
End Sub

Private Sub lblCurrencyCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picCurrency.Visible = False
    txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCurrencyCancel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgDeclineTrade_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DeclineTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgDeclineTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgLeaveShop_Click()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CCloseShop
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
    
    picCover.Visible = False
    picShop.Visible = False
    InShop = 0
    ShopAction = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgLeaveShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCurrencyOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Currency Menu Overflow Fix
   If IsNumeric(txtCurrency.text) Then
        If CurrencyMenu = 3 Then
            If Val(txtCurrency.text) > GetBankItemValue(MyIndex, tmpCurrencyItem) Then txtCurrency.text = GetBankItemValue(MyIndex, tmpCurrencyItem)
        ElseIf Val(txtCurrency.text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then txtCurrency.text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
        End If
        
        Select Case CurrencyMenu
            Case 1 ' drop item
                SendDropItem tmpCurrencyItem, Val(txtCurrency.text)
            Case 2 ' deposit item
                DepositItem tmpCurrencyItem, Val(txtCurrency.text)
            Case 3 ' withdraw item
                WithdrawItem tmpCurrencyItem, Val(txtCurrency.text)
            Case 4 ' offer trade item
                TradeItem tmpCurrencyItem, Val(txtCurrency.text)
        End Select
    Else
        AddText "โปรดใส่จำนวนที่ต้องการแลกเปลี่ยน.", BrightRed
        Exit Sub
    End If


    picCurrency.Visible = False
    tmpCurrencyItem = 0
    txtCurrency.text = vbNullString
    CurrencyMenu = 0 ' clear

' Error handler
Exit Sub
errorhandler:
HandleError "lblCurrencyOk_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub imgShopBuy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ShopAction = 1 Then Exit Sub
    ShopAction = 1 ' buying an item
    AddText "คลิกไอเทมบนร้านค้าเพื่อซื้อ.", White
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgShopBuy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgShopSell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ShopAction = 2 Then Exit Sub
    ShopAction = 2 ' selling an item
    AddText "ดับเบิ้ลคลิกไอเทมในช่องเก็บของเพื่อขาย.", White
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgShopSell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblDialogue_Button_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' call the handler
    dialogueHandler Index
    
    picDialogue.Visible = False
    dialogueIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblDialogue_Button_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblEventChatContinue_Click()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CEventChatReply
Buffer.WriteLong EventReplyID
Buffer.WriteLong EventReplyPage
Buffer.WriteLong 0
SendData Buffer.ToArray
Set Buffer = Nothing
ClearEventChat
InEvent = False
End Sub

Private Sub lblMe_Click()
Dim Speed As Long, Def As Long, Mdef As Long, RegenHp As Long, RegenMp As Long
Dim ATKL, numWeapon As Long, CT As Long
Dim Dodge As Long
    ' แสดงข้อมูลส่วนตัว
    picMe.Visible = True

    Def = 0
    Mdef = 0
    RegenHp = 0
    RegenMp = 0
    Dodge = 0
    
    ' lblCrit
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            If (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate > 0 And (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate < 1 Then
                lblCrit.Caption = "0" & (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate & " %"
            Else
                lblCrit.Caption = (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate & " %"
            End If
        Else
            If (GetPlayerStat(MyIndex, willpower) / 2.5) > 0 And (GetPlayerStat(MyIndex, willpower) / 2.5) < 1 Then
                lblCrit.Caption = "0" & (GetPlayerStat(MyIndex, willpower) / 2.5) & " %"
            Else
                lblCrit.Caption = (GetPlayerStat(MyIndex, willpower) / 2.5) & " %"
            End If
        End If
    
    ' Dodge
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Weapon)).Dodge
    End If
    
    If GetPlayerEquipment(MyIndex, Armor) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Armor)).Dodge
    End If
    
    If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Helmet)).Dodge
    End If
    
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Shield)).Dodge
    End If
    
    ' lblStrLHand
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        If Item(GetPlayerEquipment(MyIndex, Shield)).LHand > 0 Then
            lblStrLHand.Caption = ((((Item(GetPlayerEquipment(MyIndex, Shield)).Data2) + (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Shield)).DmgLow) / 100) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Shield)).Data2) + (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Shield)).DmgHigh) / 100)
        Else
            lblStrLHand.Caption = "0 - 0"
        End If
    Else
        lblStrLHand.Caption = "0 - 0"
    End If
    
    ' Mdef
    If GetPlayerEquipment(MyIndex, Armor) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Armor)).MATK
    End If
    
    If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Helmet)).MATK
    End If
    
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Shield)).MATK
    End If
    
    Mdef = Mdef + (GetPlayerLevel(MyIndex) * 2)
    
    ' lblMdef
    lblDefInt.Caption = GetPlayerStat(MyIndex, Intelligence) + Mdef & " -  " & (GetPlayerStat(MyIndex, Intelligence) * 2) + Mdef
    
    ' lblDodge
        If Dodge + (GetPlayerStat(MyIndex, Agility) / 2.5) > 0 And Dodge + (GetPlayerStat(MyIndex, Agility) / 2.5) < 1 Then
            lblDodge.Caption = "0" & Dodge + (GetPlayerStat(MyIndex, Agility) / 4) & " %"
        Else
            lblDodge.Caption = Dodge + (GetPlayerStat(MyIndex, Agility) / 4) & " %"
        End If
        
    ' lblBlock
        If GetPlayerEquipment(MyIndex, Shield) > 0 Then
            lblBlock.Caption = Item(GetPlayerEquipment(MyIndex, Shield)).Data2 + (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
        Else
            If (GetPlayerStat(MyIndex, Endurance) / 10) > 0 And (GetPlayerStat(MyIndex, Endurance) / 10) < 1 Then
                lblBlock.Caption = "0" & (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
            Else
                lblBlock.Caption = (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
            End If
        End If

    ' Fixed bug <For now>
    lblStr.Caption = (((GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) / 2) & " - " & (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)
    frmMain.lblLongAttack.Caption = "0 - 0"
    
    ' lblInt
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        lblInt.Caption = (((Item(GetPlayerEquipment(MyIndex, Weapon)).MATK + (GetPlayerStat(MyIndex, Intelligence) * 2) + (GetPlayerLevel(MyIndex) * 2))) * Item(GetPlayerEquipment(MyIndex, Weapon)).MagicLow) / 100 & " - " & ((Item(GetPlayerEquipment(MyIndex, Weapon)).MATK + (GetPlayerStat(MyIndex, Intelligence) * 2) + (GetPlayerLevel(MyIndex) * 2))) * (Item(GetPlayerEquipment(MyIndex, Weapon)).MagicHigh / 100)
    Else
        lblInt.Caption = ((GetPlayerLevel(MyIndex) * 2) + (GetPlayerStat(MyIndex, Intelligence) * 2)) / 2 & " - " & (GetPlayerLevel(MyIndex) * 2) + (GetPlayerStat(MyIndex, Intelligence) * 2)
    End If
    
    ATKL = (GetPlayerStat(MyIndex, willpower) * 2) + (GetPlayerLevel(MyIndex) * 2)

    ' Fixed lblStr And lblLong
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        lblStr.Caption = ((((Item(GetPlayerEquipment(MyIndex, Weapon)).Data2) + (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgLow) / 100) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Weapon)).Data2) + (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgHigh) / 100)
        ' Fixed ProjecTile Damage
        If (Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic) > 0 Then
            frmMain.lblLongAttack.Caption = ((((Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Damage + ATKL) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgLow) / 100)) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Damage + ATKL) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgHigh) / 100))
        Else
            frmMain.lblLongAttack.Caption = "0 - 0"
        End If
    Else
        lblStr.Caption = (((GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) / 2) & " - " & (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)
        frmMain.lblLongAttack.Caption = "0 - 0"
    End If
    
    ' lblCritATK
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            If (Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic) > 0 Then
                lblCritATK.Caption = (120 + (GetPlayerStat(MyIndex, willpower))) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritATK & " %"
            Else
                lblCritATK.Caption = (120 + (GetPlayerStat(MyIndex, willpower))) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritATK & " %"
            End If
        Else ' ถ้าไม่ใส่อาวุธ
            lblCritATK.Caption = 120 + (GetPlayerStat(MyIndex, willpower))
        End If
        
    ' lblWalk
    If Int(WALK_SPEED + Int(GetPlayerStat(MyIndex, Stats.Agility) / 50)) < 6 Then
        lblWalk.Caption = Int(WALK_SPEED + Int(GetPlayerStat(MyIndex, Stats.Agility) / 50)) & " หน่วย."
    Else
        lblWalk.Caption = "6 หน่วย (สูงสุด)."
    End If
    
    ' lblAttackspeed
    ' ความเร็วในการโจมตี
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            Speed = ((2000 + Item(GetPlayerEquipment(MyIndex, Weapon)).Speed) - Item(GetPlayerEquipment(MyIndex, Weapon)).SpeedLow) - ((GetPlayerStat(MyIndex, Stats.Agility) * 5))
            If Speed > 1000 Then
                lblAttackspeed.Caption = FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            Else
                If Speed > 200 Then
                    lblAttackspeed.Caption = "0" & FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
                Else
                    lblAttackspeed.Caption = "0.100 ครั้ง/วินาที"
                End If
            End If
        Else
            Speed = 2000 - ((GetPlayerStat(MyIndex, Stats.Agility) * 5))
            If Speed > 1000 Then
                lblAttackspeed.Caption = FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            Else
                lblAttackspeed.Caption = "0" & FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            End If
        End If
        
        ' lblNDEF
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblNDEF.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).NDEF & " %"
        Else
            lblNDEF.Caption = "0 %"
        End If
        
        ' lblKick
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblKick.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).Kick & " %"
        Else
            lblKick.Caption = "0 %"
        End If
        
        ' lblCastTime ร่ายเวทย์ v2
        
        Select Case (1 + (GetPlayerStat(MyIndex, willpower) / 50))
            Case 1: CT = 100
            Case 2: CT = 50
            Case 3: CT = 33
            Case 4: CT = 25
            Case 5: CT = 20
            Case 6: CT = 16
            Case Else: CT = 0
        End Select
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblCastTime.Caption = (100 - (Item(GetPlayerEquipment(MyIndex, Weapon)).DelayDown * 100)) & " + " & (100 - CT) & " %"
        Else
            lblCastTime.Caption = (100 - CT) & " %"
        End If
                
        ' lblVampire
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblVampire.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).Vampire & " %"
        Else
            lblVampire.Caption = "0 %"
        End If
    
        ' lblRegenHP
        
            RegenHp = (GetPlayerStat(MyIndex, Stats.Endurance) * 2) + (GetPlayerLevel(MyIndex) / 2) + (GetPlayerMaxVital(MyIndex, HP) * 0.01) + 2
        
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Weapon)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Weapon)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Armor) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Armor)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Armor)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Helmet)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Helmet)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Shield) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Shield)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Shield)).RegenHp)
                End If
            End If
    
            lblRegenHP.Caption = RegenHp & " หน่วย."
    
        ' lblRegenMP
        
            RegenMp = (GetPlayerStat(MyIndex, Stats.Intelligence)) + (GetPlayerLevel(MyIndex) / 4) + (GetPlayerMaxVital(MyIndex, MP) * 0.01) + 1
            
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Weapon)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Weapon)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Armor) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Armor)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Armor)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Helmet)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Helmet)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Shield) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Shield)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Shield)).RegenMp)
                End If
            End If
    
            lblRegenMP.Caption = RegenMp & " หน่วย."
    
        ' lblDEF
        
        If GetPlayerEquipment(MyIndex, Armor) > 0 Then
            Def = Def + Item(GetPlayerEquipment(MyIndex, Armor)).Data2
        End If
    
        If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
            Def = Def + Item(GetPlayerEquipment(MyIndex, Helmet)).Data2
        End If
    
        Def = Def + (GetPlayerStat(MyIndex, Endurance) * 2) + (GetPlayerLevel(MyIndex) * 2)
        
        ' Check berserker
        If GetPlayerClass(MyIndex) = 4 Then ' Berserker Class None def
            Def = 0
        End If
    
        lblDEF.Caption = Def & " หน่วย."
    
End Sub

Private Sub lblNext1_Click()
    picParty2.Visible = True
    picParty.Visible = False
    picParty3.Visible = False
End Sub

Private Sub lblNext2_Click()
    picParty2.Visible = False
    picParty.Visible = False
    picParty3.Visible = True
End Sub

Private Sub lblPartyInvite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
        SendPartyRequest
    Else
        AddText "ไม่พบเป้าหมายที่ต้องการเชิญเข้าปาร์ตี้.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblPartyLeave_Click()
        ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If party.Leader > 0 Then
        If party.MemberCount > 2 Then
            Call SendPartyChatMsg("มอบหัวหน้าปาร์ตี้ให้กับ " & frmMain.lblPartyMember(2).Caption)
        End If
        
        SendPartyLeave
    Else
        AddText "คุณไม่มีปาร์ตี้.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblPetAttack_Click()
Call PetAttack(MyIndex)
End Sub

Private Sub lblPetDisband_Click()
 Call PetDisband(MyIndex)
End Sub

Private Sub lblPetFollow_Click()
 Call PetFollow(MyIndex)
End Sub

Private Sub lblPetWander_Click()
Call PetWander(MyIndex)
End Sub


Private Sub lblTrainStat_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
    SendTrainStat Index
    
    Dim Speed As Long, Def As Long, Mdef As Long, RegenHp As Long, RegenMp As Long
    Dim ATKL, numWeapon As Long, CT As Double
    Dim Dodge As Long

    Def = 0
    Mdef = 0
    RegenHp = 0
    RegenMp = 0
    Dodge = 0
    
    ' lblCrit
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            If (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate > 0 And (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate < 1 Then
                lblCrit.Caption = "0" & (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate & " %"
            Else
                lblCrit.Caption = (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate & " %"
            End If
        Else
            If (GetPlayerStat(MyIndex, willpower) / 2.5) > 0 And (GetPlayerStat(MyIndex, willpower) / 2.5) < 1 Then
                lblCrit.Caption = "0" & (GetPlayerStat(MyIndex, willpower) / 2.5) & " %"
            Else
                lblCrit.Caption = (GetPlayerStat(MyIndex, willpower) / 2.5) & " %"
            End If
        End If
    
    ' Dodge
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Weapon)).Dodge
    End If
    
    If GetPlayerEquipment(MyIndex, Armor) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Armor)).Dodge
    End If
    
    If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Helmet)).Dodge
    End If
    
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Shield)).Dodge
    End If
    
    ' lblStrLHand
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        If Item(GetPlayerEquipment(MyIndex, Shield)).LHand > 0 Then
            lblStrLHand.Caption = ((((Item(GetPlayerEquipment(MyIndex, Shield)).Data2) + (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Shield)).DmgLow) / 100) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Shield)).Data2) + (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Shield)).DmgHigh) / 100)
        Else
            lblStrLHand.Caption = "0 - 0"
        End If
    Else
        lblStrLHand.Caption = "0 - 0"
    End If
    
    ' Mdef
    If GetPlayerEquipment(MyIndex, Armor) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Armor)).MATK
    End If
    
    If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Helmet)).MATK
    End If
    
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Shield)).MATK
    End If
    
    Mdef = Mdef + (GetPlayerLevel(MyIndex) * 2)
    
    ' lblMdef
    lblDefInt.Caption = GetPlayerStat(MyIndex, Intelligence) + Mdef & " -  " & (GetPlayerStat(MyIndex, Intelligence) * 2) + Mdef
    
    ' lblDodge
        If Dodge + (GetPlayerStat(MyIndex, Agility) / 2.5) > 0 And Dodge + (GetPlayerStat(MyIndex, Agility) / 2.5) < 1 Then
            lblDodge.Caption = "0" & Dodge + (GetPlayerStat(MyIndex, Agility) / 4) & " %"
        Else
            lblDodge.Caption = Dodge + (GetPlayerStat(MyIndex, Agility) / 4) & " %"
        End If
        
    ' lblBlock
        If GetPlayerEquipment(MyIndex, Shield) > 0 Then
            lblBlock.Caption = Item(GetPlayerEquipment(MyIndex, Shield)).Data2 + (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
        Else
            If (GetPlayerStat(MyIndex, Endurance) / 10) > 0 And (GetPlayerStat(MyIndex, Endurance) / 10) < 1 Then
                lblBlock.Caption = "0" & (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
            Else
                lblBlock.Caption = (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
            End If
        End If

    ' Fixed bug <For now>
    lblStr.Caption = (((GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) / 2) & " - " & (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)
    frmMain.lblLongAttack.Caption = "0 - 0"
    
    ' lblInt
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        lblInt.Caption = (((Item(GetPlayerEquipment(MyIndex, Weapon)).MATK + (GetPlayerStat(MyIndex, Intelligence) * 4) + (GetPlayerLevel(MyIndex) * 2))) * Item(GetPlayerEquipment(MyIndex, Weapon)).MagicLow) / 100 & " - " & ((Item(GetPlayerEquipment(MyIndex, Weapon)).MATK + (GetPlayerStat(MyIndex, Intelligence) * 4) + (GetPlayerLevel(MyIndex) * 2))) * (Item(GetPlayerEquipment(MyIndex, Weapon)).MagicHigh / 100)
    Else
        lblInt.Caption = ((GetPlayerLevel(MyIndex) * 2) + (GetPlayerStat(MyIndex, Intelligence) * 4)) / 2 & " - " & (GetPlayerLevel(MyIndex) * 2) + (GetPlayerStat(MyIndex, Intelligence) * 4)
    End If
    
    ATKL = (GetPlayerStat(MyIndex, willpower) * 3.5) + (GetPlayerLevel(MyIndex) * 2)

    ' Fixed lblStr And lblLong
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        lblStr.Caption = ((((Item(GetPlayerEquipment(MyIndex, Weapon)).Data2) + (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgLow) / 100) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Weapon)).Data2) + (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgHigh) / 100)
        ' Fixed ProjecTile Damage
        If (Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic) > 0 Then
            frmMain.lblLongAttack.Caption = ((((Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Damage + ATKL) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgLow) / 100)) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Damage + ATKL) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgHigh) / 100))
        Else
            frmMain.lblLongAttack.Caption = "0 - 0"
        End If
    Else
        lblStr.Caption = (((GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)) / 2) & " - " & (GetPlayerStat(MyIndex, Strength) * 3.5) + (GetPlayerLevel(MyIndex) * 2)
        frmMain.lblLongAttack.Caption = "0 - 0"
    End If
    
    ' lblCritATK
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            If (Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic) > 0 Then
                lblCritATK.Caption = (120 + (GetPlayerStat(MyIndex, willpower))) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritATK & " %"
            Else
                lblCritATK.Caption = (120 + (GetPlayerStat(MyIndex, willpower))) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritATK & " %"
            End If
        Else ' ถ้าไม่ใส่อาวุธ
            lblCritATK.Caption = 120 + (GetPlayerStat(MyIndex, willpower))
        End If
        
    ' lblWalk
    If Int(WALK_SPEED + Int(GetPlayerStat(MyIndex, Stats.Agility) / 50)) < 6 Then
        lblWalk.Caption = Int(WALK_SPEED + Int(GetPlayerStat(MyIndex, Stats.Agility) / 50)) & " หน่วย."
    Else
        lblWalk.Caption = "6 หน่วย (สูงสุด)."
    End If
    
    ' lblAttackspeed
    ' ความเร็วในการโจมตี
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            Speed = ((2000 + Item(GetPlayerEquipment(MyIndex, Weapon)).Speed) - Item(GetPlayerEquipment(MyIndex, Weapon)).SpeedLow) - ((GetPlayerStat(MyIndex, Stats.Agility) * 5))
            If Speed > 1000 Then
                lblAttackspeed.Caption = FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            Else
                If Speed > 200 Then
                    lblAttackspeed.Caption = "0" & FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
                Else
                    lblAttackspeed.Caption = "0.100 ครั้ง/วินาที"
                End If
            End If
        Else
            Speed = 2000 - ((GetPlayerStat(MyIndex, Stats.Agility) * 5))
            If Speed > 1000 Then
                lblAttackspeed.Caption = FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            Else
                lblAttackspeed.Caption = "0" & FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            End If
        End If
        
        ' lblNDEF
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblNDEF.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).NDEF & " %"
        Else
            lblNDEF.Caption = "0 %"
        End If
        
        ' lblKick
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblKick.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).Kick & " %"
        Else
            lblKick.Caption = "0 %"
        End If
        
        ' lblCastTime ร่ายเวทย์ v2
        
        Select Case (1 + (GetPlayerStat(MyIndex, willpower) / 50))
            Case 7: CT = 100
            Case 6: CT = 50
            Case 5: CT = 33
            Case 4: CT = 100 - 25
            Case 3: CT = 100 - 33
            Case 2: CT = 100 - 50
            Case 1: CT = 0
            Case Else: CT = 0
        End Select
        
        CT = 100 - (100 / (1 + (GetPlayerStat(MyIndex, willpower) / 50)))
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblCastTime.Caption = (100 - (Item(GetPlayerEquipment(MyIndex, Weapon)).DelayDown * 100)) & " + " & FormatNumber(CT, 2) & " %"
        Else
            lblCastTime.Caption = FormatNumber(CT, 2) & " %"
        End If
                
        ' lblVampire
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblVampire.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).Vampire & " %"
        Else
            lblVampire.Caption = "0 %"
        End If
    
        ' lblRegenHP
        
            RegenHp = (GetPlayerStat(MyIndex, Stats.Endurance) * 2) + (GetPlayerLevel(MyIndex) / 2) + (GetPlayerMaxVital(MyIndex, HP) * 0.01) + 2
        
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Weapon)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Weapon)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Armor) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Armor)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Armor)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Helmet)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Helmet)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Shield) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Shield)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Shield)).RegenHp)
                End If
            End If
    
            lblRegenHP.Caption = RegenHp & " หน่วย."
    
        ' lblRegenMP
        
            RegenMp = (GetPlayerStat(MyIndex, Stats.Intelligence)) + (GetPlayerLevel(MyIndex) / 4) + (GetPlayerMaxVital(MyIndex, MP) * 0.01) + 1
            
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Weapon)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Weapon)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Armor) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Armor)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Armor)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Helmet)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Helmet)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Shield) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Shield)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Shield)).RegenMp)
                End If
            End If
    
            lblRegenMP.Caption = RegenMp & " หน่วย."
    
        ' lblDEF
        
        If GetPlayerEquipment(MyIndex, Armor) > 0 Then
            Def = Def + Item(GetPlayerEquipment(MyIndex, Armor)).Data2
        End If
    
        If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
            Def = Def + Item(GetPlayerEquipment(MyIndex, Helmet)).Data2
        End If
    
        Def = Def + (GetPlayerStat(MyIndex, Endurance) * 2) + (GetPlayerLevel(MyIndex) * 2)
        
        ' Check berserker
        If GetPlayerClass(MyIndex) = 4 Then ' Berserker Class None def
            Def = 0
        End If
    
        lblDEF.Caption = Def & " หน่วย."

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblTrainStat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Option1_Click()
Chat1(0) = 1
End Sub

Private Sub Option2_Click()
Chat1(0) = 0
End Sub

Private Sub optMOff_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Music = 0
    ' stop music playing
    StopMidi
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMOn_Click()
Dim MusicFile As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Music = 1
    ' start music playing
    
    MusicFile = Trim$(Map.Music)
    If Not MusicFile = Music_Playing Then
        If Not MusicFile = "None." Then
            PlayMidi MusicFile
        Else
            StopMidi
        End If
    End If
    
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOff_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Sound = 0
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Sound = 1
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCover_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picHotbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SlotNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SlotNum = IsHotbarSlot(X, Y)

    If Button = 1 Then
        If SlotNum <> 0 Then
            SendHotbarUse SlotNum
        End If
    ElseIf Button = 2 Then
        If SlotNum <> 0 Then
            SendHotbarChange 0, 0, SlotNum
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picHotbar_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picHotbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SlotNum As Long
Dim spellslot As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SlotNum = IsHotbarSlot(X, Y)
    
    If SlotNum <> 0 Then
        If Hotbar(SlotNum).sType = 2 Then ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Hotbar(SlotNum).Slot = PlayerSpells(i) Then
                    spellslot = i
                    Exit For
                End If
            Next
        End If
    End If

    If SlotNum <> 0 Then
        If Hotbar(SlotNum).sType = 1 Then ' item
            X = X + picHotbar.Left + 1
            Y = Y + picHotbar.Top - picItemDesc.Height - 1
            UpdateDescWindow Hotbar(SlotNum).Slot, X, Y
            LastItemDesc = Hotbar(SlotNum).Slot ' set it so you don't re-set values
            Exit Sub
        ElseIf Hotbar(SlotNum).sType = 2 Then ' spell
            X = X + picHotbar.Left + 1
            Y = Y + picHotbar.Top - picSpellDesc.Height - 1
            UpdateSpellWindow Hotbar(SlotNum).Slot, X, Y, spellslot ' fixed
            LastSpellDesc = Hotbar(SlotNum).Slot  ' set it so you don't re-set values
            Exit Sub
        End If
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' no spell was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picHotbar_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InMapEditor Then
        Call MapEditorMouseDown(Button, X, Y, False)
    Else
        ' left click
        If Button = vbLeftButton Then
            ' targetting
            Call PlayerSearch(CurX, CurY)
        ' right click
        ElseIf Button = vbRightButton Then
            If ShiftDown Then
                ' admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
            End If
        End If
    End If

    Call SetFocusOnChat
    If frmEditor_Events.Visible Then frmEditor_Events.SetFocus
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picScreen_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' ย้อนกลับเมาส์เป็นปกติ
    
   ' If oldCursor > 0 Then
   '     Call SetClassLong(frmMain.hwnd, GCL_HCURSOR, oldCursor)
   '     Call SetClassLong(picScreen.hwnd, GCL_HCURSOR, oldCursor)
   '     DestroyCursor newCursor
   ' End If
    
    ' cursor ภาพเมาส์
  '     newCursor = LoadCursorFromFile("data files\graphics\cursor\Default.ani")
       
   '    If newCursor > 0 Then
    '        oldCursor = SetClassLong(frmMain.hwnd, GCL_HCURSOR, newCursor)
    '        oldCursor = SetClassLong(picScreen.hwnd, GCL_HCURSOR, newCursor)
 '     End If

    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((Y + Camera.Top) \ PIC_Y)

    If InMapEditor Then
        frmEditor_Map.shpLoc.Visible = False

        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, X, Y)
        End If
    End If
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picScreen_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsShopItem = 0

    For i = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(i).Item > 0 And Shop(InShop).TradeItem(i).Item <= MAX_ITEMS Then
            With tempRec
                .Top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                .Bottom = .Top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picShop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
End Sub

Private Sub picShopItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim shopItem As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    shopItem = IsShopItem(X, Y)
    
    If shopItem > 0 Then
        Select Case ShopAction
            Case 0 ' no action, give cost
                With Shop(InShop).TradeItem(shopItem)
                    AddText "ไอเทมนี้ราคา " & .CostValue & " " & Trim$(Item(.CostItem).Name) & " กรุณากดปุ่ม Buy และคลิกไอเทมนี้ เพื่อซื้อ.", White
                End With
            Case 1 ' buy item
                ' buy item code
                BuyItem shopItem
        End Select
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picShopItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim shopslot As Long
Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    shopslot = IsShopItem(X, Y)

    If shopslot <> 0 Then
        X2 = X + picShop.Left + picShopItems.Left + 1
        Y2 = Y + picShop.Top + picShopItems.Top + 1
        UpdateDescWindow Shop(InShop).TradeItem(shopslot).Item, X2, Y2
        LastItemDesc = Shop(InShop).TradeItem(shopslot).Item
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpellDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picSpellDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpellDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_DblClick()
Dim spellnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = IsPlayerSpell(SpellX, SpellY)

    If spellnum <> 0 Then
        If SpellBuffer = spellnum Then Exit Sub
        Call CastSpell(spellnum)
        Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim spellnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = IsPlayerSpell(SpellX, SpellY)
    If Button = 1 Then ' left click
        If spellnum <> 0 Then
            DragSpell = spellnum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If spellnum <> 0 Then
            Dialogue "ลบสกิล", "คุณต้องการลบสกิล " & Trim$(Spell(PlayerSpells(spellnum)).Name) & " ใช่ไหม?", DIALOGUE_TYPE_FORGET, True, spellnum
            Exit Sub
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim spellslot As Long
Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellX = X
    SpellY = Y
    
    spellslot = IsPlayerSpell(X, Y)
    
    If DragSpell > 0 Then
        Call BltDraggedSpell(X + picSpells.Left, Y + picSpells.Top)
    Else
        If spellslot <> 0 Then
            X2 = X + picSpells.Left - picSpellDesc.Width - 1
            Y2 = Y + picSpells.Top - picSpellDesc.Height - 1
            UpdateSpellWindow PlayerSpells(spellslot), X2, Y2, spellslot
            LastSpellDesc = PlayerSpells(spellslot)
            Exit Sub
        End If
    End If
    
    picSpellDesc.Visible = False
    LastSpellDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DragSpell > 0 Then
        ' drag + drop
        For i = 1 To MAX_PLAYER_SPELLS
            With rec_pos
                .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If DragSpell <> i Then
                        SendChangeSpellSlots DragSpell, i
                        Exit For
                    End If
                End If
            End If
        Next
        ' hotbar
        For i = 1 To MAX_HOTBAR
            With rec_pos
                .Top = picHotbar.Top - picSpells.Top
                .Left = picHotbar.Left - picSpells.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.Top - picSpells.Top + 32
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 2, DragSpell, i
                    DragSpell = 0
                    picTempSpell.Visible = False
                    Exit Sub
                End If
            End If
        Next
    End If

    DragSpell = 0
    picTempSpell.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Picture1_Click()
    Dim Speed As Long, Def As Long, Mdef As Long, RegenHp As Long, RegenMp As Long
Dim ATKL, numWeapon As Long, CT As Long
Dim Dodge As Long
    ' แสดงข้อมูลส่วนตัว
    picMe.Visible = True

    Def = 0
    Mdef = 0
    RegenHp = 0
    RegenMp = 0
    Dodge = 0
    
    ' lblCrit
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            If (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate > 0 And (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate < 1 Then
                lblCrit.Caption = "0" & (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate & " %"
            Else
                lblCrit.Caption = (GetPlayerStat(MyIndex, willpower) / 2.5) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritRate & " %"
            End If
        Else
            If (GetPlayerStat(MyIndex, willpower) / 2.5) > 0 And (GetPlayerStat(MyIndex, willpower) / 2.5) < 1 Then
                lblCrit.Caption = "0" & (GetPlayerStat(MyIndex, willpower) / 2.5) & " %"
            Else
                lblCrit.Caption = (GetPlayerStat(MyIndex, willpower) / 2.5) & " %"
            End If
        End If
    
    ' Dodge
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Weapon)).Dodge
    End If
    
    If GetPlayerEquipment(MyIndex, Armor) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Armor)).Dodge
    End If
    
    If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Helmet)).Dodge
    End If
    
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        Dodge = Dodge + Item(GetPlayerEquipment(MyIndex, Shield)).Dodge
    End If
    
    ' lblStrLHand
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        If Item(GetPlayerEquipment(MyIndex, Shield)).LHand > 0 Then
            lblStrLHand.Caption = ((((Item(GetPlayerEquipment(MyIndex, Shield)).Data2) + (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Shield)).DmgLow) / 100) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Shield)).Data2) + (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Shield)).DmgHigh) / 100)
        Else
            lblStrLHand.Caption = "0 - 0"
        End If
    Else
        lblStrLHand.Caption = "0 - 0"
    End If
    
    ' Mdef
    If GetPlayerEquipment(MyIndex, Armor) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Armor)).MATK
    End If
    
    If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Helmet)).MATK
    End If
    
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        Mdef = Mdef + Item(GetPlayerEquipment(MyIndex, Shield)).MATK
    End If
    
    Mdef = Mdef + (GetPlayerLevel(MyIndex) * 2)
    
    ' lblMdef
    lblDefInt.Caption = GetPlayerStat(MyIndex, Intelligence) + Mdef & " -  " & (GetPlayerStat(MyIndex, Intelligence) * 2) + Mdef
    
    ' lblDodge
        If Dodge + (GetPlayerStat(MyIndex, Agility) / 2.5) > 0 And Dodge + (GetPlayerStat(MyIndex, Agility) / 2.5) < 1 Then
            lblDodge.Caption = "0" & Dodge + (GetPlayerStat(MyIndex, Agility) / 4) & " %"
        Else
            lblDodge.Caption = Dodge + (GetPlayerStat(MyIndex, Agility) / 4) & " %"
        End If
        
    ' lblBlock
        If GetPlayerEquipment(MyIndex, Shield) > 0 Then
            lblBlock.Caption = Item(GetPlayerEquipment(MyIndex, Shield)).Data2 + (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
        Else
            If (GetPlayerStat(MyIndex, Endurance) / 10) > 0 And (GetPlayerStat(MyIndex, Endurance) / 10) < 1 Then
                lblBlock.Caption = "0" & (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
            Else
                lblBlock.Caption = (GetPlayerStat(MyIndex, Endurance) / 10) & " %"
            End If
        End If

    ' Fixed bug <For now>
    lblStr.Caption = (((GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) / 2) & " - " & (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)
    frmMain.lblLongAttack.Caption = "0 - 0"
    
    ' lblInt
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        lblInt.Caption = (((Item(GetPlayerEquipment(MyIndex, Weapon)).MATK + (GetPlayerStat(MyIndex, Intelligence) * 2) + (GetPlayerLevel(MyIndex) * 2))) * Item(GetPlayerEquipment(MyIndex, Weapon)).MagicLow) / 100 & " - " & ((Item(GetPlayerEquipment(MyIndex, Weapon)).MATK + (GetPlayerStat(MyIndex, Intelligence) * 2) + (GetPlayerLevel(MyIndex) * 2))) * (Item(GetPlayerEquipment(MyIndex, Weapon)).MagicHigh / 100)
    Else
        lblInt.Caption = ((GetPlayerLevel(MyIndex) * 2) + (GetPlayerStat(MyIndex, Intelligence) * 2)) / 2 & " - " & (GetPlayerLevel(MyIndex) * 2) + (GetPlayerStat(MyIndex, Intelligence) * 2)
    End If
    
    ATKL = (GetPlayerStat(MyIndex, willpower) * 2) + (GetPlayerLevel(MyIndex) * 2)

    ' Fixed lblStr And lblLong
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        lblStr.Caption = ((((Item(GetPlayerEquipment(MyIndex, Weapon)).Data2) + (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgLow) / 100) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Weapon)).Data2) + (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgHigh) / 100)
        ' Fixed ProjecTile Damage
        If (Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic) > 0 Then
            frmMain.lblLongAttack.Caption = ((((Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Damage + ATKL) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgLow) / 100)) & " - " & ((((Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Damage + ATKL) * Item(GetPlayerEquipment(MyIndex, Weapon)).DmgHigh) / 100))
        Else
            frmMain.lblLongAttack.Caption = "0 - 0"
        End If
    Else
        lblStr.Caption = (((GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)) / 2) & " - " & (GetPlayerStat(MyIndex, Strength) * 2) + (GetPlayerLevel(MyIndex) * 2)
        frmMain.lblLongAttack.Caption = "0 - 0"
    End If
    
    ' lblCritATK
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            If (Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic) > 0 Then
                lblCritATK.Caption = (120 + (GetPlayerStat(MyIndex, willpower))) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritATK & " %"
            Else
                lblCritATK.Caption = (120 + (GetPlayerStat(MyIndex, willpower))) + Item(GetPlayerEquipment(MyIndex, Weapon)).CritATK & " %"
            End If
        Else ' ถ้าไม่ใส่อาวุธ
            lblCritATK.Caption = 120 + (GetPlayerStat(MyIndex, willpower))
        End If
        
    ' lblWalk
    If Int(WALK_SPEED + Int(GetPlayerStat(MyIndex, Stats.Agility) / 50)) < 6 Then
        lblWalk.Caption = Int(WALK_SPEED + Int(GetPlayerStat(MyIndex, Stats.Agility) / 50)) & " หน่วย."
    Else
        lblWalk.Caption = "6 หน่วย (สูงสุด)."
    End If
    
    ' lblAttackspeed
    ' ความเร็วในการโจมตี
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            Speed = ((2000 + Item(GetPlayerEquipment(MyIndex, Weapon)).Speed) - Item(GetPlayerEquipment(MyIndex, Weapon)).SpeedLow) - ((GetPlayerStat(MyIndex, Stats.Agility) * 5))
            If Speed > 1000 Then
                lblAttackspeed.Caption = FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            Else
                If Speed > 200 Then
                    lblAttackspeed.Caption = "0" & FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
                Else
                    lblAttackspeed.Caption = "0.100 ครั้ง/วินาที"
                End If
            End If
        Else
            Speed = 2000 - ((GetPlayerStat(MyIndex, Stats.Agility) * 5))
            If Speed > 1000 Then
                lblAttackspeed.Caption = FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            Else
                lblAttackspeed.Caption = "0" & FormatNumber(Speed / 1000, 3) & " ครั้ง/วินาที"
            End If
        End If
        
        ' lblNDEF
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblNDEF.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).NDEF & " %"
        Else
            lblNDEF.Caption = "0 %"
        End If
        
        ' lblKick
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblKick.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).Kick & " %"
        Else
            lblKick.Caption = "0 %"
        End If
        
        ' lblCastTime ร่ายเวทย์ v2
        
        Select Case (1 + (GetPlayerStat(MyIndex, willpower) / 50))
            Case 1: CT = 100
            Case 2: CT = 50
            Case 3: CT = 33
            Case 4: CT = 25
            Case 5: CT = 20
            Case 6: CT = 16
            Case Else: CT = 0
        End Select
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblCastTime.Caption = (100 - (Item(GetPlayerEquipment(MyIndex, Weapon)).DelayDown * 100)) & " + " & (100 - CT) & " %"
        Else
            lblCastTime.Caption = (100 - CT) & " %"
        End If
                
        ' lblVampire
        
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            lblVampire.Caption = Item(GetPlayerEquipment(MyIndex, Weapon)).Vampire & " %"
        Else
            lblVampire.Caption = "0 %"
        End If
    
        ' lblRegenHP
        
            RegenHp = (GetPlayerStat(MyIndex, Stats.Endurance) * 2) + (GetPlayerLevel(MyIndex) / 2) + (GetPlayerMaxVital(MyIndex, HP) * 0.01) + 2
        
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Weapon)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Weapon)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Armor) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Armor)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Armor)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Helmet)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Helmet)).RegenHp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Shield) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Shield)).RegenHp > 0 Then
                    RegenHp = RegenHp + (Item(GetPlayerEquipment(MyIndex, Shield)).RegenHp)
                End If
            End If
    
            lblRegenHP.Caption = RegenHp & " หน่วย."
    
        ' lblRegenMP
        
            RegenMp = (GetPlayerStat(MyIndex, Stats.Intelligence)) + (GetPlayerLevel(MyIndex) / 4) + (GetPlayerMaxVital(MyIndex, MP) * 0.01) + 1
            
            If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Weapon)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Weapon)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Armor) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Armor)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Armor)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Helmet)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Helmet)).RegenMp)
                End If
            End If
            
            If GetPlayerEquipment(MyIndex, Shield) > 0 Then
                If Item(GetPlayerEquipment(MyIndex, Shield)).RegenMp > 0 Then
                    RegenMp = RegenMp + (Item(GetPlayerEquipment(MyIndex, Shield)).RegenMp)
                End If
            End If
    
            lblRegenMP.Caption = RegenMp & " หน่วย."
    
        ' lblDEF
        
        If GetPlayerEquipment(MyIndex, Armor) > 0 Then
            Def = Def + Item(GetPlayerEquipment(MyIndex, Armor)).Data2
        End If
    
        If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
            Def = Def + Item(GetPlayerEquipment(MyIndex, Helmet)).Data2
        End If
    
        Def = Def + (GetPlayerStat(MyIndex, Endurance) * 2) + (GetPlayerLevel(MyIndex) * 2)
        
        ' Check berserker
        If GetPlayerClass(MyIndex) = 4 Then ' Berserker Class None def
            Def = 0
        End If
    
        lblDEF.Caption = Def & " หน่วย."

End Sub

Private Sub picYourTrade_DblClick()
Dim TradeNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(TradeX, TradeY, True)

    If TradeNum <> 0 Then
        UntradeItem TradeNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picYourTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeX = X
    TradeY = Y
    
    TradeNum = IsTradeItem(X, Y, True)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picYourTrade.Left + 4
        Y = Y + picTrade.Top + picYourTrade.Top + 4
        UpdateDescWindow GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).Num), X, Y
        LastItemDesc = GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).Num) ' set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTheirTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(X, Y, False)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picTheirTrade.Left + 4
        Y = Y + picTrade.Top + picTheirTrade.Top + 4
        UpdateDescWindow TradeTheirOffer(TradeNum).Num, X, Y
        LastItemDesc = TradeTheirOffer(TradeNum).Num ' set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTheirTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAAmount.Caption = "จำนวน : " & scrlAAmount.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAAmount_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAItem.Value <= 0 Then
        lblAItem.Caption = "เสกไอเทม : ไม่มี"
    Else
        If Trim$(Item(scrlAItem.Value).Name) <> vbNullString Then
            lblAItem.Caption = "เสกไอเทม : " & Trim$(Item(scrlAItem.Value).Name)
        Else
            lblAItem.Caption = "เสกไอเทม : ไม่มี"
        End If
    End If

    If Item(scrlAItem.Value).Type = ITEM_TYPE_CURRENCY Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    scrlAAmount.Enabled = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlVolume_Change()
lblVolume.Caption = "ความดังเสียง : " & scrlVolume.Value
DefaultVolume = scrlVolume.Value
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If
    
    'exitgame
    If GetKeyState(vbKeyEscape) < 0 Then
         frmExitMenu.Show
         ' logoutGame
     End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access > 0 Then
                If picAdmin.Visible = True Then
                    picAdmin.Visible = False
                    If frmMain.picAdmin.Visible = False Then frmMain.Width = 11340
                    SetFocusOnGame
                Else
                    picAdmin.Visible = True
                    If frmMain.picAdmin.Visible = True Then frmMain.Width = 14220
                    SetFocusOnGame
                End If
                ' fixed close picAdmin focus cbMap
                ' picAdmin.Visible = Not picAdmin.Visible
            End If
            
        Case vbKeyDelete
            'MsgBox Player(MyIndex).PlayerQuest(1).Status
            'PlayerHandleQuest 1, 2
    End Select
        
    If chaton = True Then
        Else
    Select Case KeyCode
        Case vbKeyI

                ' show the window
                If Not picInventory.Visible Then
                    picInventory.Visible = True
                Else
                    picInventory.Visible = False
                End If
                
                picCharacter.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = False
                picPet.Visible = False
                BltInventory
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
        Case vbKeyK
                
                ' send packet
                Set Buffer = New clsBuffer
                Buffer.WriteLong CSpells
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                
                ' show the window
                If Not picSpells.Visible Then
                    picSpells.Visible = True
                Else
                    picSpells.Visible = False
                End If
                
                picInventory.Visible = False
                picCharacter.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = False
                picPet.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
        Case vbKeyC
                
                ' send packet
                SendRequestPlayerData
                
                ' show the window
                If Not picCharacter.Visible Then
                    picCharacter.Visible = True
                Else
                    picCharacter.Visible = False
                End If
                
                picInventory.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = False
                picPet.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                ' Render
                BltEquipment
                BltFace
                SetFocusOnGame
        Case vbKeyO

                ' show the window
                picCharacter.Visible = False
                picInventory.Visible = False
                picSpells.Visible = False
                
                If Not picOptions.Visible Then
                    picOptions.Visible = True
                Else
                    picOptions.Visible = False
                End If
                
                picParty.Visible = False
                picParty2.Visible = False
                picParty3.Visible = False
                picQuestLog.Visible = False
                picPet.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick
                SetFocusOnGame
        Case vbKeyP

                ' show the window
            picCharacter.Visible = False
            picInventory.Visible = False
            picSpells.Visible = False
            picOptions.Visible = False
            
            If Not picParty.Visible Then
                picParty.Visible = True
            Else
                picParty.Visible = False
            End If
            
            picParty2.Visible = False
            picParty3.Visible = False
            picQuestLog.Visible = False
            picPet.Visible = False
            ' play sound
            PlaySound Sound_ButtonClick
            SetFocusOnGame
        Case vbKeyL
            Call AddText("ยังไม่เปิดให้บริการในส่วนนี้ค่ะ.", BrightRed)
            Exit Sub
            
                ' show the window
            picCharacter.Visible = False
            picInventory.Visible = False
            picSpells.Visible = False
            picOptions.Visible = False
            picParty.Visible = False
            picParty2.Visible = False
            picParty3.Visible = False
            picQuestLog.Visible = False
            
            If Not picPet.Visible Then
                picPet.Visible = True
            Else
                picPet.Visible = False
            End If

            ' play sound
            PlaySound Sound_ButtonClick
            SetFocusOnGame
        Case vbKeyQ
        
                UpdateQuestLog
            ' show the window
            picCharacter.Visible = False
            picInventory.Visible = False
            picSpells.Visible = False
            picOptions.Visible = False
            picParty.Visible = False
            picParty2.Visible = False
            picParty3.Visible = False
            
            If Not picQuestLog.Visible Then
                picQuestLog.Visible = True
            Else
                picQuestLog.Visible = False
            End If
            
            picPet.Visible = False
            ' play sound
            PlaySound Sound_ButtonClick
            SetFocusOnGame
        Case vbKeyF
            Call PetFollow(MyIndex)
            PlaySound Sound_ButtonClick
            'Call AddText("สั่งสัตว์เลี้ยงให้ติดตามคุณแล้ว...", BrightGreen)
            ' pet followed
            SetFocusOnGame
        Case vbKeyE
            Call PetAttack(MyIndex)
            PlaySound Sound_ButtonClick
            'Call AddText("สั่งสัตว์เลี้ยงโจมตี !!", BrightGreen)
            ' pet attack
            SetFocusOnGame
        Case vbKeyR
            Call PetDisband(MyIndex)
            PlaySound Sound_ButtonClick
            'Call AddText("เก็บสัตว์เลี้ยง !!", BrightGreen)
            ' pet remove
            SetFocusOnGame
        Case vbKeyM
            ShowMiniMap = Not ShowMiniMap
            
                 '<---- add new things
        End Select
    End If
        
    ' hotbar
    If picAdmin.Visible = False Then
    
    For i = 1 To MAX_HOTBAR
    ' Key code
    ' 111 = F1-F12 / 96 = 0-9 [Numlock] / 48 = 1 to =
        If KeyCode = 111 + i Then
            SendHotbarUse i
        End If
    Next
    
    If Not txtMyChat.Visible Then
        For i = 1 To 9
            If KeyCode = 48 + i Then
                SendHotbarUse i
            End If
        Next
        
        If KeyCode = 48 Then ' 0
            SendHotbarUse 10
        ElseIf KeyCode = 189 Then ' -
            SendHotbarUse 11
        ElseIf KeyCode = 187 Then ' =
            SendHotbarUse 12
        End If
    End If
    
    End If
    
    ' handles delete events
    If KeyCode = vbKeyDelete Then
        If InMapEditor Then DeleteEvent CurX, CurY
    End If
    
    ' handles copy + pasting events
    If KeyCode = vbKeyC Then
        If ControlDown Then
            If InMapEditor Then
                CopyEvent_Map CurX, CurY
            End If
        End If
    End If
    If KeyCode = vbKeyV Then
        If ControlDown Then
            If InMapEditor Then
                PasteEvent_Map CurX, CurY
            End If
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub tmrChat_Timer()
    Call SayMsg(vbNullString)
    tmrChat.Enabled = False
End Sub

Private Sub txtMyChat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    MyText = txtMyChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMyChat_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChat_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SetFocusOnChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtChat_GotFocus", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ***************
' ** Inventory **
' ***************
Private Sub picInventory_DblClick()
    Dim InvNum As Long
    Dim Value As Long
    Dim multiplier As Double
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragInvSlotNum = 0
    InvNum = IsInvItem(InvX, InvY)

    If InvNum <> 0 Then
    
        ' are we in a shop?
        If InShop > 0 Then
            Select Case ShopAction
                Case 0 ' nothing, give value
                    multiplier = Shop(InShop).BuyRate / 100
                    Value = Item(GetPlayerInvItemNum(MyIndex, InvNum)).Price * multiplier
                    If Value > 0 Then
                        AddText "คุณขายไอเทมได้ " & Value & " .", White
                    Else
                        AddText "ไอเทมนี้เป็นไอเทมที่ไม่สามารถขายได้.", BrightRed
                    End If
                Case 2 ' 2 = sell
                    SellItem InvNum
            End Select
            
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 2 ' deposit
                lblCurrency.Caption = "คุณต้องการฝากเท่าไร?"
                tmpCurrencyItem = InvNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
                
            Call DepositItem(InvNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For i = 1 To MAX_INV
                If TradeYourOffer(i).Num = InvNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)).Type = ITEM_TYPE_CURRENCY Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).Num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 4 ' offer in trade
                lblCurrency.Caption = "คุณต้องการแลกเปลี่ยนจำนวนเท่าไร?"
                tmpCurrencyItem = InvNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
            
            Call TradeItem(InvNum, 0)
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(InvNum)
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsEqItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsEqItem = 0

    For i = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, i) > 0 And GetPlayerEquipment(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .Top = EqTop
                .Bottom = .Top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsEqItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsInvItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsInvItem = 0

    For i = 1 To MAX_INV

        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS

        If PlayerSpells(i) > 0 And PlayerSpells(i) <= MAX_SPELLS Then

            With tempRec
                .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean) As Long
    Dim tempRec As RECT
    Dim i As Long
    Dim itemnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsTradeItem = 0

    For i = 1 To MAX_INV
    
        If Yours Then
            itemnum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)
        Else
            itemnum = TradeTheirOffer(i).Num
        End If

        If itemnum > 0 And itemnum <= MAX_ITEMS Then

            With tempRec
                .Top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsTradeItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTradeItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InvNum = IsInvItem(X, Y)

    If Button = 1 Then
        If InvNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = InvNum
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If InvNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(MyIndex, InvNum) > 0 Then
                        CurrencyMenu = 1 ' drop
                        lblCurrency.Caption = "คุณต้องการทิ้งจำนวน?"
                        tmpCurrencyItem = InvNum
                        txtCurrency.text = vbNullString
                        picCurrency.Visible = True
                        txtCurrency.SetFocus
                    End If
                Else
                    Call SendDropItem(InvNum, 0)
                End If
            End If
        End If
    End If

    SetFocusOnChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InvX = X
    InvY = Y

    If DragInvSlotNum > 0 Then
        If InTrade > 0 Then Exit Sub
        If InBank Or InShop Then Exit Sub
        Call BltInventoryItem(X + picInventory.Left, Y + picInventory.Top)
    Else
        InvNum = IsInvItem(X, Y)

        If InvNum <> 0 Then
            ' exit out if we're offering that item
            If InTrade Then
                For i = 1 To MAX_INV
                    If TradeYourOffer(i).Num = InvNum Then
                        ' is currency?
                        If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)).Type = ITEM_TYPE_CURRENCY Then
                            ' only exit out if we're offering all of it
                            If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).Num) Then
                                Exit Sub
                            End If
                        Else
                            Exit Sub
                        End If
                    End If
                Next
            End If
            X = X + picInventory.Left - picItemDesc.Width - 1
            Y = Y + picInventory.Top - picItemDesc.Height - 1
            UpdateDescWindow GetPlayerInvItemNum(MyIndex, InvNum), X, Y
            LastItemDesc = GetPlayerInvItemNum(MyIndex, InvNum) ' set it so you don't re-set values
            Exit Sub
        End If
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        ' drag + drop
        For i = 1 To MAX_INV
            With rec_pos
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then '
                    If DragInvSlotNum <> i Then
                        SendChangeInvSlots DragInvSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
        ' hotbar
        For i = 1 To MAX_HOTBAR
            With rec_pos
                .Top = picHotbar.Top - picInventory.Top
                .Left = picHotbar.Left - picInventory.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.Top - picInventory.Top + 32
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 1, DragInvSlotNum, i
                    DragInvSlotNum = 0
                    picTempInv.Visible = False
                    BltHotbar
                    Exit Sub
                End If
            End If
        Next
    End If

    DragInvSlotNum = 0
    picTempInv.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picItemDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picItemDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picItemDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *****************
' ** Char window **
' *****************

Private Sub picCharacter_Click()
    Dim EqNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EqNum = IsEqItem(EqX, EqY)

    If EqNum <> 0 Then
        SendUnequip EqNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim EqNum As Long
    Dim X2 As Long, Y2 As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EqX = X
    EqY = Y
    EqNum = IsEqItem(X, Y)

    If EqNum <> 0 Then
        Y2 = Y + picCharacter.Top - frmMain.picItemDesc.Height - 1
        X2 = X + picCharacter.Left - frmMain.picItemDesc.Width - 1
        UpdateDescWindow GetPlayerEquipment(MyIndex, EqNum), X2, Y2
        LastItemDesc = GetPlayerEquipment(MyIndex, EqNum) ' set it so you don't re-set values
        Exit Sub
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ****************
' ** Admin Menu **
' ****************

Private Sub cmdALoc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    BLoc = Not BLoc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdALoc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendRequestEditMap
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp2Me_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarpMe2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("หมายเลขแผนที่ผิดพลาด.", Red)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASprite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtASprite.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtASprite.text)) Then
        Exit Sub
    End If

    SendSetSprite CLng(Trim$(txtASprite.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASprite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMapReport_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    SendMapReport
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMapReport_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdARespawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendMapRespawn
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdARespawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdABan_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdABan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdANpc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditNpc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdANpc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditResource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAResource_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditShop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditSpell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Or Not IsNumeric(Trim$(txtAAccess.text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.text), CLng(Trim$(txtAAccess.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdADestroy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    SendBanDestroy
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdADestroy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If
    
    SendSpawnItem scrlAItem.Value, scrlAAmount.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' bank
Private Sub picBank_DblClick()
Dim bankNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragBankSlotNum = 0

    bankNum = IsBankItem(BankX, BankY)
    If bankNum <> 0 Then
         If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_NONE Then Exit Sub
         
             If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 3 ' withdraw
                lblCurrency.Caption = "คุณต้องการถอนออกจำนวน?"
                tmpCurrencyItem = bankNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
            
         WithdrawItem bankNum, 0
         Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bankNum As Long
                        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    bankNum = IsBankItem(X, Y)
    
    If bankNum <> 0 Then
        
        If Button = 1 Then
            DragBankSlotNum = bankNum
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

' TODO : Add sub to change bankslots client side first so there's no delay in switching
    If DragBankSlotNum > 0 Then
        For i = 1 To MAX_BANK
            With rec_pos
                .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If DragBankSlotNum <> i Then
                        ChangeBankSlots DragBankSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    DragBankSlotNum = 0
    picTempBank.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bankNum As Long, itemnum As Long, ItemType As Long
Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    BankX = X
    BankY = Y
    
    If DragBankSlotNum > 0 Then
        Call BltBankItem(X + picBank.Left, Y + picBank.Top)
    Else
        bankNum = IsBankItem(X, Y)
        
        If bankNum <> 0 Then
            
            X2 = X + picBank.Left + 1
            Y2 = Y + picBank.Top + 1
            UpdateDescWindow Bank.Item(bankNum).Num, X2, Y2
            LastItemDesc = Bank.Item(bankNum).Num
            Exit Sub
        End If
    End If
    
    frmMain.picItemDesc.Visible = False
    LastBankDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsBankItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsBankItem = 0
    
    For i = 1 To MAX_BANK
        If GetBankItemNum(i) > 0 And GetBankItemNum(i) <= MAX_ITEMS Then
        
            With tempRec
                .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    
                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBankItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

'ALATAR

'QuestDialogue:

Private Sub lblQuestAccept_Click()
    PlayerHandleQuest CLng(lblQuestAccept.Tag), 1
    picQuestDialogue.Visible = False
    lblQuestAccept.Visible = False
    lblQuestAccept.Tag = vbNullString
    lblQuestSay = "-"
    RefreshQuestLog
End Sub

Private Sub lblQuestExtra_Click()
    RunQuestDialogueExtraLabel
End Sub

Private Sub lblQuestClose_Click()
    picQuestDialogue.Visible = False
    lblQuestExtra.Visible = False
    lblQuestAccept.Visible = False
    lblQuestAccept.Tag = vbNullString
    lblQuestSay = "-"
End Sub

'QuestLog:

Private Sub picQuestButton_Click()
    'Need to be replaced with imgButton(X) and a proper image
    UpdateQuestLog
    picQuestLog.Visible = Not picQuestLog.Visible
    PlaySound Sound_ButtonClick
End Sub

Private Sub imgQuestButton_Click(Index As Integer)
    If Trim$(lstQuestLog.text) = vbNullString Then Exit Sub
    LoadQuestlogBox Index
End Sub

'/ALATAR

'// Event Allstar

Private Sub lblEventChatContinue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = False
End Sub

Private Sub lblEventChatContinue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = True
End Sub

Private Sub lblEventChatContinue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = True
End Sub

Private Sub lblEventChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = False
End Sub

Private Sub picEventChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = False
End Sub
