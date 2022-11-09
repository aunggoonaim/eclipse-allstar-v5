VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help Support"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7995
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton Command5 
         Caption         =   "ข้าอยากกลับเมือง"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ข้อมูลเกม"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "คำสั่งต่างๆ"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ระบบในเกม"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "วิธีเล่น"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "การคำนวนอัตราดรอป ?"
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   $"frmHelp.frx":1CFA
         Height          =   975
         Left            =   2640
         TabIndex        =   7
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "คุณต้องการความช่วยเหลืออะไร?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
