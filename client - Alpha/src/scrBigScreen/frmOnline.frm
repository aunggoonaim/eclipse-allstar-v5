VERSION 5.00
Begin VB.Form frmOnline 
   BackColor       =   &H00FFC0C0&
   Caption         =   "ใครออนไลน์บ้าง ?"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4830
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWarp 
      Caption         =   "วาร์ปไปหา"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิดหน้านี้"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin VB.ListBox lstOnline 
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    frmOnline.Visible = False
    frmOnline.Hide
End Sub

Private Sub cmdWarp_Click()
    Call AddText("ยังไม่เปิดให้บริการค่ะ..", BrightRed)
End Sub
