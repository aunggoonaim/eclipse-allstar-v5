VERSION 5.00
Begin VB.Form frmDialog 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6900
   ClientLeft      =   5355
   ClientTop       =   2085
   ClientWidth     =   9765
   ControlBox      =   0   'False
   Icon            =   "frmDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "����"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�к���Т�������"
      Height          =   5895
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "������� �ա�ҡ��� ����ҹ�������ö�к������������� �ѹ���Ш�ԧ � ��Ѻ"
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   8535
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "9.���ظ 2 ��� / ���ظ����ͧ / ���ظ�����ѡ"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "8.Event / ���ͧ�˵ء�ó� ����͹�к��ͧ�� RPG"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "7.�ѵ������§ (�͹Ҥ�����ʡ���ѭ�ԭ�ѵ��ͧ�����ͤ)"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   3975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "6.�Թ����"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "5.����Ẻ Rang (���/�׹)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "4.�����ʡ���� "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "3.�������� ��Ե����"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "2.������ ��� �����������"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "1.��Ҿ�ҡ�� ���� ���� ���Թ ���ط���"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "�к���㹻Ѩ�غѹ�����к����仹��"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�����������ͧ��"
      Height          =   5895
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   9255
      Begin VB.PictureBox Picture1 
         Height          =   1815
         Left            =   1080
         Picture         =   "frmDialog.frx":1CFA
         ScaleHeight     =   1755
         ScaleWidth      =   7035
         TabIndex        =   22
         Top             =   4080
         Width           =   7095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDialog.frx":3CF5
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   3480
         Width           =   8535
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDialog.frx":3DA7
         Height          =   975
         Left            =   240
         TabIndex        =   20
         Top             =   2520
         Width           =   8415
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "����Ἱ������ � : �Թ����شἹ����� � ������ ������駤�������§�Ἱ������ ���ͤ�¡Ѻ npc warp"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   8415
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl : ���� , Space bar : �红ͧ , 1 �֧ 0 ��� - ��� = : �����Ѵ (��駤����Ѵ�����ҡʡ�� ������� � ��ҧ)"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   8535
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "�����١�� : �� ��ҧ ���� ��� ��� W A S D = �Ǻ����������͹��� , Enter : ᪷"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   8535
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Game basic : ��鹰ҹ"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   8535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDialog.frx":3F2B
         Height          =   1215
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   8775
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "�Դ˹�ҹ��"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
frmDialog.Hide
End Sub

Private Sub OKButton_Click()

If Frame1.Visible = True Then
    Frame1.Visible = False
    Frame2.Visible = True
    Exit Sub
End If

If Frame2.Visible = True Then
    Frame2.Visible = False
    Frame1.Visible = True
    Exit Sub
End If

End Sub
