VERSION 5.00
Begin VB.Form frmExitMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Exit"
   ClientHeight    =   3315
   ClientLeft      =   9375
   ClientTop       =   4215
   ClientWidth     =   1605
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   1605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd01 
      Caption         =   "�Ѻ����ŧ"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "����͡"
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "�͡"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�ҡ�� �������?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�س��ͧ����͡"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmExitMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmd01_Click()
    frmMain.WindowState = vbMinimized
    frmExitMenu.Hide
End Sub


Private Sub cmdNo_Click()
    frmExitMenu.Hide
End Sub

Private Sub cmdYes_Click()
    
    If party.Leader > 0 Then
        If party.MemberCount > 2 Then
            Call SendPartyChatMsg("�ͺ���˹�һ��������Ѻ " & frmMain.lblPartyMember(2).Caption)
            SendPartyLeave
            Call AddText("�س�͡�ҡ���������� ��سҡ��͡���ա���駤��", BrightRed)
            hasParty = False
            'Exit Sub

        Else
    
            SendPartyLeave
            Call AddText("�س�͡�ҡ���������� ��سҡ��͡���ա���駤��", BrightRed)
            'Exit Sub
        End If
        
    Else
    
    'Call PetDisband(MyIndex)
    exitGame = True
    frmExitMenu.Hide
    Call DestroyGame
        
    End If
    
End Sub
