VERSION 5.00
Begin VB.Form frmEditor_Doors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "แก้ไขประตู"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11085
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Doors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
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
      Left            =   5760
      TabIndex        =   26
      Top             =   6000
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
      Left            =   7320
      TabIndex        =   25
      Top             =   6000
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
      Left            =   4200
      TabIndex        =   24
      Top             =   6000
      Width           =   1455
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
      Left            =   120
      TabIndex        =   23
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "ประตู/สวิตช์ ข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   7695
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
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
      Begin VB.Frame fraDoor 
         Caption         =   "ประตู/สวิตช์ ข้อมูล 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   7455
         Begin VB.Frame Frame5 
            Caption         =   "อะไรปลดล๊อคประตู?"
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
            Left            =   3480
            TabIndex        =   27
            Top             =   480
            Width           =   3135
            Begin VB.HScrollBar scrlSwitch 
               Height          =   255
               Left            =   240
               TabIndex        =   28
               Top             =   480
               Width           =   2415
            End
            Begin VB.Label lblSwitch 
               Caption         =   "ประตู : ไม่มี"
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
               Top             =   240
               Width           =   2895
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "วาร์ป"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   720
            TabIndex        =   16
            Top             =   480
            Width           =   2175
            Begin VB.HScrollBar scrlY 
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   1440
               Width           =   1935
            End
            Begin VB.HScrollBar scrlX 
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   960
               Width           =   1935
            End
            Begin VB.HScrollBar scrlMap 
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   480
               Width           =   1935
            End
            Begin VB.Label lblY 
               Caption         =   "แผนที่ y : 0"
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
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label lblX 
               Caption         =   "แผนที่ x : 0"
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
               TabIndex        =   18
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lblMap 
               Caption         =   "แผนที่ : 0"
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
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame fraToUnlock 
            Caption         =   "ปลดล๊อคด้วยอะไร?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   1560
            TabIndex        =   9
            Top             =   2520
            Width           =   4335
            Begin VB.OptionButton OptUnlock 
               Caption         =   "ไม่มี"
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
               Index           =   2
               Left            =   2760
               TabIndex        =   15
               Top             =   240
               Width           =   975
            End
            Begin VB.Frame Frame4 
               Caption         =   "กุญแจ"
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
               Left            =   480
               TabIndex        =   12
               Top             =   480
               Width           =   3375
               Begin VB.HScrollBar scrlKey 
                  Height          =   255
                  Left            =   240
                  TabIndex        =   13
                  Top             =   480
                  Width           =   2415
               End
               Begin VB.Label lblKey 
                  Caption         =   "กุญแจ : ไม่มี"
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
                  TabIndex        =   14
                  Top             =   240
                  Width           =   3015
               End
            End
            Begin VB.OptionButton OptUnlock 
               Caption         =   "สวิตช์"
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
               Index           =   1
               Left            =   1560
               TabIndex        =   11
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton OptUnlock 
               Caption         =   "กุญแจ"
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
               Index           =   0
               Left            =   600
               TabIndex        =   10
               Top             =   240
               Width           =   975
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ประตู / สวิตช์?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton optDoor 
            Caption         =   "สวิตช์"
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
            Left            =   1200
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optDoor 
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
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ประตู/สวิตช์ รายชื่อ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
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
         Height          =   5325
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Doors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

scrlMap.Max = MAX_MAPS
scrlX.Max = MAX_BYTE
scrlY.Max = MAX_BYTE
scrlSwitch.Max = MAX_DOORS
scrlKey.Max = MAX_ITEMS
End Sub

Private Sub cmdCancel_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call DoorEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_DOORS Then Exit Sub
    
    ClearDoor EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Doors(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    DoorEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call DoorEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DoorEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optDoor_Click(Index As Integer)
Doors(EditorIndex).DoorType = Index
If Index = 0 Then
    Frame6.Visible = True
    fraToUnlock.Visible = True
    Frame5.Visible = False
Else
    Frame6.Visible = False
    fraToUnlock.Visible = False
    Frame5.Visible = True
End If
End Sub

Private Sub OptUnlock_Click(Index As Integer)
Doors(EditorIndex).UnlockType = Index
If Index = 0 Then
    Frame4.Visible = True
Else
    Frame4.Visible = False
End If
End Sub

Private Sub scrlKey_Change()
If scrlKey.Value > 0 Then
lblKey.Caption = "กุญแจ : " & Trim$(Item(scrlKey.Value).Name)
Else
lblKey.Caption = "กุญแจ : ไม่มี"
End If
Doors(EditorIndex).key = scrlKey.Value
End Sub



Private Sub scrlMap_Change()

lblMap.Caption = "แผนที่ : " & scrlMap.Value
Doors(EditorIndex).WarpMap = scrlMap.Value

End Sub

Private Sub scrlSwitch_Change()
If (scrlSwitch.Value > 0) Then
lblSwitch.Caption = "ประตู : " & Trim$(Doors(scrlSwitch.Value).Name)
Else
lblSwitch.Caption = "ประตู : ไม่มี"
End If
Doors(EditorIndex).Switch = scrlSwitch.Value
End Sub



Private Sub scrlX_Change()
lblX.Caption = "แผนที่ x : " & scrlX.Value
Doors(EditorIndex).WarpX = scrlX.Value
End Sub

Private Sub scrlY_Change()
lblY.Caption = "แผนที่ y : " & scrlY.Value
Doors(EditorIndex).WarpY = scrlY.Value
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Doors(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & " : " & Doors(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
