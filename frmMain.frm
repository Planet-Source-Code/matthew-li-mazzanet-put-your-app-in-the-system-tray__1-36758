VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Systray Example"
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dialog 
      Left            =   3480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Icon"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "Systray Example"
      Top             =   1680
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   1800
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide from tray"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show in tray"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Icon:"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Tooltip:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Begin VB.Menu mnuTrayShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuTrayHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then Exit Sub
ShowTray
End Sub

Private Sub Command2_Click()
HideTray
End Sub

Private Sub Command3_Click()
dialog.Filter = "*.ico | *.ico"
dialog.ShowOpen
If dialog.FileName = "" Then Exit Sub
Picture1.Picture = LoadPicture(dialog.FileName)
HideTray
ShowTray
End Sub

Private Sub Form_Unload(Cancel As Integer)
HideTray
End Sub

Private Sub mnuTrayExit_Click()
Unload Me
End Sub

Private Sub mnuTrayHide_Click()
Me.Hide
End Sub

Private Sub mnuTrayShow_Click()
Me.Show
End Sub
