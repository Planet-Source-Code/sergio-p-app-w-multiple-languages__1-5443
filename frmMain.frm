VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmMain.caption"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1000"
   Begin VB.Frame frm1001 
      Caption         =   "Frame "
      Height          =   2835
      Left            =   60
      TabIndex        =   0
      Tag             =   "1001"
      Top             =   0
      Width           =   4125
      Begin VB.CommandButton cmdMenu 
         Caption         =   "cmdMenu"
         Height          =   585
         Left            =   2130
         TabIndex        =   8
         Tag             =   "1007"
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "cmdExit"
         Height          =   345
         Left            =   210
         TabIndex        =   7
         Tag             =   "1005"
         Top             =   2340
         Width           =   3705
      End
      Begin VB.TextBox txtSample 
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   6
         Tag             =   "1006"
         Text            =   "txtSample"
         Top             =   630
         Width           =   1755
      End
      Begin VB.TextBox txtSelected 
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "txtSelected"
         Top             =   1980
         Width           =   1755
      End
      Begin VB.ListBox lstLanguages 
         Height          =   1230
         Left            =   210
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1725
      End
      Begin VB.Label lblSample 
         AutoSize        =   -1  'True
         Caption         =   "lblSample"
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Tag             =   "1004"
         Top             =   390
         Width           =   675
      End
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         Caption         =   "lblSelected"
         Height          =   195
         Left            =   2160
         TabIndex        =   4
         Tag             =   "1003"
         Top             =   1740
         Width           =   780
      End
      Begin VB.Label lblLanguages 
         AutoSize        =   -1  'True
         Caption         =   "lblLanguages"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Tag             =   "1002"
         Top             =   390
         Width           =   945
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "mnuFile"
      Index           =   2000
      Begin VB.Menu mnuSample 
         Caption         =   "mnuSample"
         Index           =   2001
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    ''Unload the Form then Exit
    Unload Me
    Unload frmMenu
    End
End Sub

Private Sub cmdMenu_Click()
    frmMenu.Show
End Sub

Private Sub Form_Load()
    ''Search for ALL the Language Files
    SetLanguageFile "*.lan"
    txtSelected = "english"
    ''Load the default Language
    LoadResStrings frmMain, txtSelected & ".lan"
    LoadResStrings frmMenu, txtSelected & ".lan"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ''Call the exit Sub
    ''cmdExit_Click
End Sub

Private Sub lstLanguages_Click()
    ''Set the Selected Language
    txtSelected = lstLanguages.Text
    ''Load the Selected Language
    LoadResStrings frmMain, txtSelected & ".lan"
    ''Load also in the other Form
    LoadResStrings frmMenu, txtSelected & ".lan"
End Sub
