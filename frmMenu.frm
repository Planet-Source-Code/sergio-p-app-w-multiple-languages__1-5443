VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3900
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1000"
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "3000"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "3001"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3045
      Left            =   120
      TabIndex        =   0
      Tag             =   "1006"
      Top             =   690
      Width           =   4275
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   2385
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   4207
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "tabSample"
               Object.Tag             =   "2002"
               Object.ToolTipText     =   "ToolTipText"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "tabSample"
               Object.Tag             =   "2003"
               Object.ToolTipText     =   "ToolTipText"
               ImageVarType    =   2
            EndProperty
         EndProperty
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
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'' You can also load the string from the Form_Load event:
''
'' Private Sub Form_Load()
    '' Load the Language
    '' LoadResStrings frmMenu, frmMain.txtSelected & ".lan"
'' End Sub
