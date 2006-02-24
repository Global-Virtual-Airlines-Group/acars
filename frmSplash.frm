VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Splash Screen"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmSplash.frx":08CA
   ScaleHeight     =   6750
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing something or other..."
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   3495
   End
   Begin VB.Label VersionLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Build 0.1.42"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   6360
      Width           =   5985
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isClicked As Boolean

Private Sub Form_Click()
    isClicked = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    isClicked = True
End Sub

Private Sub Form_Load()
    VersionLabel.Caption = "Version " + CStr(App.Major) + "." + CStr(App.Minor) + _
        " (Build " + Format(App.Revision, "000") + ")"
End Sub

Public Sub SetProgressLabel(ProgressInfo As String)
    lblProgress.Caption = ProgressInfo + " ..."
    DoEvents
    Sleep 150
End Sub

Public Sub ClearProgressLabel()
    lblProgress.Caption = ""
    DoEvents
End Sub
