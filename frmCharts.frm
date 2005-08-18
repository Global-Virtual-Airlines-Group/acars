VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmCharts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Approach Charts"
   ClientHeight    =   11565
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   771
   ScaleMode       =   0  'User
   ScaleWidth      =   776.987
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT CHART"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   105
      Width           =   1455
   End
   Begin VB.ComboBox cboChart 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      Text            =   "cboChart"
      Top             =   120
      Width           =   3255
   End
   Begin VB.ComboBox cboAirport 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCharts.frx":0000
      Left            =   840
      List            =   "frmCharts.frx":0007
      TabIndex        =   1
      Text            =   "cboAirport"
      Top             =   120
      Width           =   2535
   End
   Begin SHDocVwCtl.WebBrowser brwChart 
      Height          =   10800
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   10455
      ExtentX         =   18441
      ExtentY         =   19050
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblChart 
      Alignment       =   1  'Right Justify
      Caption         =   "Chart Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label lblAirport 
      Alignment       =   1  'Right Justify
      Caption         =   "Airport"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ApproachCharts As Charts

Private Sub cboChart_Click()
    Dim chartID As Integer
    Dim hostName As String
    
    'Get the chart ID
    chartID = ApproachCharts.ChartIDs(cboChart.ListIndex)
    
    'DEBUG REDIRECTION
    hostName = config.ACARSHost
    If (hostName = "polaris.sce.net") Then hostName = "dva2006.deltava.org"
    
    'Fire up the web browser
    brwChart.Navigate "http://" + frmMain.txtPilotID.Text + ":" + frmMain.txtPassword.Text + "@" + _
        hostName + "/charts/0x" + Hex(chartID)
    ShowMessage "Loading Approach Chart " + ApproachCharts.ChartNames(cboChart.ListIndex) + _
        " for " + ApproachCharts.AirportCode, ACARSTEXTCOLOR
End Sub

Private Sub cmdPrint_Click()
    brwChart.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub
