VERSION 5.00
Begin VB.Form frmProxyServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Personal Socks4 Proxy Server"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "frmProxyServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Server Control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton cmdProxyControl 
         Caption         =   "Start Proxy Server"
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
         Left            =   2400
         TabIndex        =   5
         Top             =   300
         Width           =   2055
      End
      Begin VB.TextBox txtPort 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "1080"
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Server Port:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin prjProxyServer.ctlProxyControl objProxy 
      Left            =   240
      Top             =   1440
      _extentx        =   4048
      _extenty        =   873
   End
   Begin VB.ListBox lstEvents 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1815
      ItemData        =   "frmProxyServer.frx":1042
      Left            =   120
      List            =   "frmProxyServer.frx":1044
      TabIndex        =   0
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lblEventLog 
      Caption         =   "Event Log:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4695
   End
End
Attribute VB_Name = "frmProxyServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProxyControl_Click()
    If Not IsNumeric(txtPort) Then
        MsgBox "Invalid port number specified.", vbCritical, "Error"
        Exit Sub
    End If
    Select Case cmdProxyControl.Caption
        Case "Start Proxy Server"
            lstEvents.AddItem "[" & Format(Time, "H:MM:SS AM/PM") & "]: Proxy Server Started..."
            objProxy.StartListen txtPort.Text
            cmdProxyControl.Caption = "Stop Proxy Server"
        Case "Stop Proxy Server"
            lstEvents.AddItem "[" & Format(Time, "H:MM:SS AM/PM") & "]: Proxy Server Stopped..."
            objProxy.StopListen
            cmdProxyControl.Caption = "Start Proxy Server"
    End Select
End Sub

Private Sub objProxy_SocketEvent(Index As Integer, Description As String)
    lstEvents.AddItem "[" & Format(Time, "H:MM:SS AM/PM") & "][" & Index & "]: " & Description
End Sub
