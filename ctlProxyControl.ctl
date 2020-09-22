VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ctlProxyControl 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1335
   ScaleWidth      =   2565
   Begin MSWinsockLib.Winsock RemoteSocks 
      Index           =   0
      Left            =   600
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LocalSocks 
      Index           =   0
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "ctlProxyControl.ctx":0000
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblProxyServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SOCKS Proxy Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape shpLogo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "ctlProxyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event SocketEvent(Index As Integer, Description As String)

Private Socks() As SockSocket
Private Type SockSocket
    strConnectionData As String
    strHostAddress As String
    intPortNumber As Integer
    blnFirstPacket As Boolean
End Type

Private Function Word(ByVal lngVal As Long) As String
    Dim Lo As Long
    Dim Hi As Long
    Lo = Fix(lngVal / 256)
    Hi = lngVal Mod 256
    Word = Chr(Lo) & Chr(Hi)
End Function

Private Function GetWord(ByVal strVal As String) As Long
    Dim Lo As Long
    Dim Hi As Long
    Lo = Asc(Mid(strVal, 1, 1))
    Hi = Asc(Mid(strVal, 2, 1))
    GetWord = (Lo * 256) + Hi
End Function

'--------------------------------------------------------------------------------------------------------------
Private Function CreateSock() As Integer
    Dim i As Integer
    For i = 1 To UBound(Socks)
        If LocalSocks(i).State <> sckConnected Then
            CreateSock = i
            Exit Function
        End If
    Next i
    ReDim Preserve Socks(0 To UBound(Socks) + 1)
    CreateSock = UBound(Socks)
    Load LocalSocks(CreateSock)
    Load RemoteSocks(CreateSock)
End Function
 
Private Function ResetSock(Index As Integer)
    LocalSocks(Index).Close
    RemoteSocks(Index).Close
    Socks(Index).strHostAddress = ""
    Socks(Index).intPortNumber = 0
    Socks(Index).blnFirstPacket = True
End Function

Private Function CloseSock(Index As Integer)
    LocalSocks(Index).Close
    RemoteSocks(Index).Close
    ResetSock Index
End Function

Private Function CheckSock(Index As Integer)
    If LocalSocks(Index).State <> sckConnected Or RemoteSocks(Index).State <> sckConnected Then
        CloseSock Index
    End If
End Function

Private Function SendSock(SockType As Integer, Index As Integer, Data As String)
    If LocalSocks(Index).State = sckConnected And RemoteSocks(Index).State = sckConnected Then
        If SockType = 0 Then
            LocalSocks(Index).SendData Data
        Else
            RemoteSocks(Index).SendData Data
        End If
    Else
        RaiseEvent SocketEvent(Index, "Connection was unexpectedly closed.")
        CloseSock Index
    End If
End Function
'--------------------------------------------------------------------------------------------------------------

Private Sub LocalSocks_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i As Integer
    i = CreateSock
    ResetSock i
    LocalSocks(i).Close
    LocalSocks(i).Accept requestID
    RaiseEvent SocketEvent(i, "Connection received. (" & LocalSocks(Index).RemoteHostIP & ")")
End Sub

Private Sub LocalSocks_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo DeepStuff:
    Dim strData As String
    Dim lngLength As Long
    If Socks(Index).blnFirstPacket = True Then
        LocalSocks(Index).GetData strData
        If Mid(strData, 1, 1) = Chr(4) And Mid(strData, 2, 1) = Chr(1) Then
            RaiseEvent SocketEvent(Index, "Creating outgoing socket. (" & Asc(Mid(strData, 5, 1)) & "." & Asc(Mid(strData, 6, 1)) & "." & Asc(Mid(strData, 7, 1)) & "." & Asc(Mid(strData, 8, 1)) & ")")
            Socks(Index).blnFirstPacket = False
            Socks(Index).intPortNumber = GetWord(Mid(strData, 3, 2))
            Socks(Index).strHostAddress = Asc(Mid(strData, 5, 1)) & "." & Asc(Mid(strData, 6, 1)) & "." & Asc(Mid(strData, 7, 1)) & "." & Asc(Mid(strData, 8, 1))
            Socks(Index).strConnectionData = Mid(strData, 3, 6)
            RemoteSocks(Index).Close
            RemoteSocks(Index).Connect Socks(Index).strHostAddress, Socks(Index).intPortNumber
        End If
    Else
        CheckSock Index
        LocalSocks(Index).GetData strData
        SendSock 1, Index, strData
    End If
    Exit Sub
DeepStuff:
    RaiseEvent SocketEvent(Index, "Local Socking Failure: " & Err.Description)
End Sub

Private Sub LocalSocks_Close(Index As Integer)
    RaiseEvent SocketEvent(Index, "Local Socket Closed")
    RemoteSocks(Index).Close
    ResetSock Index
End Sub

Private Sub RemoteSocks_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent SocketEvent(Index, "Socket Error: " & Description & " (" & RemoteSocks(Index).RemoteHostIP & ")")
    If Socks(Index).blnFirstPacket = True Then
        LocalSocks(Index).SendData Chr(4) & Chr(91) & Socks(Index).strConnectionData
    End If
    CloseSock Index
End Sub

Private Sub RemoteSocks_Connect(Index As Integer)
    CheckSock Index
    SendSock 0, Index, Chr(4) & Chr(90) & Socks(Index).strConnectionData
End Sub

Private Sub RemoteSocks_Close(Index As Integer)
    RaiseEvent SocketEvent(Index, "Remote Socket Closed")
    'LocalSocks(Index).Close
    'ResetSock Index
End Sub

Private Sub RemoteSocks_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo CrappedOut
    Dim strData As String
    Dim lngLength As Long
    CheckSock Index
    RemoteSocks(Index).GetData strData
    SendSock 0, Index, strData
    Exit Sub
CrappedOut:
    RaiseEvent SocketEvent(Index, "Remote Socking Failure: " & Err.Description)
End Sub

Private Sub UserControl_Initialize()
    ReDim Preserve Socks(0)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = shpLogo.Height
    UserControl.Width = shpLogo.Width
End Sub

Public Sub StartListen(Port As Integer)
    LocalSocks(0).Close
    LocalSocks(0).LocalPort = Port
    LocalSocks(0).Listen
End Sub

Public Sub StopListen()
    Dim i As Integer
    For i = 1 To LocalSocks.UBound
        LocalSocks(i).Close
        Unload LocalSocks(i)
        RemoteSocks(i).Close
        Unload RemoteSocks(i)
    Next i
    ReDim Preserve Socks(0)
    LocalSocks(0).Close
    RemoteSocks(0).Close
End Sub

