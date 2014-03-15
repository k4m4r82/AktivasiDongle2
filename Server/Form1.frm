VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Toekang Cek Doengle Ver. 0.0.0.1"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   2640
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   " [ Info Server ] "
      Height          =   1150
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      Begin VB.Label lblStatus 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   810
         Width           =   3135
      End
      Begin VB.Label lblIP 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label lblHostName 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   555
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " [ Info ] "
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1395
      Width           =   4695
      Begin VB.Label lblInfo 
         Caption         =   "Label1"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit

Private Const LOCAL_PORT As Long = 1007

Private Function startListening(ByVal localPort As Long) As Boolean
    'On Error GoTo errHandle
    
    If localPort > 0 Then
        'If the socket is already listening, and it's listening on the same port, don't bother restarting it.
        If (Socket(0).State <> sckListening) Or (Socket(0).localPort <> localPort) Then
            With Socket(0)
                .Close
                .localPort = localPort
                .Listen
            End With
        End If
        
        'Return true, since the server is now listening for clients.
        startListening = True
   End If
   
   Exit Function
errHandle:
   startListening = False
End Function

Private Sub startServer()
    If startListening(LOCAL_PORT) Then
        lblStatus.Caption = "Status Listening : ON"
    Else
        lblStatus.Caption = "Status Listening : OFF"
    End If
End Sub

Private Function send(ByVal lngIndex As Long, ByVal strData As String) As Boolean
    If Socket(lngIndex).State = sckConnected Then
        Call Socket(lngIndex).SendData(strData)
        DoEvents
        
    Else
        send = False
        Exit Function
    End If
   
    send = True
End Function

Private Sub Form_Load()
    lblIP.Caption = "IP : " & Socket(0).LocalIP
    lblHostName.Caption = "Host Name : " & Socket(0).LocalHostName
    
    Call startServer
    
    lblInfo.Caption = "Aplikasi ini akan dijalankan setiap startup. " & _
                      "Tugasnya hanya menerima request dari klien " & _
                      "dan mengembalikan status dongle=true|false"
End Sub

Private Sub Socket_Close(Index As Integer)
    ' Close the socket and raise the event to the parent.
    Call Socket(Index).Close
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i           As Long
    Dim j           As Long
    
    'On Error GoTo errHandle
    
    ' We shouldn't get ConnectionRequests on any other socket than the listener
    ' (index 0), but check anyway. Also check that we're not going to exceed
    ' the MaxClients property.
    If (Index = 0) Then
        ' Check to see if we've got any sockets that are free.
        For i = 1 To Socket.UBound
            If Socket(i).State = sckClosed Or Socket(i).State = sckClosing Then
                j = i
                Exit For
            End If
        Next i
      
        ' If we don't have any free sockets, load another on the array.
        If (j = 0) Then
            Call Load(Socket(Socket.UBound + 1))
            j = Socket.Count - 1
        End If
        
        ' With the selected socket, reset it and accept the new connection.
        With Socket(j)
            Call .Close
            Call .Accept(requestID)
        End With
    End If
    
    Exit Sub
    '
errHandle:
    ' Close the Winsock that caused the error.
    Call Socket(0).Close
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim ret             As Boolean
    
    Dim strData         As String
    Dim statusDongle    As String
    
    
    On Error GoTo errHandle
    
    ' Grab the data from the specified Winsock object, and pass it to the parent.
    Call Socket(Index).GetData(strData)
    DoEvents
    
    'hanya data dengan string 'reqStatusDongle' yg akan diproses
    If strData = "reqStatusDongle" Then
        If isValidDongle Then
            statusDongle = "true"
        Else
            statusDongle = "false"
        End If
        
        ret = send(Index, statusDongle) 'kirim status dongle ke klien
    End If
    
    Exit Sub
errHandle:
   Call Socket(Index).Close
End Sub

Private Sub Socket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call Socket(Index).Close
End Sub
