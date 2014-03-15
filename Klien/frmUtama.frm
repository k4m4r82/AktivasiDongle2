VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCekDongle 
   Caption         =   "Form1"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Left            =   2280
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCekDongle"
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

Private Const LOCAL_PORT    As Long = 1007

Private Function startConnect(ByVal ipServer As String) As Boolean
    On Error Resume Next
    
    If Socket.State <> sckClosed Then Socket.Close ' close existing connection
    Call Socket.Connect(ipServer, LOCAL_PORT)
    With Socket
        Do While .State <> sckConnected
            DoEvents
            If .State = sckError Then Exit Function
        Loop
    End With
    
    startConnect = True
End Function

Private Function send(ByVal strData As String) As Boolean
    If Socket.State = sckConnected Then
        Call Socket.SendData(strData)
        DoEvents
        
    Else
        send = False
        Exit Function
    End If
   
    send = True
End Function

Private Sub Form_Load()
    If startConnect("127.0.0.1") Then
        'ingat hanya string 'reqStatusDongle' yang akan diproses oleh server
        If Not send("reqStatusDongle") Then
            MsgBox "Aplikasi Toekang Cek Doengle belum aktif", vbExclamation, "Warning"
            Unload Me
        End If
        
    Else
        MsgBox "Aplikasi Toekang Cek Doengle belum aktif", vbExclamation, "Warning"
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Socket.State <> sckClosed Then Socket.Close        ' close existing connection
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim dataMasuk   As String
    
    'On Error Resume Next
    
    Socket.GetData dataMasuk
    If dataMasuk = "true" Then 'dongle sudah terpasang dan valid
        MsgBox "dongle valid"
        'TODO : tampilkan form utama disini
    Else
        MsgBox "Maaf dongle belum terpasang, aplikasi tidak bisa dilanjutkan", vbExclamation, "Peringatan"
    End If
    
    Unload Me
End Sub

