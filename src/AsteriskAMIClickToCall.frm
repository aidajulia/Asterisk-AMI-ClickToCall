VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form AsteriskAMIClickToCall 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asterisk AMI ClickToCall"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5970
   Icon            =   "AsteriskAMIClickToCall.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton GUIClear 
      Caption         =   "Clear the Graphical Interface"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   4600
      Width           =   5490
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click-To-Call Mode "
      Height          =   2200
      Left            =   240
      TabIndex        =   3
      Top             =   300
      Width           =   5500
      Begin VB.OptionButton optMode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mode 1 - Live Callback Mode"
         ForeColor       =   &H00000080&
         Height          =   200
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   380
         Value           =   -1  'True
         Width           =   4815
      End
      Begin VB.OptionButton optMode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mode 2 - Automated Message Mode"
         ForeColor       =   &H00000080&
         Height          =   200
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1380
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Receive a prerecorded message during the call with instructions or predefined information "
         Height          =   400
         Left            =   525
         TabIndex        =   8
         Top             =   1620
         Width           =   4605
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"AsteriskAMIClickToCall.frx":0442
         Height          =   600
         Left            =   525
         TabIndex        =   7
         Top             =   620
         Width           =   4605
      End
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   1800
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   20000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "AsteriskAMIClickToCall.frx":04E0
      ToolTipText     =   "Rx Info from AMI!"
      Top             =   2800
      Width           =   5740
   End
   Begin VB.PictureBox LedConnect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   220
      Index           =   0
      Left            =   5680
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   1
      ToolTipText     =   " Connection NO Active! "
      Top             =   80
      Visible         =   0   'False
      Width           =   200
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5400
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Init Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   5120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   5160
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   4440
      Tag             =   "Asterisk Client - AMI"
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "10.10.11.5"
      RemotePort      =   5038
      LocalPort       =   700
   End
End
Attribute VB_Name = "AsteriskAMIClickToCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' https://github.com/aidajulia/Asterisk-AMI-ClickToCall.git
' More info: Luis Fernando Echeverri Lozano (lfeche@hotmail.com)

Option Explicit

Private Sub Form_Load()
    
    On Error GoTo ErrorHandler
    
    ' Configure Timer to check every 1 second
    Timer1.Interval = 1000
    Timer1.Enabled = False
    
    LoadConfig
    
    Winsock1.LocalPort = CStr(CnfApp.WebServerLocalPort)
    Winsock1.Listen
    
    Exit Sub
    
ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Private Sub Command1_Click()
    
    On Error GoTo ErrorHandler
    
    ' Open default browser to VB6-served page
    Shell "cmd.exe /c start http://localhost:" & CnfApp.WebServerLocalPort, vbHide
    
    ' Start continuous monitoring
    Timer1.Enabled = True
    Command1.Enabled = False ' Disable button to prevent multiple starts
    
    Exit Sub
    
ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Private Sub GUIClear_Click()
    'Clear GUI
    AsteriskAMIClickToCall.txtInfo.Text = ""
    StreamFromAMI = ""
    
End Sub

Private Sub optMode_Click(Index As Integer)
    'Change Click-To-Call Mode
    If Index = 0 Then
        CTCMode = 0
        'CnfApp.PartyA_Mode1 = "SIP/103"                ' Default Agent "SIP/103" in PBX-Demo1
        CnfApp.ModeDialer = "Exten"
        CnfApp.TypeDialer = "Callback"
        
    ElseIf Index = 1 Then
        CTCMode = 1
        'CnfApp.PartyA_Mode2 = "Data: demo-thanks"      ' Message to Playback for App, "demo-thanks"
        CnfApp.ModeDialer = "App"
        CnfApp.TypeDialer = "Message"
        
    End If

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    
    On Error GoTo ErrorHandler
    
    ' Accept incoming connection
    If Winsock1.State <> sckClosed Then Winsock1.Close
    Winsock1.Accept requestID
    
    Exit Sub
    
ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Dim phoneNumber As String
Dim htmlResponse As String
Dim fullResponse As String

    On Error GoTo ErrorHandler
    
    Winsock1.GetData strData
    
    If InStr(strData, "GET /") > 0 Then
        htmlResponse = "<!DOCTYPE html>" & vbCrLf & _
                       "<html><head><title>ClickToCall</title></head>" & vbCrLf & _
                       "<body><h3>Enter phone number to call:</h3>" & vbCrLf & _
                       "<form method='POST' action='/submit'>" & vbCrLf & _
                       "<input type='text' name='phoneNumber' placeholder='e.g: 1234567890'>" & vbCrLf & _
                       "<input type='submit' value='Submit'>" & vbCrLf & _
                       "</form></body></html>"
        fullResponse = "HTTP/1.1 200 OK" & vbCrLf & _
                       "Content-Type: text/html" & vbCrLf & _
                       "Content-Length: " & Len(htmlResponse) & vbCrLf & _
                       "Connection: close" & vbCrLf & _
                       vbCrLf & _
                       htmlResponse
        Winsock1.SendData fullResponse
        DoEvents    ' Allow data to be sent
        Winsock1.Close
        Winsock1.Listen
    ElseIf InStr(strData, "POST /submit") > 0 Then
        phoneNumber = Mid(strData, InStr(strData, "phoneNumber=") + 12)
        If InStr(phoneNumber, vbCrLf) > 0 Then
            phoneNumber = Left(phoneNumber, InStr(phoneNumber, vbCrLf) - 1)
        'Else
        End If
        '
        ' Debug
        ' MsgBox "Obtained number: " & phoneNumber, vbInformation
        CallToMake = phoneNumber
        ' Here we can send data to buffer
        ' Simple Queued calls
        If Trim(CallToMake) <> "" Then
            '
            ' Make a call / without "Sanitizenumbers(CallToMake)"
            ' Public Function ActionDlrMini(PartyA As String, PartyB As String, CallRoute As String, CallerIDLabel As String, ContextLabel As String, Optional ModeDialer As String, Optional TypeDialer As String) As String
            ' MakeCall = ActionDlrMini(PartyA, PartyB, CallRoute, CallerIDLabel, ContextLabel, ModeDialer, TypeDialer)
            '
            ' Defined Click-To-Call "Mode#" and "TypeDialer" values of Function "ActionDlrMini()"
            If CTCMode = 0 Then
                CnfApp.PartyB = CallToMake
                'CnfApp.PartyA_Mode1 = "SIP/103"                ' Default Agent "SIP/103" in PBX-Demo1
                CnfApp.ModeDialer = "Exten"
                CnfApp.TypeDialer = "Callback"
                CnfApp.CallerIDLabel = "Dialler Calling to "
                MakeCall = ActionDlrMini(CnfApp.PartyA_Mode1, CnfApp.PartyB, CnfApp.CallRoute, CnfApp.CallerIDLabel, CnfApp.ContextLabel, CnfApp.ModeDialer, CnfApp.TypeDialer)
                
            ElseIf CTCMode = 1 Then
                CnfApp.PartyB = CallToMake
                'CnfApp.PartyA_Mode2 = "Data: demo-thanks"      ' Message to Playback for App, "demo-thanks"
                CnfApp.ModeDialer = "App"
                CnfApp.TypeDialer = "Message"
                CnfApp.CallerIDLabel = "App/Playback-"
                MakeCall = ActionDlrMini(CnfApp.PartyA_Mode2, CnfApp.PartyB, CnfApp.CallRoute, CnfApp.CallerIDLabel, CnfApp.ContextLabel, CnfApp.ModeDialer, CnfApp.TypeDialer)
                
            End If
            
        'Else
            ' nothing to send
        End If
        '
        ' Save or Show the last ActionID for Call
        ' MakeCall=ActionID
        '
        
        htmlResponse = "<html><body><h3>Number submitted. Enter another:</h3>" & _
                       "<a href='/'>Return</a></body></html>"
        fullResponse = "HTTP/1.1 200 OK" & vbCrLf & _
                       "Content-Type: text/html" & vbCrLf & _
                       "Content-Length: " & Len(htmlResponse) & vbCrLf & _
                       "Connection: close" & vbCrLf & _
                       vbCrLf & _
                       htmlResponse
        Winsock1.SendData fullResponse
        DoEvents    ' Allow data to be sent
        Winsock1.Close
        Winsock1.Listen
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Private Sub Timer1_Timer()

    ' Here - All temporal processes - how a scheduler handles them
    
    On Error GoTo ErrorHandler
    
    ' Keep socket listening continuously
    If Winsock1.State <> sckListening Then
        Winsock1.Close
        Winsock1.Listen
    End If
    '
    
    '
    ' Other proccess for Connection with VoIP Asterisk
    If (AsteriskAMIClickToCall.Winsock2.State = sckClosed) Then
        '
        AsteriskAMIClickToCall.LedConnect(0).Visible = False
        AsteriskAMIClickToCall.LedConnect(0).BackColor = &H8000000F
        SessionAMI = 0
        LedConnect(0).ToolTipText = " Not Connected! "
        '
        If (CnfApp.LocalPort = 0) Or (CnfApp.LocalPort > 799) Then
            CnfApp.LocalPort = 700
        Else
            CnfApp.LocalPort = CnfApp.LocalPort + 1
        End If
        '
        ' Socket Close, then set Setup and connect
        AsteriskAMIClickToCall.Winsock2.RemotePort = CnfApp.RemotePort
        AsteriskAMIClickToCall.Winsock2.RemoteHost = CnfApp.RemoteHost
        AsteriskAMIClickToCall.Winsock2.LocalPort = CnfApp.LocalPort
        ' Invoke the Connect method to initiate a connection.
        AsteriskAMIClickToCall.Winsock2.Connect
        ' Connected without session
        
    ElseIf (AsteriskAMIClickToCall.Winsock2.State = sckConnected) And (SessionAMI = 1) Then
        ' Handshaking for Session
        AsteriskAMIClickToCall.Winsock2.SendData CnfApp.Word1 & vbCrLf & CnfApp.Word2 & vbCrLf & CnfApp.Word3 & vbCrLf & CnfApp.Word4 & vbCrLf & "" & vbCrLf & ""
        
        ' Update Status of connection
        AsteriskAMIClickToCall.LedConnect(0).Visible = True
        AsteriskAMIClickToCall.LedConnect(0).BackColor = &H4133EC
        LedConnect(0).ToolTipText = " Connected! "
        
    ElseIf (AsteriskAMIClickToCall.Winsock2.State = sckConnected) And (SessionAMI = 2) Then
        ' Update Status of connection and session
        AsteriskAMIClickToCall.LedConnect(0).Visible = True
        AsteriskAMIClickToCall.LedConnect(0).BackColor = &HE698&
        LedConnect(0).ToolTipText = " On Session "
        
        ' Send Calls from buffer
        If (Trim(CallCmdToAMI) <> "") Then
            AsteriskAMIClickToCall.Winsock2.SendData CallCmdToAMI
            
            ' Clear the buffer
            CallCmdToAMI = ""
            
            ' Update Status of Sending Call
            AsteriskAMIClickToCall.LedConnect(0).Visible = True
            AsteriskAMIClickToCall.LedConnect(0).BackColor = &HFFFF&
            LedConnect(0).ToolTipText = " Call in progress "
            
        Else
            ' Nothing to send
        End If
        
    ElseIf (AsteriskAMIClickToCall.Winsock2.State = sckError) Then
        ' In error, close the socket
        AsteriskAMIClickToCall.Winsock2.Close
        AsteriskAMIClickToCall.LedConnect(0).Visible = False
        AsteriskAMIClickToCall.LedConnect(0).BackColor = &H8000000F
        SessionAMI = 0
        LedConnect(0).ToolTipText = " Not Connected! "
    
    ElseIf (AsteriskAMIClickToCall.Winsock2.State = sckConnected) Then
        '
        AsteriskAMIClickToCall.LedConnect(0).Visible = True
        AsteriskAMIClickToCall.LedConnect(0).BackColor = &H4133EC
        SessionAMI = 1
        LedConnect(0).ToolTipText = " Connected! "
        
    Else
        ' AsteriskAMIClickToCall.Winsock2.?
    End If
    '
    
    '
    ' Update GUI
    ' Rx-Downstream
    If StreamFromAMI <> "" Then
        AsteriskAMIClickToCall.txtInfo.Text = StreamFromAMI
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Private Sub Winsock2_Close()
    
    On Error GoTo ErrorHandler
    
    AsteriskAMIClickToCall.Winsock2.Close
    AsteriskAMIClickToCall.LedConnect(0).Visible = False
    AsteriskAMIClickToCall.LedConnect(0).BackColor = &H8000000F
    SessionAMI = 0
    LedConnect(0).ToolTipText = " Not Connected! "
    
    Exit Sub

ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Private Sub Winsock2_Connect()
    
    On Error GoTo ErrorHandler
    
    AsteriskAMIClickToCall.LedConnect(0).Visible = True
    AsteriskAMIClickToCall.LedConnect(0).BackColor = &H4133EC
    SessionAMI = 1
    LedConnect(0).ToolTipText = " Connected! "
    
    Exit Sub

ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
    
    On Error GoTo ErrorHandler
    
    Exit Sub
        
ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Public Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Dim Sincro As Integer

    On Error GoTo ErrorHandler
    
    strData = ""
    If AsteriskAMIClickToCall.Winsock2.State = sckConnected Then
        AsteriskAMIClickToCall.Winsock2.GetData strData, vbString
        If (InStr(strData, "Message: Authentication accepted") > 0) Then
            AsteriskAMIClickToCall.LedConnect(0).Visible = True
            AsteriskAMIClickToCall.LedConnect(0).BackColor = &HE698&
            SessionAMI = 2
            LedConnect(0).ToolTipText = " On Session! "
        End If
        
        ' Rx-Downstream
        If Len(StreamFromAMI) > 10000 Then
            StreamFromAMI = Mid(StreamFromAMI, 5000)
            Sincro = InStr(StreamFromAMI, "Event: ")
            If Sincro > 0 Then
                StreamFromAMI = Mid(StreamFromAMI, Sincro)
            End If
        End If
        StreamFromAMI = StreamFromAMI & strData
        ' Later, GUI Display
        ' txtInfo.Text = StreamFromAMI
        
    End If
    
    Exit Sub
        
ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    On Error GoTo ErrorHandler
    
    AsteriskAMIClickToCall.LedConnect(0).Visible = False
    AsteriskAMIClickToCall.LedConnect(0).BackColor = &H8000000F
    SessionAMI = 0
    LedConnect(0).ToolTipText = " Not Connected! "
    
    Exit Sub
    
ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Private Sub Winsock2_SendComplete()
    
    On Error GoTo ErrorHandler
    
    Exit Sub
        
ErrorHandler:
    ' Oops!
    Err.Clear
End Sub

Private Sub Winsock2_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    
    On Error GoTo ErrorHandler

    Exit Sub

ErrorHandler:
    ' Oops!
    Err.Clear
End Sub
