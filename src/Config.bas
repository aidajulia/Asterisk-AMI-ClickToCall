Attribute VB_Name = "ConfigApp"
' https://github.com/aidajulia/Asterisk-AMI-ClickToCall.git
' More info: Luis Fernando Echeverri Lozano (lfeche@hotmail.com)

' Config Module (Config.bas)
Option Explicit

' =============================================
' Global Settings
' =============================================
Public MakeCall As String           ' ActionID
Public CallCmdToAMI As String       ' Call Command to AMI/ Call Buffer
Public StreamFromAMI As String      ' Rx Data From AMI
Public CallToMake As String         ' Obtained number from Web
Public CTCMode As Integer           ' Click-To-Call Mode

' Data structure needed for Windows API call (GetSystemTime)
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

' Windows API call to get system time from os
Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public UniqueActionID As SYSTEMTIME

Public Type AppConfig
    '
    'Web Server listen in port
    WebServerLocalPort As Integer   ' Set up socket as server on port 8080
    
    ' Asterisk AMI Connection
    ' Default Setup for PBX-Demo1, FreePBX16, Asterisk-13/16/18 (See Proxmox Backups)
    ' https://drive.google.com/file/d/1v4sGf8kRSnP3Nu8SmZDS0qMfFtQTa_bK/view?usp=sharing
    RemoteHost As String            ' Asterisk PBX-Demo1 Server IP Address,          "10.10.11.5"
    RemotePort As Integer           ' Asterisk PBX-Demo1 Server Default AMI Port,    "5038"
    LocalHost As String             ' Windows PC, Default IP Address,                "127.0.0.1"
    LocalPort As Integer            ' Windows PC, Default App Port,                  "700"
    
    ' Asterisk AMI Session          ' Default Session Values for PBX-Demo1.
    Word1 As String                 ' "Action: Login"
    Word2 As String                 ' "Username: adminag"
    Word3 As String                 ' "Secret: ipbxcntrlBWBox1AMI"
    Word4 As String                 ' "Events: On"
    
    ' Dialer Variables
    PartyA_Mode1 As String          ' Local Extension                       ' Default Agent "SIP/103" in PBX-Demo1
    PartyA_Mode2 As String          ' Message to Playback for App.          ' "Data: demo-thanks"
    PartyB As String                ' Telephone Number for Remote Client    ' Number to Call "203", in this sample this is the Extension 203 in PBX-Demo2
    CallRoute As String             ' Digits Route                          ' Trunk Access Line Code to Call "92"
    CallerIDLabel As String         ' CallerID                              ' Maybe "Dialler Calling to "
    ContextLabel As String          ' Call context                          ' Default Context "from-internal"
    ModeDialer As String            ' Dialer Mode                           ' Exten/App
    TypeDialer As String            ' Dialer Type                           ' Callback/Message
    RingTime As String              ' RingTime                              ' Ring timer before hangup (1000=1Second)
    CallTime As String              ' TIMEOUT                               ' Duration Call if present (Seconds)

End Type

Public CnfApp As AppConfig          ' INI file for App
Public SessionAMI As Integer        ' not connect=0, connect=1, session=2

Public Sub LoadConfig()
    
    On Error GoTo UseDefaults
    
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    If fso.FileExists(App.Path & "\AsteriskAMIClickToCall.ini") Then
        Set ts = fso.OpenTextFile(App.Path & "\AsteriskAMIClickToCall.ini", ForReading)
        
        ' Web Server listen in port
        CnfApp.WebServerLocalPort = CInt(ReadINI(ts, "WebServerLocalPort"))
        
        ' Asterisk AMI Connection
        CnfApp.RemoteHost = ReadINI(ts, "RemoteHost")
        CnfApp.RemotePort = CInt(ReadINI(ts, "RemotePort"))
        CnfApp.LocalHost = ReadINI(ts, "LocalHost")
        CnfApp.LocalPort = CInt(ReadINI(ts, "LocalPort"))
        
        ' Asterisk AMI Session
        CnfApp.Word1 = ReadINI(ts, "Word1")
        CnfApp.Word2 = ReadINI(ts, "Word2")
        CnfApp.Word3 = ReadINI(ts, "Word3")
        CnfApp.Word4 = ReadINI(ts, "Word4")
        
        ' Dialer Variables
        CnfApp.PartyA_Mode1 = ReadINI(ts, "PartyA_Mode1")
        CnfApp.PartyA_Mode2 = ReadINI(ts, "PartyA_Mode2")
        CnfApp.PartyB = ReadINI(ts, "PartyB")
        CnfApp.CallRoute = ReadINI(ts, "CallRoute")
        CnfApp.CallerIDLabel = ReadINI(ts, "CallerIDLabel")
        CnfApp.ContextLabel = ReadINI(ts, "ContextLabel")
        CnfApp.ModeDialer = ReadINI(ts, "ModeDialer")
        CnfApp.TypeDialer = ReadINI(ts, "TypeDialer")
        CnfApp.RingTime = ReadINI(ts, "RingTime")
        CnfApp.CallTime = ReadINI(ts, "CallTime")
        
        ts.Close
        Exit Sub
    End If
    
UseDefaults:
    ' Default Setup in first Run if no exist "AsteriskAMIClickToCall.ini"
    ' All default Setup are for PBX-Demo1, FreePBX16, Asterisk-13/16/18 (See Proxmox Backups)
    ' https://drive.google.com/file/d/1v4sGf8kRSnP3Nu8SmZDS0qMfFtQTa_bK/view?usp=sharing
    
    ' Default Setup for Server Click To Call
    CnfApp.WebServerLocalPort = 8080

    ' Default Asterisk AMI Connection
    CnfApp.RemoteHost = "10.10.11.5"
    CnfApp.RemotePort = 5038
    CnfApp.LocalHost = "127.0.0.1"
    CnfApp.LocalPort = 700
    
    ' Default Asterisk AMI Session
    CnfApp.Word1 = "Action: Login"
    CnfApp.Word2 = "Username: adminag"
    CnfApp.Word3 = "Secret: ipbxcntrlBWBox1AMI"
    CnfApp.Word4 = "Events: On"
    
    ' Default Dialer Variables for PBX-Demo1 and PBX-Demo2 Setups (See Proxmox Backups)
    CnfApp.PartyA_Mode1 = "SIP/103"                 ' Default Agent "SIP/103" in PBX-Demo1
    CnfApp.PartyA_Mode2 = "Data: demo-thanks"       ' Message to Playback for App, "demo-thanks"
    CnfApp.PartyB = "203"                           ' Extension "203" in PBX-Demo2, this is the Client in this Demo.
    CnfApp.CallRoute = "92"                         ' Digit Route "92" from PBX-Demo1 to PBX-Demo2
    CnfApp.CallerIDLabel = "Dialler Calling to "
    CnfApp.ContextLabel = "from-internal"
    CnfApp.ModeDialer = "Exten"
    CnfApp.TypeDialer = "Callback"
    CnfApp.RingTime = "12"
    CnfApp.CallTime = "120"
    
    SaveConfig
End Sub

Public Sub SaveConfig()
    On Error Resume Next
    
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = fso.OpenTextFile(App.Path & "\AsteriskAMIClickToCall.ini", ForWriting, True)
    
    ' Default Setup for Server Click To Call
    ts.WriteLine "WebServerLocalPort=" & CnfApp.WebServerLocalPort
    
    ' Connection
    ts.WriteLine "RemoteHost=" & CnfApp.RemoteHost
    ts.WriteLine "RemotePort=" & CnfApp.RemotePort
    ts.WriteLine "LocalHost=" & CnfApp.LocalHost
    ts.WriteLine "LocalPort=" & CnfApp.LocalPort
    
    ' Session
    ts.WriteLine "Word1=" & CnfApp.Word1
    ts.WriteLine "Word2=" & CnfApp.Word2
    ts.WriteLine "Word3=" & CnfApp.Word3
    ts.WriteLine "Word4=" & CnfApp.Word4
    
    ' Dialer Variables
    ts.WriteLine "PartyA_Mode1=" & CnfApp.PartyA_Mode1
    ts.WriteLine "PartyA_Mode2=" & CnfApp.PartyA_Mode2
    ts.WriteLine "PartyB=" & CnfApp.PartyB
    ts.WriteLine "CallRoute=" & CnfApp.CallRoute
    ts.WriteLine "CallerIDLabel=" & CnfApp.CallerIDLabel
    ts.WriteLine "ContextLabel=" & CnfApp.ContextLabel
    ts.WriteLine "ModeDialer=" & CnfApp.ModeDialer
    ts.WriteLine "TypeDialer=" & CnfApp.TypeDialer
    ts.WriteLine "RingTime=" & CnfApp.RingTime
    ts.WriteLine "CallTime=" & CnfApp.CallTime
    
    ts.Close
End Sub

Private Function ReadINI(ts As TextStream, key As String) As String
    Dim line As String
    
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        If Left(line, Len(key) + 1) = key & "=" Then
            ReadINI = Mid(line, Len(key) + 2)
            Exit Function
        End If
    Loop
End Function

Public Function ActionDlrMini(PartyA As String, PartyB As String, CallRoute As String, CallerIDLabel As String, ContextLabel As String, Optional ModeDialer As String, Optional TypeDialer As String) As String
Dim UAID As String
Dim ActionID As String                  ' Timestamp
Dim VarBufferAction_Dlr As String       ' Local Buffer
Dim CmdAst As String                    ' PartyA
Dim ExtraCmdAst As String               ' PartyB

    On Error GoTo ErrorHandler
    
    '
    ' UniqueActionID
    GetSystemTime UniqueActionID
    UAID = Format(Now, "YYYYMMDDHHMMSS") & Format(Trim(UniqueActionID.wMilliseconds), "000")
    
    '
    If (ModeDialer = "Exten") Or (ModeDialer = "") Then
        '
        If (TypeDialer = "Callback") Or (TypeDialer = "") Then
            '
            ' Unique ID for Command
            ActionID = UAID
            CmdAst = Trim(PartyA)
            ExtraCmdAst = Trim(CallRoute) & Trim(PartyB)
            CallerIDLabel = CallerIDLabel & ExtraCmdAst
            '
            ' Default operation in this Mode
            CmdAst = "Action: Originate" & vbCrLf & "ActionID: " & ActionID & vbCrLf & "Channel: " & CmdAst & vbCrLf & "Context: " & ContextLabel & vbCrLf & "Exten: " & ExtraCmdAst & vbCrLf & "Priority: 1" & vbCrLf & "Callerid: " & CallerIDLabel & vbCrLf & "Account: " & Format(Now, "YYMMDD") & Mid(ActionID, 9) & vbCrLf & "Timeout: " & Val(CnfApp.RingTime * 1000) & vbCrLf & "Variable: TIMEOUT(absolute)=" & Val(CnfApp.CallTime) & vbCrLf & "Async: true" & vbCrLf & vbCrLf
            '
            ' SAMPLE OK!!!
            '===================================
            '"Action: Originate
            'ActionID: 20250421173030261
            'Channel: SIP/103
            'Context: from-internal
            'Exten: 92203
            'Priority: 1
            'Callerid: Dialler Calling to 92203
            'Account: 250421173030261
            'Timeout: 12000
            'Variable: TIMEOUT(absolute)=120
            'Async: true
            '
            '"
            '===================================
            '
            
            '
            ' Local Buffer
            VarBufferAction_Dlr = CmdAst
            '
            ' Insert to Command Queue
            If (Trim(CallCmdToAMI) = "") And (VarBufferAction_Dlr <> "") Then
                CallCmdToAMI = VarBufferAction_Dlr
                VarBufferAction_Dlr = ""
            ElseIf (Trim(CallCmdToAMI) <> "") And (VarBufferAction_Dlr <> "") Then
                CallCmdToAMI = CallCmdToAMI & VarBufferAction_Dlr
                VarBufferAction_Dlr = ""
            End If
        
        ElseIf (ModeDialer = "Queue") Or (ModeDialer = "PlayRecrd") Then
            ' Other Modes!
        
        End If
        
    ElseIf ModeDialer = "App" Then
        '
        If (TypeDialer = "Message") Or (TypeDialer = "") Then
            '
            ' Unique ID for Command
            ActionID = UAID
            CmdAst = Trim(PartyA)
            ExtraCmdAst = Trim(CallRoute) & Trim(PartyB)
            CallerIDLabel = CallerIDLabel & ExtraCmdAst
            '
            ' Default operation in this Mode
            CmdAst = "Action: Originate" & vbCrLf & "ActionID: " & ActionID & vbCrLf & "Channel: " & "Local/" & ExtraCmdAst & "@from-internal" & vbCrLf & "Application: Playback" & vbCrLf & CmdAst & vbCrLf & "Callerid: " & CallerIDLabel & vbCrLf & "Context: from-internal" & vbCrLf & "Exten: s" & vbCrLf & "Account: " & Format(Now, "YYMMDD") & Mid(ActionID, 9) & vbCrLf & "Timeout: " & Val(CnfApp.RingTime * 1000) & vbCrLf & "Variable: TIMEOUT(absolute)=" & Val(CnfApp.CallTime) & vbCrLf & "Priority: 1" & vbCrLf & "Async: true" & vbCrLf & vbCrLf
            '
            ' SAMPLE OK!!!
            '===================================
            '"Action: Originate
            'ActionID: 20250421190732273
            'Channel: Local/92203@from-internal
            'Application: Playback
            'Data: demo-thanks
            'Callerid: App/Playback-92203
            'Context: from-internal
            'Exten: s
            'Account: 250421190732273
            'Timeout: 12000
            'Variable: TIMEOUT(absolute)=120
            'Priority: 1
            'Async: true
            '
            '"
            '===================================
            '
            
            '
            ' Local Buffer
            VarBufferAction_Dlr = CmdAst
            '
            ' Insert to Command Queue
            If (Trim(CallCmdToAMI) = "") And (VarBufferAction_Dlr <> "") Then
                CallCmdToAMI = VarBufferAction_Dlr
                VarBufferAction_Dlr = ""
            ElseIf (Trim(CallCmdToAMI) <> "") And (VarBufferAction_Dlr <> "") Then
                CallCmdToAMI = CallCmdToAMI & VarBufferAction_Dlr
                VarBufferAction_Dlr = ""
            End If
        
        Else
            ' Other Modes!
        
        End If
        
    End If
    '
    
    ActionDlrMini = ActionID
    
    Exit Function
    
ErrorHandler:
    ActionDlrMini = ""
    ' Oops!
    Err.Clear
End Function
