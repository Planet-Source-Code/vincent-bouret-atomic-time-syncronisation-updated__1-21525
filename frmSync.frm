VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSync 
   Caption         =   "Atomic Time Syncronisation"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSys 
      Caption         =   "Auto-Update System Time"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   2400
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock wskAtom 
      Left            =   2640
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
      RemotePort      =   13
   End
   Begin VB.CommandButton cMajSys 
      Caption         =   "Update Local Clock"
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
      Left            =   480
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cSyncro 
      Caption         =   "Syncronise"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cMaj 
      Caption         =   "Retrieve Atomic Time"
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
      Left            =   3120
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Timer timeRemote 
      Interval        =   100
      Left            =   4920
      Top             =   1320
   End
   Begin VB.Timer timeLocal 
      Interval        =   100
      Left            =   120
      Top             =   1320
   End
   Begin VB.ComboBox cmbGMT 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox cmbServer 
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   200
      Width           =   4575
   End
   Begin VB.Label lAtomicTime 
      Alignment       =   2  'Center
      Caption         =   "00:00:00"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lInfo2 
      Alignment       =   2  'Center
      Caption         =   "Atomic Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lInfo1 
      Alignment       =   2  'Center
      Caption         =   "Local Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lLocalTime 
      Alignment       =   2  'Center
      Caption         =   "00:00:00"
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
      Left            =   960
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lFuseau 
      Alignment       =   1  'Right Justify
      Caption         =   "GMT Decal:"
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
      Left            =   960
      TabIndex        =   2
      Top             =   645
      Width           =   1215
   End
   Begin VB.Label lServeur 
      Alignment       =   1  'Right Justify
      Caption         =   "Server:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JJJJJ YR-MO-DA HH:MM:SS TT L H msADV UTC(NIST) OTM

Const RemPort = 13
Dim ServerName(11) As String
Dim ServerIp(11) As String
Dim ServerLocation(11) As String
Dim sAtomicTime As String
Dim sBaseTime As String
Dim LocSyncroTime As Single
Dim SyncroTime As Single


Private Sub cMaj_Click()
wskAtom.Close
wskAtom.RemoteHost = ServerName(cmbServer.ListIndex)
wskAtom.Connect
End Sub



Private Sub cMajSys_Click()
Time = lLocalTime.Caption
LocSyncroTime = Timer
sBaseTime = Time
lLocalTime.Caption = Time
SyncroTime = Timer
sSyncroTime = Time
lAtomicTime.Caption = Time

End Sub

Private Sub cSyncro_Click()
LocSyncroTime = Timer
SyncroTime = LocSyncroTime
sBaseTime = lAtomicTime.Caption
sAtomicTime = sBaseTime


End Sub

Private Sub Form_Load()
Dim i As Integer

ServerName(0) = "time-a.nist.gov"
ServerIp(0) = "129.6.15.28"
ServerLocation(0) = "NIST, Gaithersburg, Maryland"
ServerName(1) = "time-b.nist.gov"
ServerIp(1) = "129.6.15.29"
ServerLocation(1) = "NIST, Gaithersburg, Maryland"
ServerName(2) = "time-a.timefreq.bldrdoc.gov"
ServerIp(2) = "132.163.4.101"
ServerLocation(2) = "NIST, Boulder, Colorado"
ServerName(3) = "time-b.timefreq.bldrdoc.gov"
ServerIp(3) = "132.163.4.102"
ServerLocation(3) = "NIST, Boulder, Colorado"
ServerName(4) = "time-c.timefreq.bldrdoc.gov"
ServerIp(4) = "132.163.4.103"
ServerLocation(4) = "NIST, Boulder, Colorado"
ServerName(5) = "utcnist.colorado.edu"
ServerIp(5) = "128.138.140.44"
ServerLocation(5) = "University of Colorado, Boulder"
ServerName(6) = "time.nist.gov"
ServerIp(6) = "192.43.244.18"
ServerLocation(6) = "NCAR, Boulder, Colorado"
ServerName(7) = "time-nw.nist.gov"
ServerIp(7) = "131.107.1.10"
ServerLocation(7) = "Microsoft, Redmond, Washington"
ServerName(8) = "nist1.datum.com"
ServerIp(8) = "209.0.72.7"
ServerLocation(8) = "Datum, San Jose, California"
ServerName(9) = "nist1.dc.certifiedtime.com"
ServerIp(9) = "216.200.93.8"
ServerLocation(9) = "Abovnet, Virginia"
ServerName(10) = "nist1.nyc.certifiedtime.com"
ServerIp(10) = "208.184.49.9"
ServerLocation(10) = "Abovnet, New York City"
ServerName(11) = "nist1.sjc.certifiedtime.com"
ServerIp(11) = "208.185.146.41"
ServerLocation(11) = "Abovnet, San Jose, California"

For i = 0 To 11
    cmbServer.AddItem "[" & ServerName(i) & "] " & ServerLocation(i), i
Next i

For i = 12 To 1 Step -1
    cmbGMT.AddItem "[GMT -" & i & " h.]"
Next i
cmbGMT.AddItem "[GMT +0]"
For i = 1 To 12
    cmbGMT.AddItem "[GMT +" & i & " h.]"
Next i

cmbServer.ListIndex = 0
cmbGMT.ListIndex = 7
wskAtom.RemotePort = RemPort
LocSyncroTime = Timer
sBaseTime = Time
lLocalTime.Caption = Time
End Sub

Private Sub timeLocal_Timer()

Dim ElapSec As Single

ElapSec = Int(Timer - LocSyncroTime)
'Debug.Print ElapSec
If lLocalTime.Caption <> "00:00:00" Then
    lLocalTime.Caption = GetRealTime(sBaseTime, ElapSec, cmbGMT.ListIndex - 12)
End If


End Sub

Private Sub timeRemote_Timer()

Dim ElapSec As Single

ElapSec = Int(Timer - SyncroTime)
'Debug.Print ElapSec
If lAtomicTime.Caption <> "00:00:00" Then
    lAtomicTime.Caption = GetRealTime(sAtomicTime, ElapSec, cmbGMT.ListIndex - 12)
End If

End Sub

Private Sub wskAtom_DataArrival(ByVal bytesTotal As Long)

Dim tempData As String

wskAtom.GetData tempData
SyncroTime = Timer
lAtomicTime.Caption = ApplyGMT(Mid(tempData, 17, 8), cmbGMT.ListIndex - 12)
sAtomicTime = lAtomicTime.Caption
'Debug.Print sAtomicTime
If chkSys.Value = 1 Then
    Time = sAtomicTime
    DoEvents
    lLocalTime.Caption = Time
    sBaseTime = Time
    LocSyncroTime = Timer
    SyncroTime = LocSyncroTime
End If

wskAtom.Close

End Sub

Private Function GetRealTime(BaseATM As String, ElapSec As Single, GMT As Integer) As String

Dim sHours As String
Dim sMinutes As String
Dim sSeconds As String
Dim sOSeconds As String
Dim sOMinutes As String
Dim CarryHour As Integer
Dim iHours As Single

sHours = Mid(BaseATM, 1, 2)
sMinutes = Mid(BaseATM, 4, 2)
sSeconds = Mid(BaseATM, 7, 2)
sOSeconds = sSeconds
sOMinutes = sMinutes

sSeconds = CStr((Val(sSeconds) + Int(ElapSec)) Mod 60)
If Len(sSeconds) <> 2 Then sSeconds = "0" & sSeconds
sMinutes = CStr((Val(sMinutes) + ((Int(ElapSec) + Val(sOSeconds)) \ 60)) Mod 60)
If Len(sMinutes) <> 2 Then sMinutes = "0" & sMinutes
If sMinutes < sOMinutes Then CarryHour = 1
iHours = Val(sHours) + (Int(ElapSec) \ 3600) + CarryHour '+ GMT
'If iHours < 0 Then
'    sHours = CStr(24 + iHours)
'Else
'    sHours = CStr(iHours)
'End If

If Len(sHours) <> 2 Then sHours = "0" & sHours

GetRealTime = sHours & ":" & sMinutes & ":" & sSeconds
End Function


Private Function ApplyGMT(BaseATM, GMT As Integer)

Dim sHours As String
Dim iHours As Single
Dim sRest As String

sHours = Mid(BaseATM, 1, 2)
sRest = Mid(BaseATM, 3, 6)
iHours = Val(sHours) + GMT
If iHours < 0 Then
    sHours = CStr(24 + iHours)
Else
    iHours = iHours Mod 24
    sHours = CStr(iHours)
End If
If Len(sHours) <> 2 Then sHours = "0" & sHours
ApplyGMT = sHours & sRest

End Function
