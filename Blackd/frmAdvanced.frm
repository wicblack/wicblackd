VERSION 5.00
Begin VB.Form frmAdvanced 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Settings"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8205
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAdvanced.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAlternativeBinding 
      BackColor       =   &H00000000&
      Caption         =   "Alternative mc binding . CLOSE BLACKD PROXY AND ALL TIBIA MCS AFTER CHANGING THIS !!!!"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   5570
      Width           =   8055
   End
   Begin VB.ComboBox cmbTibiaServers 
      Height          =   315
      Left            =   4560
      TabIndex        =   29
      Top             =   7920
      Width           =   3495
   End
   Begin VB.TextBox txtLoginCharacter 
      Height          =   285
      Left            =   4560
      TabIndex        =   27
      Top             =   7560
      Width           =   3495
   End
   Begin VB.CheckBox chkWantBypass 
      BackColor       =   &H00000000&
      Caption         =   "I want to bypass login server. I understand that I will be banished / deleted if I use it on official servers."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   7080
      Width           =   7935
   End
   Begin VB.CommandButton cmdUpdatePIDs 
      BackColor       =   &H00C0FFFF&
      Caption         =   "UPDATE THIS LIST"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdSet25 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Set 25 FPS (to restore it)"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdSet0 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Set 0 FPS"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox cmbClients 
      Height          =   315
      Left            =   840
      TabIndex        =   14
      Text            =   "0 - First press UPDATETHIS LIST"
      Top             =   3960
      Width           =   3495
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CLICK HERE TO load a config for a different Tibia Version. This will close ALL. Reload later."
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3000
      Width           =   7695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Lock priorities"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   240
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.ComboBox cmbTibiaPriority 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Text            =   "Default - NORMAL"
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdApplyPriorities 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Apply Priorities"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox cmbMyPriority 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "Very High - HIGH"
      Top             =   600
      Width           =   2895
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8040
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "You still have to do the login virtually in the same way, with the correct account number and the correct password."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   8280
      Width           =   8055
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Game server where this character belong, or IP :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   7920
      Width           =   4335
   End
   Begin VB.Label lblPacketChar 
      BackColor       =   &H00000000&
      Caption         =   "Total exact name of the character you want to connect :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   7560
      Width           =   4335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   $"frmAdvanced.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   6360
      Width           =   7935
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Bypass login server:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8040
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   $"frmAdvanced.frx":053E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   7695
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Warning 1: 0 FPS is not compatible with ""Show exp in Tibia window tittle"" . This option will be disabled when you set 0 FPS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   7695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "FPS limiter:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblChar 
      BackColor       =   &H00000000&
      Caption         =   "Clients:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3960
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8040
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8040
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "If you will use MC, then set Tibia client priority to Default - NORMAL !!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   7095
   End
   Begin VB.Label lblWarningCPU 
      BackColor       =   &H00000000&
      Caption         =   "Warning: ABOVE_NORMAL and BELOW_NORMAL only works under Windows XP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   7095
   End
   Begin VB.Label lblForTibia 
      BackColor       =   &H00000000&
      Caption         =   "Default - NORMAL"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label lblForMe 
      BackColor       =   &H00000000&
      Caption         =   "Very High - HIGH"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Recommended settings:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Tibia clients:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Blackd Proxy:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "CPU Priority:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit
Public Sub LoadMyPriorityValue()
  Select Case MyPriorityID
  Case 0
    cmbMyPriority.Text = "Lowest - IDLE"
  Case 1
    cmbMyPriority.Text = "Low - BELOW_NORMAL"
  Case 2
    cmbMyPriority.Text = "Default - NORMAL"
  Case 3
    cmbMyPriority.Text = "High - ABOVE_NORMAL"
  Case 4
    cmbMyPriority.Text = "Very High - HIGH"
  Case 5
    cmbMyPriority.Text = "Highest - REALTIME"
  Case Else
    MyPriorityID = 2
    cmbMyPriority.Text = "Default - NORMAL"
  End Select
End Sub
Public Sub LoadTibiaPriorityValue()
  Select Case TibiaPriorityID
  Case 0
    cmbTibiaPriority.Text = "Lowest - IDLE"
  Case 1
    cmbTibiaPriority.Text = "Low - BELOW_NORMAL"
  Case 2
    cmbTibiaPriority.Text = "Default - NORMAL"
  Case 3
    cmbTibiaPriority.Text = "High - ABOVE_NORMAL"
  Case 4
    cmbTibiaPriority.Text = "Very High - HIGH"
  Case 5
    cmbTibiaPriority.Text = "Highest - REALTIME"
  Case Else
    TibiaPriorityID = 2
    cmbTibiaPriority.Text = "Default - NORMAL"
  End Select
End Sub



Private Sub Check1_Click()
  If Check1.Value = 1 Then
    cmbMyPriority.enabled = False
    cmbTibiaPriority.enabled = False
    cmdApplyPriorities.enabled = False
  Else
    cmbMyPriority.enabled = True
    cmbTibiaPriority.enabled = True
    cmdApplyPriorities.enabled = True
  End If
End Sub



Private Sub cmbMyPriority_Click()
  MyPriorityID = cmbMyPriority.ListIndex
  lblMessage.Caption = "<- PRESS TO APPLY CHANGES"
  lblMessage.ForeColor = &HFFFF&
End Sub

Private Sub cmbTibiaPriority_Click()
  TibiaPriorityID = cmbTibiaPriority.ListIndex
  lblMessage.Caption = "<- PRESS TO APPLY CHANGES"
  lblMessage.ForeColor = &HFFFF&
End Sub

Private Sub cmdApplyPriorities_Click()
  Dim pok As Boolean
  pok = UpdateMyPriority()
  If pok <> False Then
    pok = UpdateTibiaPriority()
  End If
  If pok = False Then
    LogOnFile "errors.txt", PriorityErrors
  End If
End Sub



'Private Sub cmdChange_Click()
'Dim strInfo As String
'Dim here As String
'Dim i As Long
'Select Case cmbConfig.Text
'Case "Tibia 7.4"
'  configPath = "config740"
'Case "Tibia 7.6"
'  configPath = "config760"
'Case "Tibia 7.7"
'  configPath = "config770"
'Case "Tibia 7.72"
'  configPath = "config772"
'Case "Tibia 7.8"
'  configPath = "config780"
'Case "Tibia 7.81"
'  configPath = "config781"
'Case "Tibia TEST"
'  configPath = "configTEST"
'Case "Tibia 7.9"
'  configPath = "config790"
'Case "Tibia 7.92"
'  configPath = "config792"
'Case "Tibia 8.00"
'  configPath = "config800"
'Case "Tibia 8.1"
'  configPath = "config810"
'Case "Tibia 8.11"
'  configPath = "config811"
'Case "Tibia 8.2"
'  configPath = "config820"
'Case "Tibia 8.21"
'  configPath = "config821"
'Case "Tibia 8.22"
'  configPath = "config822"
'Case "Tibia 8.3"
'  configPath = "config830"
'Case "Tibia 8.31"
'  configPath = "config831"
'Case "Tibia 8.4"
'  configPath = "config840"
'Case "Tibia 8.41"
'  configPath = "config841"
'Case "Tibia 8.42"
'  configPath = "config842"
'Case "Tibia 8.5"
'  configPath = "config850"
'Case "Tibia 8.52"
'  configPath = "config852"
'Case "Tibia 8.53"
'  configPath = "config853"
'Case "Tibia 8.54"
'  configPath = "config854"
'Case "Tibia 8.55"
'  configPath = "config855"
'Case "Tibia 8.56"
'  configPath = "config856"
'Case "Tibia 8.57"
'  configPath = "config857"
'Case "Tibia 8.6"
'  configPath = "config860"
'Case "Tibia 8.61"
'  configPath = "config861"
'Case "Tibia 8.62"
'  configPath = "config862"
'Case "Tibia 8.7"
'  configPath = "config870"
'Case "Tibia 8.71"
'  configPath = "config871"
'Case "Tibia 8.72"
'  configPath = "config872"
'Case "Tibia 8.73"
'  configPath = "config873"
'Case "Tibia 8.74"
'  configPath = "config874"
'Case "Tibia 9.00"
'  configPath = "config900"
'Case "Tibia 9.1"
'  configPath = "config910"
'Case "Tibia 9.2"
'  configPath = "config920"
'Case "Tibia 9.31"
'  configPath = "config931"
'Case "Tibia 9.4"
'  configPath = "config940"
'Case "Tibia 9.41"
'  configPath = "config941"
'Case "Tibia 9.42"
'  configPath = "config942"
'Case "Tibia 9.43"
'  configPath = "config943"
'Case "Tibia 9.44"
'  configPath = "config944"
'Case "Tibia 9.45"
'  configPath = "config945"
'Case "Tibia 9.46"
'  configPath = "config946"
'Case "Tibia 9.5"
'  configPath = "config950"
'Case "Tibia 9.51"
'  configPath = "config951"
'Case Else
'  configPath = ""
'End Select
'  here = App.path & "\config.ini"
'
'  strInfo = configPath
'  i = WritePrivateProfileString("Proxy", "configPath", strInfo, here)
'End
'End Sub

Private Function getSelectedPID() As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
Dim str As String
Dim chrstr As String
Dim i As Long
Dim lenstr As Long
Dim okcont As Boolean
Dim numstr As String
Dim lonnum As Long
okcont = True
numstr = ""
str = cmbClients.Text
lenstr = Len(str)
i = 0
Do
  i = i + 1
  If i > lenstr Then
    okcont = False
  Else
    chrstr = Mid(str, i, 1)
    If chrstr = " " Then
      okcont = False
    Else
      numstr = numstr & chrstr
    End If
  End If
Loop While (okcont = True)
If numstr = "" Then
  lonnum = 0
Else
  lonnum = CLng(numstr)
End If
getSelectedPID = lonnum
Exit Function
goterr:
getSelectedPID = 0
End Function
Private Sub ChangeInternalFPS(clientpid As Long, internalFPS As Long)
  Dim b1 As Byte
  Dim b2 As Byte
  Dim adrRealInternalFPS As Long
  If TibiaVersionLong < 770 Then
    adrRealInternalFPS = adrInternalFPS
  Else
    adrRealInternalFPS = &H5D + Memory_ReadLong(adrPointerToInternalFPSminusH5D, clientpid)
  End If
  b1 = LowByteOfLong(internalFPS)
  b2 = HighByteOfLong(internalFPS)
  Memory_WriteByte (adrRealInternalFPS), b1, clientpid, True
  Memory_WriteByte (adrRealInternalFPS + 1), b2, clientpid, True
End Sub

Private Sub cmdChange_Click()
  SaveConfigWizard True
  End
End Sub

Private Sub cmdSet0_Click()
Dim clientpid As Long
If frmHardcoreCheats.chkCaptionExp.Value = 1 Then
frmHardcoreCheats.chkCaptionExp.Value = 0
End If
clientpid = getSelectedPID()
If clientpid <> 0 Then
  ChangeInternalFPS clientpid, 40800
  lbl2.Caption = "Client #" & CStr(clientpid) & " now running at 0.50 FPS"
  lbl2.ForeColor = &HFF00&
Else
  lbl2.Caption = "ERROR: select a client first"
  lbl2.ForeColor = &HFFFF&
End If
End Sub


Private Sub cmdSet25_Click()
Dim clientpid As Long
clientpid = getSelectedPID()
If clientpid <> 0 Then
  ChangeInternalFPS clientpid, 17408
  lbl2.Caption = "Client #" & CStr(clientpid) & " now running at 25.00 FPS"
  lbl2.ForeColor = &HFF00&
Else
  lbl2.Caption = "ERROR: select a client first"
  lbl2.ForeColor = &HFFFF&
End If
End Sub

Private Sub cmdUpdatePIDs_Click()
  Dim i As Integer
  Dim compareID As String
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim IsConnected As Long
  Dim IsConnectedByte As Byte
  Dim relatedtothis As String
  Dim Message As String
  Dim addedc As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  addedc = 0
  GetProcessAllProcessIDs
  'hWndDesktop = GetDesktopWindow()
  tibiaclient = 0
  cmbClients.Clear
  Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      addedc = addedc + 1
      relatedtothis = ""
      For i = 1 To MAXCLIENTS
        If ((tibiaclient = ProcessID(i)) And (GameConnected(i) = True)) Then
          relatedtothis = CharacterName(i)
        End If
      Next i
      If relatedtothis = "" Then
        cmbClients.AddItem CStr(tibiaclient) & " - Not connected"
      Else
        cmbClients.AddItem CStr(tibiaclient) & " - " & relatedtothis
      End If
    End If
  Loop
  If addedc = 0 Then
    cmbClients.AddItem "0 - No Tibia clients found"
  End If
  cmbClients.Text = cmbClients.List(0)
  Exit Sub
goterr:
  cmbClients.Text = "0 - ERROR: " & Err.Description
End Sub

Private Sub Form_Load()
  Dim runningOnWinXP As Boolean
  'Dim verinfo As OSVERSIONINFO
  'Dim ret As Long
  'verinfo.dwOSVersionInfoSize = Len(verinfo)
  'ret = GetVersionEx(verinfo)
  'If ret = 0 Then
  '  runningOnWinXP = False
 ' Else
  '  If verinfo.dwMajorVersion = 5 Then
  '    runningOnWinXP = True
  '  Else
     runningOnWinXP = False
  '  End If
  'End If

  'If runningOnWinXP = True Then
  '  lblForTibia.Caption = "High - ABOVE_NORMAL (unless you use MC)"
  '  lblWarningCPU.Caption = ""
  'Else
    lblForTibia.Caption = "Default - NORMAL"
  'End If
'  With cmbConfig
'   .Clear
'   .AddItem "Tibia TEST"
'   .AddItem "Tibia 9.51"
'   .AddItem "Tibia 9.5"
'   .AddItem "Tibia 9.46"
'   .AddItem "Tibia 9.45"
'   .AddItem "Tibia 9.44"
'   .AddItem "Tibia 9.43"
'   .AddItem "Tibia 9.42"
'   .AddItem "Tibia 9.41"
'   .AddItem "Tibia 9.4"
'   .AddItem "Tibia 9.31"
'   .AddItem "Tibia 9.2"
'   .AddItem "Tibia 9.1"
'   .AddItem "Tibia 9.00"
'   .AddItem "Tibia 8.74"
'   .AddItem "Tibia 8.73"
'   .AddItem "Tibia 8.72"
'   .AddItem "Tibia 8.71"
'   .AddItem "Tibia 8.7"
'   .AddItem "Tibia 8.62"
'   .AddItem "Tibia 8.61"
'   .AddItem "Tibia 8.6"
'   .AddItem "Tibia 8.57"
'   .AddItem "Tibia 8.56"
'   .AddItem "Tibia 8.55"
'   .AddItem "Tibia 8.54"
'   .AddItem "Tibia 8.53"
'   .AddItem "Tibia 8.52"
'   .AddItem "Tibia 8.5"
'   .AddItem "Tibia 8.42"
'   .AddItem "Tibia 8.41"
'   .AddItem "Tibia 8.4"
'   .AddItem "Tibia 8.31"
'   .AddItem "Tibia 8.3"
'   .AddItem "Tibia 8.22"
'   .AddItem "Tibia 8.21"
'   .AddItem "Tibia 8.2"
'   .AddItem "Tibia 8.11"
'   .AddItem "Tibia 8.1"
'   .AddItem "Tibia 8.00"
'   .AddItem "Tibia 7.92"
'   .AddItem "Tibia 7.9"
'   .AddItem "Tibia 7.81"
'   .AddItem "Tibia 7.8"
'   .AddItem "Tibia 7.72"
'   .AddItem "Tibia 7.7"
'   .AddItem "Tibia 7.6"
'   .AddItem "Tibia 7.4"
'   .Text = "Tibia " & TibiaVersionDefaultString
'  End With
'  Select Case configPath
'    Case "config740"
'      cmbConfig.Text = "Tibia 7.4"
'    Case "config760"
'      cmbConfig.Text = "Tibia 7.6"
'    Case "config770"
'      cmbConfig.Text = "Tibia 7.7"
'    Case "config772"
'      cmbConfig.Text = "Tibia 7.72"
'    Case "config780"
'      cmbConfig.Text = "Tibia 7.8"
'    Case "config781"
'      cmbConfig.Text = "Tibia 7.81"
'    Case "config790"
'      cmbConfig.Text = "Tibia 7.9"
'    Case "config792"
'      cmbConfig.Text = "Tibia 7.92"
'    Case "config800"
'      cmbConfig.Text = "Tibia 8.00"
'    Case "config810"
'      cmbConfig.Text = "Tibia 8.1"
'    Case "config811"
'      cmbConfig.Text = "Tibia 8.11"
'    Case "config820"
'      cmbConfig.Text = "Tibia 8.2"
'    Case "config821"
'      cmbConfig.Text = "Tibia 8.21"
'    Case "config822"
'      cmbConfig.Text = "Tibia 8.22"
'    Case "config830"
'      cmbConfig.Text = "Tibia 8.3"
'    Case "config831"
'      cmbConfig.Text = "Tibia 8.31"
'    Case "config840"
'      cmbConfig.Text = "Tibia 8.4"
'    Case "config841"
'      cmbConfig.Text = "Tibia 8.41"
'    Case "config842"
'      cmbConfig.Text = "Tibia 8.42"
'    Case "config850"
'      cmbConfig.Text = "Tibia 8.5"
'    Case "config852"
'      cmbConfig.Text = "Tibia 8.52"
'    Case "config853"
'      cmbConfig.Text = "Tibia 8.53"
'    Case "config854"
'      cmbConfig.Text = "Tibia 8.54"
'    Case "config855"
'      cmbConfig.Text = "Tibia 8.55"
'    Case "config856"
'      cmbConfig.Text = "Tibia 8.56"
'    Case "config857"
'      cmbConfig.Text = "Tibia 8.57"
'    Case "config860"
'      cmbConfig.Text = "Tibia 8.6"
'    Case "config861"
'      cmbConfig.Text = "Tibia 8.61"
'    Case "config862"
'      cmbConfig.Text = "Tibia 8.62"
'    Case "config870"
'      cmbConfig.Text = "Tibia 8.7"
'    Case "config871"
'      cmbConfig.Text = "Tibia 8.71"
'    Case "config872"
'      cmbConfig.Text = "Tibia 8.72"
'    Case "config873"
'      cmbConfig.Text = "Tibia 8.73"
'    Case "config874"
'      cmbConfig.Text = "Tibia 8.74"
'    Case "config900"
'      cmbConfig.Text = "Tibia 9.00"
'    Case "config910"
'      cmbConfig.Text = "Tibia 9.1"
'    Case "config920"
'      cmbConfig.Text = "Tibia 9.2"
'    Case "config931"
'      cmbConfig.Text = "Tibia 9.31"
'    Case "config940"
'      cmbConfig.Text = "Tibia 9.4"
'    Case "config941"
'      cmbConfig.Text = "Tibia 9.41"
'    Case "config942"
'      cmbConfig.Text = "Tibia 9.42"
'    Case "config943"
'      cmbConfig.Text = "Tibia 9.43"
'    Case "config944"
'      cmbConfig.Text = "Tibia 9.44"
'    Case "config945"
'      cmbConfig.Text = "Tibia 9.45"
'    Case "config946"
'      cmbConfig.Text = "Tibia 9.46"
'    Case "config950"
'      cmbConfig.Text = "Tibia 9.5"
'    Case "config951"
'      cmbConfig.Text = "Tibia 9.51"
'    Case "configTEST"
'      cmbConfig.Text = "Tibia TEST"
'    Case Else
'      cmbConfig.Text = "Tibia " & TibiaVersionDefaultString
'  End Select
  With cmbMyPriority
   .Clear
   .AddItem "Lowest - IDLE"
   .AddItem "Low - BELOW_NORMAL"
   .AddItem "Default - NORMAL"
   .AddItem "High - ABOVE_NORMAL"
   .AddItem "Very High - HIGH"
   .AddItem "Highest - REALTIME"
  End With
  LoadMyPriorityValue
  With cmbTibiaPriority
   .Clear
   .AddItem "Lowest - IDLE"
   .AddItem "Low - BELOW_NORMAL"
   .AddItem "Default - NORMAL"
   .AddItem "High - ABOVE_NORMAL"
   .AddItem "Very High - HIGH"
   .AddItem "Highest - REALTIME"
  End With
  LoadTibiaPriorityValue
  LoadServerIps
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Cancel = BlockUnload
End Sub




