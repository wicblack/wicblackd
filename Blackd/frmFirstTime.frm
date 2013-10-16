VERSION 5.00
Begin VB.Form frmFirstTime 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackd Proxy - First run - Config"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmFirstTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "..."
      Height          =   285
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtTibiaMapsPath 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   5175
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00C0FFFF&
      Caption         =   ">> CONTINUE >>"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Frame frmOT 
      BackColor       =   &H00000000&
      Caption         =   "ENTER YOUR OT SERVER INFO:"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   7455
      Begin VB.TextBox txtOTport 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Text            =   "7171"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtOTip 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "<- leave this number unless they tell you other port!"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "OT server PORT:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblOTIP 
         BackColor       =   &H00000000&
         Caption         =   "OT server IP or domain:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00C0FFFF&
      Caption         =   "..."
      Height          =   285
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtTibiaClientPath 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
   Begin VB.ComboBox cmbVersion 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.CheckBox chkShowAgain 
      BackColor       =   &H00000000&
      Caption         =   "Show this window again next time"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label lblPath2 
      BackColor       =   &H00000000&
      Caption         =   "Your Tibia Maps Path:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblPath1 
      BackColor       =   &H00000000&
      Caption         =   "Your Tibia Client Path:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00000000&
      Caption         =   "Your Tibia Version:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmFirstTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit
Private Const cteOfficial As String = " - OFFICIAL TIBIA"
'Private Const cteOfficial As String = " - PREVIEW SERVERS"
Private useSamePath As Boolean


Private Sub cmbVersion_Click()
    Dim strSelected As String
    Dim samePath As Boolean
    If cmbVersion.ListIndex > -1 Then
        If InStr(1, cmbVersion.List(cmbVersion.ListIndex), cteOfficial, vbTextCompare) > 0 Then
            Me.txtTibiaClientPath.Text = autoGetTibiaFolder()
            frmOT.Visible = False
        Else
            frmOT.Visible = True
        End If
        strSelected = cmbVersion.List(cmbVersion.ListIndex)
        useSamePath = False
        Select Case strSelected
        Case "Tibia 7.4 (it uses 7.72 config)"
          useSamePath = True
        Case "Tibia 7.6"
          useSamePath = True
        Case "Tibia 7.7"
          useSamePath = True
        Case "Tibia 7.72"
          useSamePath = True
        Case "Tibia 7.8"
          useSamePath = True
        Case "Tibia 7.81"
          useSamePath = True
        Case "Tibia 7.9"
          useSamePath = True
        Case "Tibia 7.92"
          useSamePath = True
        End Select
        If useSamePath = True Then
            Me.txtTibiaClientPath.Text = ""
            Me.txtTibiaMapsPath.Text = ""
        Else
            TibiaVersionLong = highestTibiaVersionLong
            Me.txtTibiaMapsPath.Text = TryAutoPath()
       End If
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim res As String
    res = BrowseForFolder(Me.hwnd, "Select your Tibia Client folder")
    If res <> "" Then
        Me.txtTibiaClientPath.Text = res
    End If
End Sub

Private Sub cmdBrowse2_Click()
    Dim res As String
    res = BrowseForFolder(Me.hwnd, "Select your Tibia Maps folder")
    If res <> "" Then
        Me.txtTibiaMapsPath.Text = res
    End If
End Sub

Private Sub cmdContinue_Click()
    Unload Me
End Sub




Private Sub Form_Load()
useSamePath = False
  Me.Caption = "Blackd Proxy " & ProxyVersion & " - First run - Config"
  With Me.cmbVersion
   .Clear
   .AddItem "Tibia 10.01" & cteOfficial
   .AddItem "Tibia 10.01"
   .AddItem "Tibia 10.0"
   .AddItem "Tibia 9.92"
   .AddItem "Tibia 9.91"
   .AddItem "Tibia 9.9"
   .AddItem "Tibia 9.86"
   .AddItem "Tibia 9.85"
   .AddItem "Tibia 9.84"
   .AddItem "Tibia 9.83"
   .AddItem "Tibia 9.82"
   .AddItem "Tibia 9.81"
   .AddItem "Tibia 9.8"
   .AddItem "Tibia 9.71"
   .AddItem "Tibia 9.7"
   .AddItem "Tibia 9.63"
   .AddItem "Tibia 9.62"
   .AddItem "Tibia 9.61"
   .AddItem "Tibia 9.6"
   .AddItem "Tibia 9.54"
   .AddItem "Tibia 9.53"
   .AddItem "Tibia 9.52"
   .AddItem "Tibia 9.51"
   .AddItem "Tibia 9.5"
   .AddItem "Tibia 9.46"
   .AddItem "Tibia 9.45"
   .AddItem "Tibia 9.44"
   .AddItem "Tibia 9.43"
   .AddItem "Tibia 9.42"
   .AddItem "Tibia 9.41"
   .AddItem "Tibia 9.4"
   .AddItem "Tibia 9.31"
   .AddItem "Tibia 9.2"
   .AddItem "Tibia 9.1"
   .AddItem "Tibia 9.00"
   .AddItem "Tibia 8.74"
   .AddItem "Tibia 8.73"
   .AddItem "Tibia 8.72"
   .AddItem "Tibia 8.71"
   .AddItem "Tibia 8.7"
   .AddItem "Tibia 8.62"
   .AddItem "Tibia 8.61"
   .AddItem "Tibia 8.6"
   .AddItem "Tibia 8.57"
   .AddItem "Tibia 8.56"
   .AddItem "Tibia 8.55"
   .AddItem "Tibia 8.54"
   .AddItem "Tibia 8.53"
   .AddItem "Tibia 8.52"
   .AddItem "Tibia 8.5"
   .AddItem "Tibia 8.42"
   .AddItem "Tibia 8.41"
   .AddItem "Tibia 8.4"
   .AddItem "Tibia 8.31"
   .AddItem "Tibia 8.3"
   .AddItem "Tibia 8.22"
   .AddItem "Tibia 8.21"
   .AddItem "Tibia 8.2"
   .AddItem "Tibia 8.11"
   .AddItem "Tibia 8.1"
   .AddItem "Tibia 8.00"
   .AddItem "Tibia 7.92"
   .AddItem "Tibia 7.9"
   .AddItem "Tibia 7.81"
   .AddItem "Tibia 7.8"
   .AddItem "Tibia 7.72"
   .AddItem "Tibia 7.7"
   .AddItem "Tibia 7.6"
   .AddItem "Tibia 7.4 (it uses 7.72 config)"
   .Text = "Tibia " & TibiaVersionForceString & cteOfficial
   End With
   Me.txtTibiaClientPath.Text = autoGetTibiaFolder(defaultSelectedTibiaFolder)
   TibiaVersionLong = highestTibiaVersionLong
   Me.txtTibiaMapsPath.Text = TryAutoPath()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim configPath As String
    OVERWRITE_OT_MODE = True
    Select Case Me.cmbVersion.Text
    Case "Tibia 7.4 (it uses 7.72 config)"
      configPath = "config772"
    Case "Tibia 7.6"
      configPath = "config760"
    Case "Tibia 7.7"
      configPath = "config770"
    Case "Tibia 7.72"
      configPath = "config772"
    Case "Tibia 7.8"
      configPath = "config780"
    Case "Tibia 7.81"
      configPath = "config781"
    Case "Tibia 7.9"
      configPath = "config790"
    Case "Tibia 7.92"
      configPath = "config792"
    Case "Tibia 8.00"
      configPath = "config800"
    Case "Tibia 8.1"
      configPath = "config810"
    Case "Tibia 8.11"
      configPath = "config811"
    Case "Tibia 8.2"
      configPath = "config820"
    Case "Tibia 8.21"
      configPath = "config821"
    Case "Tibia 8.22"
      configPath = "config822"
    Case "Tibia 8.3"
      configPath = "config830"
    Case "Tibia 8.31"
      configPath = "config831"
    Case "Tibia 8.4"
      configPath = "config840"
    Case "Tibia 8.41"
      configPath = "config841"
    Case "Tibia 8.42"
      configPath = "config842"
    Case "Tibia 8.5"
      configPath = "config850"
    Case "Tibia 8.52"
      configPath = "config852"
    Case "Tibia 8.53"
      configPath = "config853"
    Case "Tibia 8.54"
      configPath = "config854"
    Case "Tibia 8.55"
      configPath = "config855"
    Case "Tibia 8.56"
      configPath = "config856"
    Case "Tibia 8.57"
      configPath = "config857"
    Case "Tibia 8.6"
      configPath = "config860"
    Case "Tibia 8.61"
      configPath = "config861"
    Case "Tibia 8.62"
      configPath = "config862"
    Case "Tibia 8.7"
      configPath = "config870"
    Case "Tibia 8.71"
      configPath = "config871"
    Case "Tibia 8.72"
      configPath = "config872"
    Case "Tibia 8.73"
      configPath = "config873"
    Case "Tibia 8.74"
      configPath = "config874"
    Case "Tibia 9.00"
      configPath = "config900"
    Case "Tibia 9.1"
      configPath = "config910"
    Case "Tibia 9.2"
      configPath = "config920"
    Case "Tibia 9.31"
      configPath = "config931"
    Case "Tibia 9.4"
      configPath = "config940"
    Case "Tibia 9.41"
      configPath = "config941"
    Case "Tibia 9.42"
      configPath = "config942"
    Case "Tibia 9.43"
      configPath = "config943"
    Case "Tibia 9.44"
      configPath = "config944"
    Case "Tibia 9.45"
      configPath = "config945"
    Case "Tibia 9.46"
      configPath = "config946"
    Case "Tibia 9.5"
      configPath = "config950"
    Case "Tibia 9.51"
      configPath = "config951"
    Case "Tibia 9.52"
      configPath = "config952"
    Case "Tibia 9.53"
      configPath = "config953"
    Case "Tibia 9.54"
      configPath = "config954"
    Case "Tibia 9.6"
      configPath = "config960"
    Case "Tibia 9.61"
      configPath = "config961"
    Case "Tibia 9.62"
      configPath = "config962"
    Case "Tibia 9.63"
      configPath = "config963"
    Case "Tibia 9.7"
      configPath = "config970"
    Case "Tibia 9.71"
      configPath = "config971"
    Case "Tibia 9.8"
      configPath = "config980"
    Case "Tibia 9.81"
      configPath = "config981"
    Case "Tibia 9.82"
      configPath = "config982"
    Case "Tibia 9.83"
      configPath = "config983"
    Case "Tibia 9.84"
      configPath = "config984"
    Case "Tibia 9.85"
      configPath = "config985"
    Case "Tibia 9.86"
      configPath = "config986"
    Case "Tibia 9.9"
      configPath = "config990"
    Case "Tibia 9.91"
      configPath = "config991"
    Case "Tibia 9.92"
      configPath = "config992"
    Case "Tibia 10.0"
      configPath = "config1000"
    Case "Tibia 10.01"
      configPath = "config1001"
    Case "Tibia 10.01" & cteOfficial
      configPath = "config1001"
      OVERWRITE_OT_MODE = False
    Case Else
      configPath = "config" & highestTibiaVersionLong
      OVERWRITE_OT_MODE = False
    End Select
    OVERWRITE_CONFIGPATH = configPath
    OVERWRITE_CLIENT_PATH = Me.txtTibiaClientPath.Text
    OVERWRITE_MAPS_PATH = Me.txtTibiaMapsPath.Text

    If OVERWRITE_OT_MODE = True Then
        OVERWRITE_OT_IP = Me.txtOTip.Text
        OVERWRITE_OT_PORT = safeLong(Me.txtOTport.Text)
    Else
        OVERWRITE_OT_IP = ""
        OVERWRITE_OT_PORT = 7171
    End If
    If (Me.chkShowAgain.Value = 1) Then
        OVERWRITE_SHOWAGAIN = True
    Else
        OVERWRITE_SHOWAGAIN = False
    End If
End Sub

Private Sub txtTibiaClientPath_Change()
    If useSamePath = True Then
        Me.txtTibiaMapsPath.Text = Me.txtTibiaClientPath.Text
    End If
End Sub
