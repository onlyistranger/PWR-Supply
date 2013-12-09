VERSION 5.00
Object = "{B13473EC-3CCE-11D2-B401-00201832C0F5}#1.0#0"; "MWStrip.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digita Power Supply Stripchart Logger (Abhishek Smart Devices)"
   ClientHeight    =   12600
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   18735
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12600
   ScaleWidth      =   18735
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Max values registered"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   22
      Top             =   8640
      Width           =   4695
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POWER:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VOLTAGE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CURRENT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1185
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Active Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3255
      Left            =   120
      TabIndex        =   13
      Top             =   5280
      Width           =   4695
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "DigifaceWide"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "DigifaceWide"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "DigifaceWide"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POWER(Watt):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VOLTAGE(V):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CURRENT(A):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Strip Chart Modes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   4695
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "WHITE TRACK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1800
         Width           =   4455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFF00&
         Caption         =   "ENABLE VARIABLE FILL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Width           =   4455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ENABLE GRIDS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Width           =   4455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF80&
         Caption         =   "ENABLE HANDLES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "WRAP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   11760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "COM PORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4695
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "OPEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   4455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":058A
         Left            =   1200
         List            =   "Form1.frx":05D0
         TabIndex        =   5
         Text            =   "2"
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COM PORT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1050
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin Mathworks_STRIPLib.Mathworks_Strip Strip3 
      Height          =   4215
      Left            =   5040
      TabIndex        =   2
      Top             =   8400
      Width           =   13695
      _Version        =   131082
      _Version        =   65536
      _ExtentX        =   24156
      _ExtentY        =   7435
      _StockProps     =   71
      BackColor       =   -2147483633
      BevelWidth      =   0
      BorderWidth     =   0
      AreaLeft        =   0.05
      AreaTop         =   0.04
      AreaRight       =   0.98
      m_BackColor     =   14737632
      DisplayMax      =   100
      Grid            =   0
      m_lastX         =   100
      MajorTics       =   10
      Max             =   100
      MaxBufferSize   =   10000
      MinorTics       =   10
      TicDelta        =   5
      TicMin          =   0
      TicMax          =   100
      TrackSeparation =   0
      m_cursorMode    =   1
      FontSize        =   2
      BeginProperty Font0 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionSize     =   2
      Caption0        =   "Time (1division=500ms)"
      CaptionX0       =   0.52
      CaptionY0       =   0.95
      CaptionVisible0 =   -1  'True
      CaptionFontID1  =   1
      Caption1        =   "Power(watt)"
      CaptionY1       =   0.53
      CaptionVisible1 =   -1  'True
      CaptionOrientation1=   1
      TrackBackColor0 =   0
      TrackDisplayMax0=   100
      TrackDisplayMin0=   0
      TrackMajorTics0 =   8
      TrackMax0       =   100
      TrackMin0       =   0
      TrackMinorTics0 =   5
      VariableColor0  =   16776960
   End
   Begin Mathworks_STRIPLib.Mathworks_Strip Strip2 
      Height          =   4215
      Left            =   5040
      TabIndex        =   1
      Top             =   4200
      Width           =   13695
      _Version        =   131082
      _Version        =   65536
      _ExtentX        =   24156
      _ExtentY        =   7435
      _StockProps     =   71
      BackColor       =   -2147483633
      BevelWidth      =   0
      BorderWidth     =   0
      AreaLeft        =   0.05
      AreaTop         =   0.04
      AreaRight       =   0.98
      AreaBottom      =   0.84
      m_BackColor     =   12632319
      DisplayMax      =   100
      Grid            =   0
      m_lastX         =   100
      MajorTics       =   10
      Max             =   100
      MaxBufferSize   =   10000
      MinorTics       =   10
      TicDelta        =   5
      TicMin          =   0
      TicMax          =   100
      m_cursorMode    =   1
      FontSize        =   2
      BeginProperty Font0 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionSize     =   2
      Caption0        =   "Time (1division=500ms)"
      CaptionX0       =   0.52
      CaptionY0       =   0.95
      CaptionVisible0 =   -1  'True
      CaptionFontID1  =   1
      Caption1        =   "Voltage(V)"
      CaptionY1       =   0.53
      CaptionVisible1 =   -1  'True
      CaptionOrientation1=   1
      TrackBackColor0 =   0
      TrackDisplayMax0=   30
      TrackDisplayMin0=   0
      TrackMax0       =   30
      TrackMin0       =   0
      TrackMinorTics0 =   5
      VariableColor0  =   65280
   End
   Begin Mathworks_STRIPLib.Mathworks_Strip Strip1 
      Height          =   4215
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      _Version        =   131082
      _Version        =   65536
      _ExtentX        =   24156
      _ExtentY        =   7435
      _StockProps     =   71
      BackColor       =   -2147483633
      BevelOuter      =   0
      BevelWidth      =   0
      BorderWidth     =   0
      OutlineWidth    =   0
      AreaLeft        =   0.05
      AreaTop         =   0.04
      AreaRight       =   0.98
      m_BackColor     =   16761024
      DisplayMax      =   100
      Grid            =   0
      m_lastX         =   100
      MajorTics       =   10
      Max             =   100
      MaxBufferSize   =   10000
      MinorTics       =   10
      TicDelta        =   5
      TicLabelOffset  =   12
      TicMin          =   0
      TicMax          =   100
      m_cursorMode    =   1
      FontSize        =   2
      BeginProperty Font0 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionSize     =   2
      Caption0        =   "Time (1division=500ms)"
      CaptionX0       =   0.52
      CaptionY0       =   0.95
      CaptionVisible0 =   -1  'True
      CaptionFontID1  =   1
      Caption1        =   "Current(A)"
      CaptionY1       =   0.51
      CaptionVisible1 =   -1  'True
      CaptionOrientation1=   1
      TrackBackColor0 =   0
      TrackDisplayMax0=   5
      TrackDisplayMin0=   0
      TrackMax0       =   5
      TrackMin0       =   0
      TrackMinorTics0 =   5
      VariableColor0  =   255
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "abhishekkumar1902@gmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   20
      Top             =   12240
      Width           =   3360
   End
   Begin VB.Image Image3 
      Height          =   1005
      Left            =   120
      Picture         =   "Form1.frx":0623
      Top             =   10200
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   120
      Picture         =   "Form1.frx":91A1
      Top             =   11280
      Width           =   3870
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   600
      Picture         =   "Form1.frx":156EB
      Top             =   120
      Width           =   3795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Command1_Click()
If Command1.Caption = "PAN" Then
Strip1.DisplayMode = PanMode
Strip2.DisplayMode = PanMode
Strip3.DisplayMode = PanMode
Command1.Caption = "WRAP"
Else
Strip1.DisplayMode = WrapMode
Strip2.DisplayMode = WrapMode
Strip3.DisplayMode = WrapMode
Command1.Caption = "PAN"
End If
End Sub
Private Sub Command2_Click()
If Combo1.Text = "" Then
    a = MsgBox("Hey,please select a valid com port...!!!", vbCritical, Error)
    sndPlaySound App.Path & "\valid_com_port.wav", &H1
Else
    If Command2.Caption = "OPEN" Then
        MSComm.CommPort = Combo1.Text
        On Error GoTo helpme
        MSComm.PortOpen = True
        Command2.Caption = "CLOSE"
        Command2.BackColor = &HFF00& 'green colour
        sndPlaySound App.Path & "\DeviceConnect.wav", &H1
    Else
        MSComm.PortOpen = False
        Command2.Caption = "OPEN"
        Command2.BackColor = &HFF&      'red colour
        sndPlaySound App.Path & "\DeviceDisconnect", &H1
    End If
End If
helpme:
If Err.Number = 8002 Or Err.Number = 8012 Then
sndPlaySound App.Path & "\valid_com_port.wav", &H1
a = MsgBox("Hey,please select a valid com port...!!!", vbCritical, Error)
Command2.Caption = "OPEN COM PORT"
Command2.BackColor = &HFF&      'red colour
End If
End Sub
Private Sub Command3_Click()
If Command3.Caption = "ENABLE HANDLES" Then
Strip1.Handles = BothHandles
Strip2.Handles = BothHandles
Strip3.Handles = BothHandles
Command3.Caption = "DISABLE HANDLES"
Else
Strip1.Handles = NoHandles
Strip2.Handles = NoHandles
Strip3.Handles = NoHandles
Command3.Caption = "ENABLE HANDLES"
End If
End Sub
Private Sub Command4_Click()
If Command4.Caption = "ENABLE GRIDS" Then
Strip1.Grid = BothGrid
Strip2.Grid = BothGrid
Strip3.Grid = BothGrid
Strip1.GridColor = &H808080
Strip2.GridColor = &H808080
Strip3.GridColor = &H808080
Command4.Caption = "DISABLE GRIDS"
Else
Strip1.Grid = NoGrid
Strip2.Grid = NoGrid
Strip3.Grid = NoGrid
Command4.Caption = "ENABLE GRIDS"
End If
End Sub
Private Sub Command5_Click()
If Command5.Caption = "ENABLE VARIABLE FILL" Then
Strip1.VariableFill = True
Strip2.VariableFill = True
Strip3.VariableFill = True
Command5.Caption = "DISABLE VARIABLE FILL"
Else
Strip1.VariableFill = False
Strip2.VariableFill = False
Strip3.VariableFill = False
Command5.Caption = "ENABLE VARIABLE FILL"
End If
End Sub

Private Sub Command6_Click()
If Command6.Caption = "WHITE TRACK" Then
Strip1.TrackBackColor = &HFFFFFF   'white colour
Strip2.TrackBackColor = &HFFFFFF   'white colour
Strip3.TrackBackColor = &HFFFFFF   'white colour
Command6.Caption = "DARK TRACK"
Else
Strip1.TrackBackColor = &H0& 'black colour
Strip2.TrackBackColor = &H0& 'black colour
Strip3.TrackBackColor = &H0& 'black colour
Command6.Caption = "WHITE TRACK"
End If
End Sub

Private Sub Form_Load()
sndPlaySound App.Path & "\intro.wav", &H1
Text1.Text = ""
MSComm.RThreshold = 1
MSComm.InputLen = 1
Label5.Caption = "000.00"
Label6.Caption = "000.00"
Label7.Caption = "000.00"
Strip1.DisplayMode = PanMode
Strip2.DisplayMode = PanMode
Strip3.DisplayMode = PanMode
Strip1.Handles = NoHandles
Strip2.Handles = NoHandles
Strip3.Handles = NoHandles
Strip1.VariableFill = False
Strip2.VariableFill = False
Strip3.VariableFill = False
Strip1.VariableColor = &HFF&    'red
Strip2.VariableColor = &HFF00&    'green
Strip3.VariableColor = &HFFFF00   'blue
End Sub
Private Sub Image2_Click()
If MSComm.PortOpen = True Then
    On Error Resume Next
    MSComm.PortOpen = False
End If
End
End Sub
Private Sub Image3_Click()
sndPlaySound App.Path & "\reset.wav", &H1
Strip1.ClearAll
Strip1.ClearTrack (0)
Strip1.ClearVariable (0)
Strip1.CursorX = 0
Strip2.ClearAll
Strip2.ClearTrack (0)
Strip2.ClearVariable (0)
Strip2.CursorX = 0
Strip3.ClearAll
Strip3.ClearTrack (0)
Strip3.ClearVariable (0)
Strip3.CursorX = 0
Label12.Caption = "000.00"
Label13.Caption = "000.00"
Label14.Caption = "000.00"
End Sub
Private Sub MSComm_OnComm()
Dim data As String
If MSComm.CommEvent = comEvReceive Then
            data = MSComm.Input
            '-------------------------------------------------------------
            If data = "A" Then
            Text1.Text = ""
            '-------------------------------------------------------------
            ElseIf data = "B" Then
            Label5.Caption = Text1.Text 'Format(Text1.Text, "###.00")
            Strip1.Y = Val(Label5.Caption)
            Label12 = Strip1.VariableMax(0)
            Text1.Text = ""
            '-------------------------------------------------------------
            ElseIf data = "C" Then
            Label6.Caption = Text1.Text 'Format(Text1.Text, "###.00")
            Strip2.Y = Val(Label6.Caption)
            Label13 = Strip2.VariableMax(0)
            Text1.Text = ""
            '-------------------------------------------------------------
            ElseIf data = "D" Then
            Label7.Caption = Text1.Text 'Format(Text1.Text, "###.00")
            Strip3.Y = Val(Label7.Caption)
            Label14 = Strip3.VariableMax(0)
            Text1.Text = ""
            '-------------------------------------------------------------
            Else
            Text1.Text = Text1.Text & data
            End If
End If
End Sub
