VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "START COMMUNICATION"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      TabIndex        =   4
      Top             =   3600
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2760
      Width           =   1830
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "Injured"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   600
      TabIndex        =   10
      Top             =   7920
      Width           =   4335
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "Enemy Attack"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   600
      TabIndex        =   9
      Top             =   7080
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "Lost Direction"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   600
      TabIndex        =   8
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Caption         =   "Lack Of Food"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   600
      TabIndex        =   7
      Top             =   5400
      Width           =   4335
   End
   Begin VB.Label lbltemp 
      AutoSize        =   -1  'True
      Caption         =   "Temperature"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      TabIndex        =   6
      Top             =   4560
      Width           =   1710
   End
   Begin VB.Label lblheartrate 
      AutoSize        =   -1  'True
      Caption         =   "Heart Rate"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      TabIndex        =   5
      Top             =   3840
      Width           =   1470
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   885
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   9300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "G.P.S.Base Soldier Tracking and Health Indication System"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   14685
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Com Port"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10800
      TabIndex        =   1
      Top             =   2880
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recdata As String
Dim recchar As String
Dim lat1 As Variant
Dim lat2 As Variant
Dim lon1 As Variant
Dim lon2 As Variant
Dim latveh As Variant
Dim lonveh As Variant
Dim latdegree As Double
Dim latmin As Double
Dim latsec As Double
Dim longdegree As Double
Dim longmin As Double
Dim longsec As Double
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
    lpOperation As String, ByVal lpFile As String, ByVal _
    lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long


Private Sub Combo1_Change()
On Error GoTo err_handler:
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    MSComm1.CommPort = Mid(Combo1.Text, 5, Len(Combo1.Text) - 4)
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.Handshaking = comXOnXoff
    MSComm1.RThreshold = 1
    MSComm1.InputLen = 1
    Command1.Caption = "START COMMUNICATION"
Exit Sub
err_handler:
    MsgBox Err.Number & "-" & Err.Description
End Sub
Private Sub Combo1_Click()
    Call Combo1_Change
End Sub
Private Sub Command1_Click()
If Command1.Caption = "START COMMUNICATION" Then
    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
    Command1.Caption = "STOP COMMUNICATION"
Else
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Command1.Caption = "START COMMUNICATION"
End If
End Sub
Private Sub Form_Load()
On Error GoTo err_handler:


    Combo1.AddItem "COM 1"
    Combo1.AddItem "COM 2"
    Combo1.AddItem "COM 3"
    Combo1.AddItem "COM 4"
    Combo1.ListIndex = 0
Exit Sub
err_handler:
    MsgBox Err.Number & "-" & Err.Description
End Sub

Private Sub Label3_Click()

End Sub

Private Sub MSComm1_OnComm()
Select Case MSComm1.CommEvent
Case comEvReceive
    recchar = MSComm1.Input
    If recchar = "$" Then
        Label4.Caption = recdata
        Call ShowMap(recdata)
        recdata = ""
    Else
        recdata = recdata & recchar
    End If

End Select
End Sub
Private Sub ShowMap(strlocation As String)
If Mid(strlocation, 1, 1) = "*" Then
    lblheartrate = "Heart Rate = " & Mid(strlocation, 2, 3) & " BPM"
    lbltemp = "Temperature = " & Mid(strlocation, 6, 2) & " C"
    If Mid(strlocation, 10, 1) = "N" Then
        Label5.BackColor = vbGreen
    Else
        Label5.BackColor = vbRed
    End If

    If Mid(strlocation, 12, 1) = "N" Then
        Label6.BackColor = vbGreen
    Else
        Label6.BackColor = vbRed
    End If
    
    If Mid(strlocation, 14, 1) = "N" Then
        Label7.BackColor = vbGreen
    Else
        Label7.BackColor = vbRed
    End If
    
    If Mid(strlocation, 16, 1) = "N" Then
        Label8.BackColor = vbGreen
    Else
        Label8.BackColor = vbRed
    End If

strlocation = Mid(strlocation, 18)
Dim data1, data2 As String
Dim addr As String
''18.537018,N,07349.1885
', 73.870727
''




'strlocation = Mid(strlocation, 2, Len(strlocation) - 1)
latdegree = Mid(strlocation, 1, 2)
latmin = Mid(strlocation, 3, 2)
latsec = Round(Mid(strlocation, 5, 5) * 60, 0)

longdegree = Mid(strlocation, 11, 3)
longmin = Mid(strlocation, 14, 2)
longsec = Round(Mid(strlocation, 16, 5) * 60, 0)

data1 = latdegree + (latmin / 60) + (latsec / 3600)
data2 = longdegree + (longmin / 60) + (longsec / 3600)
'data1 = 18.498494
'data2 = 73.931007
addr = "http://maps.google.com/maps?q=" & data1 & "%20" & data2
Me.Caption = addr

website = "http://www.google.com/lochp?q=" & data1 & "%20" & data2
    ShellExecute 0, "open", _
        "C:\Program Files\Mozilla Firefox\Firefox.exe", website, vbNullString, 1

End If

End Sub

