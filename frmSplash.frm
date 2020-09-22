VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "StarFleet Calculator"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10140
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H0080C0FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   2955
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBandBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   9900
      TabIndex        =   29
      Top             =   2520
      Width           =   9900
   End
   Begin VB.PictureBox picCorner2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00C0C0FF&
      Height          =   675
      Left            =   6960
      ScaleHeight     =   675
      ScaleWidth      =   3015
      TabIndex        =   28
      Top             =   480
      Width           =   3015
   End
   Begin VB.Timer tmrData 
      Interval        =   100
      Left            =   2400
      Top             =   1320
   End
   Begin VB.Timer tmrFlash 
      Interval        =   600
      Left            =   2040
      Top             =   1320
   End
   Begin VB.PictureBox picBand 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   3240
      ScaleHeight     =   150
      ScaleWidth      =   3615
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
   Begin VB.PictureBox picCorner 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00C0C0FF&
      Height          =   675
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   3015
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.PictureBox picSystem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF8080&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   885
      TabIndex        =   0
      Top             =   1200
      Width           =   880
      Begin VB.Label lblSystem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "SYSTEM"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   705
      End
   End
   Begin VB.Label lblDivision 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CEO Apollo Fleet's _Scorpion_"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   420
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Width           =   3720
   End
   Begin VB.Label lblSFCalc 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "StarFleet Calculator"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   420
      Left            =   3840
      TabIndex        =   5
      Top             =   720
      Width           =   2580
   End
   Begin VB.Label lblSD 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "by Marine Captain D'Arnak Jeffries"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   420
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
      Width           =   4545
   End
   Begin VB.Label lblConnect 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CONNECTING TO MAIN DATABASE"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   420
      Left            =   3120
      TabIndex        =   27
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label lblDataRow2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   9
      Left            =   9600
      TabIndex        =   26
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label lblDataRow2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   8
      Left            =   9600
      TabIndex        =   25
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label lblDataRow2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   7
      Left            =   9600
      TabIndex        =   24
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label lblDataRow2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   6
      Left            =   9600
      TabIndex        =   23
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label lblDataRow2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   5
      Left            =   9600
      TabIndex        =   22
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lblDataRow2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   4
      Left            =   8400
      TabIndex        =   21
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label lblDataRow2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   3
      Left            =   8400
      TabIndex        =   20
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label lblDataRow2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   2
      Left            =   8400
      TabIndex        =   19
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label lblDataRow2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   18
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label lblDataRow2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   8400
      TabIndex        =   17
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lblDataRow1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   9
      Left            =   9120
      TabIndex        =   16
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label lblDataRow1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   8
      Left            =   9120
      TabIndex        =   15
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label lblDataRow1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   7
      Left            =   9120
      TabIndex        =   14
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label lblDataRow1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   6
      Left            =   9120
      TabIndex        =   13
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label lblDataRow1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   5
      Left            =   9120
      TabIndex        =   12
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lblDataRow1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   11
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label lblDataRow1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   10
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label lblDataRow1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   2
      Left            =   7920
      TabIndex        =   9
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label lblDataRow1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   8
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label lblDataRow1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "05654"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   7
      Top             =   1200
      Width           =   465
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Load frmCalc
    Const PI = 3.14159265
'BottomLeft to TopRight LCARS Panel
    pX = picCorner.Width
    pY = picCorner.Height
    Colour = &HFFA0A0
    picCorner.AutoRedraw = True
    picCorner.FillColor = Colour
    picCorner.Circle (400, 400), 400, Colour
    picCorner.Line (0, 450)-(1000, pY), Colour, BF
    picCorner.Line (400, 0)-(pX, 500), Colour, BF
    picCorner.FillColor = picCorner.BackColor
    picCorner.Circle (1150, 400), 250, picCorner.BackColor
    picCorner.Line (1000 - 110, 500 - 125)-(pX, pY), picCorner.BackColor, BF
    picCorner.Line (1000 + 110, 150)-(pX, pY), picCorner.BackColor, BF
'TopLeft to BottomRight LCARS Panel
    pX = picCorner2.Width
    pY = picCorner2.Height
    Colour = RGB(255, 150, 100)
    picCorner2.AutoRedraw = True
    picCorner2.FillColor = Colour
    picCorner2.Circle (pX - 400, 400), 400, Colour
    picCorner2.Line (pX, 450)-(pX - 1000, pY), Colour, BF
    picCorner2.Line (pX - 400, 0)-(0, 500), Colour, BF
    picCorner2.FillColor = picCorner2.BackColor
    picCorner2.Circle (pX - 1150, 400), 250, picCorner2.BackColor
    picCorner2.Line (pX - 1000 + 110, 500 - 125)-(0, pY), picCorner2.BackColor, BF
    picCorner2.Line (pX - 1000 - 110, 150)-(0, pY), picCorner2.BackColor, BF
    Randomize Timer
    For X = 0 To 9
        lblDataRow1(X).ForeColor = RGB(200, 150, 100)
        lblDataRow1(X).Caption = Format(Rnd * 10000000, "########")
        lblDataRow2(X).ForeColor = RGB(200, 150, 100)
        lblDataRow2(X).Caption = Format(Rnd * 1000, "###")
    Next
    lblDataRow1(0).ForeColor = RGB(250, 200, 150)
    lblDataRow2(0).ForeColor = RGB(250, 200, 150)
'Bottom LCARS Band
    picBandBottom.Left = picCorner.Left
    picBandBottom.Width = picCorner2.Left + picCorner2.Width - 100
    pX = picBandBottom.Width
    pY = picBandBottom.Height
    Colour = &HA0FFA0
    picBandBottom.AutoRedraw = True
    picBandBottom.FillColor = Colour
    picBandBottom.Circle (pY / 2, pY / 2), pY / 2 - 20, Colour
    picBandBottom.Circle (pX - pY / 2, pY / 2), pY / 2 - 20, Colour
    picBandBottom.Line (pY / 2, 0)-(pX - pY / 2, pY), Colour, BF

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSystem.BackColor = &HFFC0C0
    tmrFlash.Enabled = True
End Sub

Private Sub imgLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSystem.BackColor = &HFFC0C0
    tmrFlash.Enabled = True
End Sub

Private Sub lblSystem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSystem.BackColor = &HFFE0E0
    lblSystem.ForeColor = RGB(0, 0, 0)
    tmrFlash.Enabled = False
End Sub

Private Sub picSystems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSystem.BackColor = &HFFE0E0
    lblSystem.ForeColor = RGB(0, 0, 0)
    tmrFlash.Enabled = False
End Sub


Private Sub picCorner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSystem.BackColor = &HFFC0C0
    tmrFlash.Enabled = True
End Sub

Private Sub tmrData_Timer()
    Static DataCount
    Static ConnectCount
    lblDataRow1(DataCount).ForeColor = RGB(200, 150, 100)
    lblDataRow2(DataCount).ForeColor = RGB(200, 150, 100)
    DataCount = DataCount + 1
    ConnectCount = ConnectCount + 1
    If DataCount = 10 Then DataCount = 0
    lblDataRow1(DataCount).ForeColor = RGB(250, 200, 150)
    lblDataRow1(DataCount).Caption = Format(Rnd * 10000000, "########")
    lblDataRow2(DataCount).ForeColor = RGB(250, 200, 150)
    lblDataRow2(DataCount).Caption = Format(Rnd * 1000, "###")
    If ConnectCount = 35 Then lblConnect.Caption = "DATABASE CONNECTED"
    If ConnectCount = 55 Then lblConnect.Caption = "AWAITING AUTHORISATION"
    If ConnectCount = 80 Then lblConnect.Caption = "AUTHORISATION 240-ALPHA-4-9-1-DELTA"
    If ConnectCount = 100 Then lblConnect.Caption = "PROCESSING"
    If ConnectCount = 130 Then lblConnect.Caption = "ACCESS GRANTED"
    If ConnectCount = 160 Then
        Me.Hide
        frmCalc.Show
    End If
End Sub

Private Sub tmrFlash_Timer()
    If lblSystem.ForeColor = RGB(0, 0, 0) Then
        lblSystem.ForeColor = RGB(255, 80, 80)
        lblSystem.Caption = "LOADING"
    Else
        lblSystem.ForeColor = RGB(0, 0, 0)
        lblSystem.Caption = "SYSTEM"
    End If
End Sub
