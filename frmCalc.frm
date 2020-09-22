VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Star Fleet Calculator"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Haettenschweiler"
      Size            =   9.75
      Charset         =   0
      Weight          =   800
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0080C0FF&
   Icon            =   "frmCalc.frx":0000
   LinkTopic       =   "frmCalc"
   MaxButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStarDate 
      Interval        =   1000
      Left            =   8880
      Top             =   6840
   End
   Begin MSComctlLib.StatusBar statBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   7275
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   661
      SimpleText      =   "Warpfactor to Lightspeed"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4604
            Picture         =   "frmCalc.frx":030A
            Text            =   "Warpfactor to Lightspeed"
            TextSave        =   "Warpfactor to Lightspeed"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5953
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5953
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmDistance 
      BackColor       =   &H00000000&
      Caption         =   "Enter distance to travel:"
      ForeColor       =   &H0080C0FF&
      Height          =   975
      Left            =   5880
      TabIndex        =   16
      Top             =   0
      Width           =   3495
      Begin VB.ComboBox cmbDistance 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   375
         ItemData        =   "frmCalc.frx":0626
         Left            =   1440
         List            =   "frmCalc.frx":063F
         TabIndex        =   19
         Text            =   "thousand lightyear"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtDistance 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Text            =   "100"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame frmWarpValue 
      BackColor       =   &H00000000&
      Caption         =   "Enter desired Warp speed:"
      ForeColor       =   &H0080C0FF&
      Height          =   975
      Left            =   2520
      TabIndex        =   3
      Top             =   0
      Width           =   3375
      Begin MSComCtl2.UpDown upWarpMinor 
         Height          =   615
         Index           =   3
         Left            =   3000
         TabIndex        =   15
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1085
         _Version        =   393216
         Value           =   9
         AutoBuddy       =   -1  'True
         BuddyControl    =   "lblWarpMinor(3)"
         BuddyDispid     =   196614
         BuddyIndex      =   3
         OrigLeft        =   3360
         OrigTop         =   240
         OrigRight       =   3600
         OrigBottom      =   855
         Max             =   9
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown upWarpMinor 
         Height          =   615
         Index           =   2
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1085
         _Version        =   393216
         Value           =   9
         AutoBuddy       =   -1  'True
         BuddyControl    =   "lblWarpMinor(2)"
         BuddyDispid     =   196614
         BuddyIndex      =   2
         OrigLeft        =   2640
         OrigTop         =   240
         OrigRight       =   2880
         OrigBottom      =   855
         Max             =   9
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown upWarpMinor 
         Height          =   615
         Index           =   1
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1085
         _Version        =   393216
         Value           =   9
         AutoBuddy       =   -1  'True
         BuddyControl    =   "lblWarpMinor(1)"
         BuddyDispid     =   196614
         BuddyIndex      =   1
         OrigLeft        =   1920
         OrigTop         =   240
         OrigRight       =   2160
         OrigBottom      =   855
         Max             =   9
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown upWarpMinor 
         Height          =   615
         Index           =   0
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1085
         _Version        =   393216
         Value           =   9
         AutoBuddy       =   -1  'True
         BuddyControl    =   "lblWarpMinor(0)"
         BuddyDispid     =   196614
         BuddyIndex      =   0
         OrigLeft        =   1200
         OrigTop         =   240
         OrigRight       =   1440
         OrigBottom      =   855
         Max             =   9
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown upWarpMajor 
         Height          =   615
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1085
         _Version        =   393216
         Value           =   9
         AutoBuddy       =   -1  'True
         BuddyControl    =   "lblWarpMajor"
         BuddyDispid     =   196616
         OrigLeft        =   480
         OrigTop         =   240
         OrigRight       =   720
         OrigBottom      =   855
         Max             =   9
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.Label lblWarpMinor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   3
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblWarpMinor 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblWarpMinor 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblWarpMinor 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblWarpDecimalPoint 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   75
      End
      Begin VB.Label lblWarpMajor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   270
         TabIndex        =   6
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.OptionButton optWarpScale 
      BackColor       =   &H00000000&
      Caption         =   "The New Generation scale"
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid flexWarp 
      Height          =   6615
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   26
      Cols            =   9
      BackColor       =   0
      ForeColor       =   8438015
      BackColorFixed  =   4210752
      ForeColorFixed  =   12640511
      BackColorSel    =   8421504
      ForeColorSel    =   12640511
      GridColor       =   12632319
      GridColorFixed  =   4210752
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Haettenschweiler"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frmWarpScale 
      BackColor       =   &H00000000&
      Caption         =   "Select Warp scale to use:"
      ForeColor       =   &H0080C0FF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton optWarpScale 
         BackColor       =   &H00000000&
         Caption         =   "Cochrane scale"
         ForeColor       =   &H00C0E0FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Unload frmSplash
End Sub

Private Sub Form_Load()
    statBar.Panels(2).Text = Date & " - " & Time
    For Col = 0 To 8
        flexWarp.ColWidth(Col) = (flexWarp.Width / 9) - 25
        flexWarp.ColAlignment(Col) = flexAlignCenterCenter
    Next
    For Row = 0 To 25
        flexWarp.RowHeight(Row) = (flexWarp.Height / 26) - 10
    Next
    flexWarp.Col = 0
    flexWarp.Row = 0
    flexWarp.Text = "Warp Factor"
    flexWarp.Col = 1
    flexWarp.Text = "* lightspeed"
    flexWarp.Col = 2
    flexWarp.Text = "Earth2Moon"
    flexWarp.Col = 3
    flexWarp.Text = "12 billion km"
    flexWarp.Col = 4
    flexWarp.Text = "Nearby star"
    flexWarp.Col = 5
    flexWarp.Text = "X sector (20ly)"
    flexWarp.Col = 6
    flexWarp.Text = "X Federation"
    flexWarp.Col = 7
    flexWarp.Text = "Andromeda 2M ly"
    flexWarp.Col = 8
    flexWarp.Text = "User def. dist."
    cmbDistance.ListIndex = 5
    optWarpScale(1).Value = True
End Sub

Private Sub optWarpScale_Click(Index As Integer)
    optWarpScale(1 - Index).Value = False
    Select Case Index
        Case 0 'Cochrane scale
            upWarpMajor.Max = 24
            upWarpMajor.Value = 24
            For X = 0 To 3
                lblWarpMinor(X).Enabled = False
                upWarpMinor(X).Value = 0
                upWarpMinor(X).Enabled = False
            Next
        Case 1 'TNG Scale
            upWarpMajor.Max = 9
            If upWarpMajor.Value > 9 Then upWarpMajor.Value = 9
            For X = 0 To 3
                lblWarpMinor(X).Enabled = True
                upWarpMinor(X).Enabled = True
                upWarpMinor(X).Value = 9
            Next
    End Select
    FillFlex
End Sub

Private Sub FillFlex()
    For X = 1 To 24
        flexWarp.Row = X
        flexWarp.Col = 0
        flexWarp.CellFontBold = True
        flexWarp.CellFontSize = 10
        If optWarpScale(0).Value = True Then
            LS = WarpToLight(CochScale(X), True)
            flexWarp.Text = CochScale(X)
        Else
            LS = WarpToLight(TNGScale(X), False)
            flexWarp.Text = TNGScale(X)
        End If
        flexWarp.Col = 1
        flexWarp.CellFontBold = True
        flexWarp.CellFontSize = 10
        flexWarp.Text = Format(LS, "#####0.0##")
        Travel
    Next
End Sub

Private Sub tmrStarDate_Timer()
    statBar.Panels(2).Text = Date & " - " & Time
End Sub

Private Sub txtDistance_Change()
    If Val(txtDistance.Text) Then ShowCustomDistance
End Sub

Private Sub txtDistance_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 8
        Case Asc(","), Asc(".")
            KeyAscii = Asc(".")
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub upWarpMajor_Change()
    ShowCustomWarp upWarpMajor.Value + (0.1 * upWarpMinor(0).Value) + (0.01 * upWarpMinor(1).Value) + (0.001 * upWarpMinor(2).Value) + (0.0001 * upWarpMinor(3).Value)
End Sub

Private Sub ShowCustomWarp(Factor As Double)
    flexWarp.Row = 25
    flexWarp.Col = 0
    flexWarp.CellFontSize = 12
    flexWarp.CellForeColor = RGB(100, 255, 150)
    flexWarp.Text = Factor
    flexWarp.Col = 1
    flexWarp.CellFontSize = 12
    flexWarp.CellForeColor = RGB(100, 255, 150)
    LS = WarpToLight(Factor, optWarpScale(0).Value)
    flexWarp.Text = Format(LS, "#####0.0##")
    'Time for Earth to Moon
    flexWarp.Col = 2
    flexWarp.CellFontSize = 12
    flexWarp.CellForeColor = RGB(100, 255, 150)
    flexWarp.CellAlignment = flexAlignLeftCenter
    flexWarp.Text = SecondsToTime(CDec(400000 / c / LS))
    'Time to cross Solar system
    flexWarp.Col = 3
    flexWarp.CellFontSize = 12
    flexWarp.CellForeColor = RGB(100, 255, 150)
    flexWarp.CellAlignment = flexAlignLeftCenter
    flexWarp.Text = SecondsToTime(CDec(12000000000# / c / LS))
    'Time to nearby star (5 ly)
    flexWarp.Col = 4
    flexWarp.CellFontSize = 12
    flexWarp.CellForeColor = RGB(100, 255, 150)
    flexWarp.CellAlignment = flexAlignLeftCenter
    flexWarp.Text = SecondsToTime(CDec(5 * ly / c / LS))
    'Time to cross sector (20 ly)
    flexWarp.Col = 5
    flexWarp.CellFontSize = 12
    flexWarp.CellForeColor = RGB(100, 255, 150)
    flexWarp.CellAlignment = flexAlignLeftCenter
    flexWarp.Text = SecondsToTime(CDec(20 * ly / c / LS))
    'Time to cross Federation space (8000 ly)
    flexWarp.Col = 6
    flexWarp.CellFontSize = 12
    flexWarp.CellForeColor = RGB(100, 255, 150)
    flexWarp.CellAlignment = flexAlignLeftCenter
    flexWarp.Text = SecondsToTime(CDec(8000 * ly / c / LS))
    'Time to Andromeda (2M ly)
    flexWarp.Col = 7
    flexWarp.CellFontSize = 12
    flexWarp.CellForeColor = RGB(100, 255, 150)
    flexWarp.CellAlignment = flexAlignLeftCenter
    flexWarp.Text = SecondsToTime(CDec(2000000# * ly / c / LS))
    'Time to user defined distance
    flexWarp.Col = 8
    flexWarp.CellFontSize = 12
    flexWarp.CellForeColor = RGB(100, 255, 150)
    flexWarp.CellAlignment = flexAlignLeftCenter
    TempDist = Val(txtDistance.Text)
    Select Case cmbDistance.ListIndex
        Case 0 ' TempDist * km
            flexWarp.Text = SecondsToTime(CDec(TempDist / c / LS))
        Case 1 ' TempDist * 1000 km
            flexWarp.Text = SecondsToTime(CDec(TempDist * 1000 / c / LS))
        Case 2 ' TempDist * M km
            flexWarp.Text = SecondsToTime(CDec(TempDist * 1000000# / c / LS))
        Case 3 ' TempDist * B km
            flexWarp.Text = SecondsToTime(CDec(TempDist * 1000000000# / c / LS))
        Case 4 ' TempDist * ly
            flexWarp.Text = SecondsToTime(CDec(TempDist * ly / c / LS))
        Case 5 ' TempDist * 1000 ly
            flexWarp.Text = SecondsToTime(CDec(TempDist * 1000 * ly / c / LS))
        Case 6 ' TempDist * M ly
            flexWarp.Text = SecondsToTime(CDec(TempDist * 1000000# * ly / c / LS))
    End Select
End Sub

Private Sub upWarpMinor_Change(Index As Integer)
    ShowCustomWarp upWarpMajor.Value + (0.1 * upWarpMinor(0).Value) + (0.01 * upWarpMinor(1).Value) + (0.001 * upWarpMinor(2).Value) + (0.0001 * upWarpMinor(3).Value)
End Sub

Private Sub Travel()
        'Time for Earth to Moon
        flexWarp.Col = 2
        flexWarp.CellFontBold = True
        flexWarp.CellFontSize = 10
        flexWarp.CellAlignment = flexAlignLeftCenter
        flexWarp.Text = SecondsToTime(CDec(400000 / c / LS))
        'Time to cross Solar system
        flexWarp.Col = 3
        flexWarp.CellFontBold = True
        flexWarp.CellFontSize = 10
        flexWarp.CellAlignment = flexAlignLeftCenter
        flexWarp.Text = SecondsToTime(CDec(12000000000# / c / LS))
        'Time to nearby star (5 ly)
        flexWarp.Col = 4
        flexWarp.CellFontBold = True
        flexWarp.CellFontSize = 10
        flexWarp.CellAlignment = flexAlignLeftCenter
        flexWarp.Text = SecondsToTime(CDec(5 * ly / c / LS))
        'Time to cross sector (20 ly)
        flexWarp.Col = 5
        flexWarp.CellFontBold = True
        flexWarp.CellFontSize = 10
        flexWarp.CellAlignment = flexAlignLeftCenter
        flexWarp.Text = SecondsToTime(CDec(20 * ly / c / LS))
        'Time to cross Federation space (8000 ly)
        flexWarp.Col = 6
        flexWarp.CellFontBold = True
        flexWarp.CellFontSize = 10
        flexWarp.CellAlignment = flexAlignLeftCenter
        flexWarp.Text = SecondsToTime(CDec(8000 * ly / c / LS))
        'Time to Andromeda (2M ly)
        flexWarp.Col = 7
        flexWarp.CellFontBold = True
        flexWarp.CellFontSize = 10
        flexWarp.CellAlignment = flexAlignLeftCenter
        flexWarp.Text = SecondsToTime(CDec(2000000# * ly / c / LS))
        'Time to user defined distance
        flexWarp.Col = 8
        flexWarp.CellFontSize = 12
        flexWarp.CellForeColor = RGB(100, 255, 150)
        flexWarp.CellAlignment = flexAlignLeftCenter
        TempDist = Val(txtDistance.Text)
        Select Case cmbDistance.ListIndex
            Case 0 ' TempDist * km
                flexWarp.Text = SecondsToTime(CDec(TempDist / c / LS))
            Case 1 ' TempDist * 1000 km
                flexWarp.Text = SecondsToTime(CDec(TempDist * 1000 / c / LS))
            Case 2 ' TempDist * M km
                flexWarp.Text = SecondsToTime(CDec(TempDist * 1000000# / c / LS))
            Case 3 ' TempDist * B km
                flexWarp.Text = SecondsToTime(CDec(TempDist * 1000000000# / c / LS))
            Case 4 ' TempDist * ly
                flexWarp.Text = SecondsToTime(CDec(TempDist * ly / c / LS))
            Case 5 ' TempDist * 1000 ly
                flexWarp.Text = SecondsToTime(CDec(TempDist * 1000 * ly / c / LS))
            Case 6 ' TempDist * M ly
                flexWarp.Text = SecondsToTime(CDec(TempDist * 1000000# * ly / c / LS))
        End Select
End Sub

Private Sub ShowCustomDistance()
    For X = 1 To 25
        If X < 25 Then
            If optWarpScale(0).Value = True Then
                LS = WarpToLight(CochScale(X), True)
            Else
                LS = WarpToLight(TNGScale(X), False)
            End If
        Else
            If optWarpScale(0).Value = True Then
                LS = WarpToLight(upWarpMajor.Value + (0.1 * upWarpMinor(0).Value) + (0.01 * upWarpMinor(1).Value) + (0.001 * upWarpMinor(2).Value) + (0.0001 * upWarpMinor(3).Value), True)
            Else
                LS = WarpToLight(upWarpMajor.Value + (0.1 * upWarpMinor(0).Value) + (0.01 * upWarpMinor(1).Value) + (0.001 * upWarpMinor(2).Value) + (0.0001 * upWarpMinor(3).Value), False)
            End If
        End If
        flexWarp.Row = X
        'Time to user defined distance
        flexWarp.Col = 8
        flexWarp.CellFontSize = 12
        flexWarp.CellForeColor = RGB(100, 255, 150)
        flexWarp.CellAlignment = flexAlignLeftCenter
        TempDist = Val(txtDistance.Text)
        Select Case cmbDistance.ListIndex
            Case 0 ' TempDist * km
                flexWarp.Text = SecondsToTime(CDec(TempDist / c / LS))
            Case 1 ' TempDist * 1000 km
                flexWarp.Text = SecondsToTime(CDec(TempDist * 1000 / c / LS))
            Case 2 ' TempDist * M km
                flexWarp.Text = SecondsToTime(CDec(TempDist * 1000000# / c / LS))
            Case 3 ' TempDist * B km
                flexWarp.Text = SecondsToTime(CDec(TempDist * 1000000000# / c / LS))
            Case 4 ' TempDist * ly
                flexWarp.Text = SecondsToTime(CDec(TempDist * ly / c / LS))
            Case 5 ' TempDist * 1000 ly
                flexWarp.Text = SecondsToTime(CDec(TempDist * 1000 * ly / c / LS))
            Case 6 ' TempDist * M ly
                flexWarp.Text = SecondsToTime(CDec(TempDist * 1000000# * ly / c / LS))
        End Select
    Next
End Sub

Private Sub cmbDistance_Click()
    If Val(txtDistance.Text) Then ShowCustomDistance
End Sub
