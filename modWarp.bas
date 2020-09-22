Attribute VB_Name = "modWarp"
'Speed of light in kilometres/second
Global Const c As Long = 297600

'Lightyear in kilometres
Global Const ly As Variant = c * 86400# * 365.25

'Subspace field density
Global Const a As Double = 0.0026432

'Electromagnetic flux
Global Const n As Double = 2.879267

'Cochrane refraction index
Global Const f1 As Double = 0.0627412

'Cochrande reflection index
Global Const f2 As Double = 0.325746

'Temporary value of Warp expressed in number of lightspeed
Global LS As Variant

'Temporary value of User distance
Global TempDist As Variant

'The Cochrane Warp factor scale
Global CochScale(1 To 24) As Double

'The Next Generation Warp factor scale
Global TNGScale(1 To 24) As Double

Sub Main()

For X = 1 To 24
    CochScale(X) = X
Next

TNGScale(1) = 1
TNGScale(2) = 2
TNGScale(3) = 3
TNGScale(4) = 4
TNGScale(5) = 5
TNGScale(6) = 6
TNGScale(7) = 7
TNGScale(8) = 8
TNGScale(9) = 9
TNGScale(10) = 9.1
TNGScale(11) = 9.2
TNGScale(12) = 9.3
TNGScale(13) = 9.4
TNGScale(14) = 9.5
TNGScale(15) = 9.6
TNGScale(16) = 9.7
TNGScale(17) = 9.8
TNGScale(18) = 9.9
TNGScale(19) = 9.95
TNGScale(20) = 9.975
TNGScale(21) = 9.99
TNGScale(22) = 9.995
TNGScale(23) = 9.999
TNGScale(24) = 9.9999

Load frmSplash
frmSplash.Show

End Sub

Function WarpToLight(WarpFactor As Double, WarpScale As Boolean) As Double
    
    'WarpScale=False = Cochrane Scale
    'WarpScale=True = TNG Scale

    Select Case WarpScale
        Case True
            WarpToLight = WarpFactor ^ 3
        Case False
            Select Case WarpFactor
                Case Is <= 9
                    WarpToLight = WarpFactor ^ (10 / 3)
                Case Is > 9
                    WarpToLight = WarpFactor ^ (((10 / 3) + a * (-Log(10 - WarpFactor)) ^ n) + f1 * ((WarpFactor - 9) ^ 5) + f2 * ((WarpFactor - 9) ^ 11))
            End Select
  
    End Select

End Function

Function SecondsToTime(Seconds As Variant) As String
    If Seconds < 60 Then
        SecondsToTime = Format(Seconds, "#0.0000 Sec.")
    ElseIf Seconds >= 60 And Seconds < 3600 Then
        SecondsToTime = Format(Seconds / 60, "#0.000 Min.")
    ElseIf Seconds >= 3600 And Seconds < 86400# Then
        SecondsToTime = Format(Seconds / 3600, "#0.000 Hr.")
    ElseIf Seconds >= 86400# And Seconds < 31557600# Then
        SecondsToTime = Format(Seconds / 86400#, "##0.000 Day")
    ElseIf Seconds >= 31557600# And Seconds < 3155760000# Then
        SecondsToTime = Format(Seconds / 31557600#, "##0.000 Yr.")
    ElseIf Seconds >= 3155760000# And Seconds < 31557600000# Then
        SecondsToTime = Format(Seconds / 3155760000#, "##0.000 Cnt.")
    ElseIf Seconds >= 31557600000# And Seconds < 31557600000000# Then
        SecondsToTime = Format(Seconds / 31557600000#, "##0.000 kYr.")
    ElseIf Seconds >= 31557600000000# Then
        SecondsToTime = Format(Seconds / 31557600000000#, "##0.000 MYr.")
    End If
End Function
