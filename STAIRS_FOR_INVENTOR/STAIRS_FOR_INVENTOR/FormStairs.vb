
Option Explicit On
Imports Inventor



Public Class Stairs_Class

    Dim oStairsHeight As Double
    Dim oLadderHeight As Double
    Dim oStairsAngle As Double
    Dim oStairsWidth As Double
    Dim oNumberOfSteps As Integer
    Dim oNumberOfRungs As Integer
    Dim oStepheight As Double
    Dim oRungheight As Double
    Dim oSPAA As Double
    Dim oSPAAtop As Double
    Dim oStairsHypotenusa As Double
    Dim oNumberOfSpils As Integer
    Dim oNumberOfHoops As Integer
    Dim pi As Double
    Dim oAngle As Double
    Dim oRadians As Double
    Dim oInvApp As Inventor.Application
    Dim oInvDoc As Document
    Dim oHandrailState As String
    Dim oCageState As String
    Dim oProjectPath As String
    Dim oFilename As String
    Dim oLadderFilename As String

    Dim oRailFilename As String
    Dim oRailingState As String
    Dim oRailLng1 As Double
    Dim oRailLng2 As Double
    Dim oRailLng3 As Double
    Dim oRailLng4 As Double
    Dim oRailLng5 As Double
    Dim oRailDia As Double
    Dim oNumberOfRailSpils1 As Integer
    Dim oNumberOfRailSpils2 As Integer
    Dim oNumberOfRailSpils3 As Integer
    Dim oNumberOfRailSpils4 As Integer
    Dim oNumberOfRailSpils5 As Integer
    Dim oNumberOfRailSpilsR As Integer
    Dim oRailAngle As Double
    Dim oDiaNr As Integer

    Dim oCSHeight As Double
    Dim oCSWidth As Double
    Dim oCSAngle As Double
    Dim oCSIntRadius As Double
    Dim oCSState As String
    Dim oCSFilename As String


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'TopMost = True
        'Dim oExpdate As Long
        'Dim oThisdate = Date.Now.Ticks
        'oExpdate = 636920928000000000 'Monday, April 29, 2019

        ''MsgBox(oThisdate)
        'If oThisdate > oExpdate Then
        '    MsgBox("Fatal error." & vbCrLf & "The database rebuild failed to start.", MsgBoxStyle.Critical)

        '    Close()
        'End If


        ' Attempt to get a reference to a running instance of Inventor.
        Try
            oInvApp = GetObject(, "Inventor.Application")
        Catch ex As Exception
            MsgBox("Inventor must be running.")
            btnPlaceStairsInAssem.Enabled = False
            btnPlaceLadder.Enabled = False
            btnPlaceRailing.Enabled = False
            btnPlacePlatform.Enabled = False
            btnPlaceCurvedStairs.Enabled = False
            Exit Sub
        End Try
        'MsgBox(oInvApp.Caption + " Is Actif")

        oInvApp.ScreenUpdating = True

        oStairsHeight = 4000
        oStairsAngle = 42
        oStairsWidth = 800
        oHandrailState = "Both"
        oLadderHeight = 3000
        oCageState = "Cage"
        oRailLng1 = 1500
        oRailLng2 = 2500
        oRailLng3 = 4000
        oRailLng4 = 2500
        oRailLng5 = 1500
        oRailDia = 3000
        oRailingState = "1Side"
        oPlatformAngle = 60
        oPlatformLenght = 3000
        oPlatformHeight = 3000
        oPlatformWidth = 1000
        CbxFoot.Checked = True
        CbxLegs.Checked = True
        oPlatformState = "RECT_LEGS_FOOT"
        oCSHeight = 4000
        oCSWidth = 800
        oCSAngle = 42
        oCSIntRadius = 2500
        oCSState = "CW_Both"
        cboWidth.SelectedIndex = 4
        cboHeight.SelectedIndex = 0
        cboAngle.SelectedIndex = 5
        cboCurvedStairsWidth.SelectedIndex = 4
        cboCurvedStairsAngle.SelectedIndex = 5
        CboPlatformLegNumber.SelectedIndex = 0


        get_ProjectPath()
        get_Filename()
        get_LadderFilename()
        get_PlatformFilename()
        get_Railing1Filename()
        get_CSFilename()


    End Sub



    Private Sub chkLeftHandrail_CheckedChanged(sender As Object, e As EventArgs) Handles chkLeftHandrail.CheckedChanged
        If chkLeftHandrail.CheckState = 1 AndAlso chkRightHandrail.CheckState = 1 Then
            picRailingBoth.Visible = True
            picRailingLeft.Visible = False
            picRailingRight.Visible = False
            picRailingNone.Visible = False
            oHandrailState = "Both"
        End If
        If chkLeftHandrail.CheckState = 1 AndAlso chkRightHandrail.CheckState = 0 Then
            picRailingBoth.Visible = False
            picRailingLeft.Visible = True
            picRailingRight.Visible = False
            picRailingNone.Visible = False
            oHandrailState = "Left"
        End If
        If chkLeftHandrail.CheckState = 0 AndAlso chkRightHandrail.CheckState = 1 Then
            picRailingBoth.Visible = False
            picRailingLeft.Visible = False
            picRailingRight.Visible = True
            picRailingNone.Visible = False
            oHandrailState = "Right"
        End If
        If chkLeftHandrail.CheckState = 0 AndAlso chkRightHandrail.CheckState = 0 Then
            picRailingBoth.Visible = False
            picRailingLeft.Visible = False
            picRailingRight.Visible = False
            picRailingNone.Visible = True
            oHandrailState = "None"
        End If
        get_ProjectPath()
        get_Filename()
    End Sub

    Private Sub chkRightHandrail_CheckedChanged(sender As Object, e As EventArgs) Handles chkRightHandrail.CheckedChanged
        If chkLeftHandrail.CheckState = 1 AndAlso chkRightHandrail.CheckState = 1 Then
            picRailingBoth.Visible = True
            picRailingLeft.Visible = False
            picRailingRight.Visible = False
            picRailingNone.Visible = False
            oHandrailState = "Both"
        End If
        If chkLeftHandrail.CheckState = 1 AndAlso chkRightHandrail.CheckState = 0 Then
            picRailingBoth.Visible = False
            picRailingLeft.Visible = True
            picRailingRight.Visible = False
            picRailingNone.Visible = False
            oHandrailState = "Left"
        End If
        If chkLeftHandrail.CheckState = 0 AndAlso chkRightHandrail.CheckState = 1 Then
            picRailingBoth.Visible = False
            picRailingLeft.Visible = False
            picRailingRight.Visible = True
            picRailingNone.Visible = False
            oHandrailState = "Right"
        End If
        If chkLeftHandrail.CheckState = 0 AndAlso chkRightHandrail.CheckState = 0 Then
            picRailingBoth.Visible = False
            picRailingLeft.Visible = False
            picRailingRight.Visible = False
            picRailingNone.Visible = True
            oHandrailState = "None"
        End If
        get_ProjectPath()
        get_Filename()
    End Sub

    Private Sub get_ProjectPath()
        Try
            oInvApp = GetObject(, "Inventor.Application")
            oProjectPath = oInvApp.FileLocations.Workspace
            lblPath.Text = oProjectPath & "\Mechanical\Stairs\"
            lblLadderPath.Text = oProjectPath & "\Mechanical\Ladders\"
            lblRailPath.Text = oProjectPath & "\Mechanical\Railing\"
            lblPlatformPath.Text = oProjectPath & "\Mechanical\Platform\"
            lblCurvedStairsPath.Text = oProjectPath & "\Mechanical\Stairs\"
        Catch ex As Exception
            lblPath.Text = "No path Found"
            lblLadderPath.Text = "No path Found"
            lblRailPath.Text = "No path Found"
            lblPlatformPath.Text = "No path Found"
            lblCurvedStairsPath.Text = "No path Found"
            Exit Sub
        End Try

    End Sub

    Private Sub cboHeight_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles cboHeight.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If cboHeight.Text = Nothing Then
                cboHeight.Text = "600"
            End If
            If CDbl(cboHeight.Text) < 600 Then
                oStairsHeight = 600
                cboHeight.Text = 600
            End If
            If CDbl(cboHeight.Text) > 6000 Then
                oStairsHeight = 6000
                cboHeight.Text = 6000
            End If
            oStairsHeight = CDbl(cboHeight.Text)
        End If
        get_Filename()
    End Sub
    Private Sub cboHeight_MouseLeave(sender As Object, e As EventArgs) Handles cboHeight.MouseLeave, btnPlaceStairsInAssem.MouseEnter
        If cboHeight.Text = Nothing Then
            cboHeight.Text = 600
        End If
        oStairsHeight = CDbl(cboHeight.Text)
        If oStairsHeight < 600 Then
            oStairsHeight = 600
            cboHeight.Text = 600
        End If
        If oStairsHeight > 6000 Then
            oStairsHeight = 6000
            cboHeight.Text = 6000
        End If
        get_ProjectPath()
        get_Filename()

    End Sub
    Private Sub cboHeight_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboHeight.SelectedIndexChanged
        oStairsHeight = CDbl(cboHeight.Text)
        If oStairsHeight < 600 Then
            oStairsHeight = 600
            cboHeight.Text = 600
        End If
        If oStairsHeight > 6000 Then
            oStairsHeight = 6000
            cboHeight.Text = 6000
        End If
        get_ProjectPath()
        get_Filename()
    End Sub
    Private Sub cboHeight_LostFocus(sender As Object, e As EventArgs) Handles cboHeight.LostFocus
        If cboHeight.Text = Nothing Then
            cboHeight.Text = 600
        End If
        oStairsHeight = CDbl(cboHeight.Text)
        If oStairsHeight < 600 Then
            oStairsHeight = 600
            cboHeight.Text = 600
        End If
        If oStairsHeight > 6000 Then
            oStairsHeight = 6000
            cboHeight.Text = 6000
        End If
        get_ProjectPath()
        get_Filename()
    End Sub






    Private Sub cboLadderHeight_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles cboLadderHeight.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If cboLadderHeight.Text = Nothing Then
                cboLadderHeight.Text = "1200"
            End If
            If CDbl(cboLadderHeight.Text) < 600 Then
                oLadderHeight = 600
                cboLadderHeight.Text = 600
            End If
            If CDbl(cboLadderHeight.Text) > 6000 Then
                oLadderHeight = 6000
                cboLadderHeight.Text = 6000
            End If
            oLadderHeight = CDbl(cboLadderHeight.Text)
            get_LadderFilename()
        End If
    End Sub
    Private Sub cboLadderHeight_MouseLeave(sender As Object, e As EventArgs) Handles cboLadderHeight.MouseLeave
        If cboLadderHeight.Text = Nothing Then
            cboLadderHeight.Text = 1200
        End If
        oLadderHeight = CDbl(cboLadderHeight.Text)
        If oLadderHeight < 1200 Then
            oLadderHeight = 1200
            cboLadderHeight.Text = 1200
        End If
        If oLadderHeight < 2300 Then
            oCageState = "NoCage"
            chkCage.Enabled = False
            chkCage.Checked = False
        End If

        If oLadderHeight >= 2300 Then
            oCageState = "Cage"
            chkCage.Enabled = True
            chkCage.Checked = True

        End If
        If oLadderHeight > 6000 Then
            oLadderHeight = 6000
            cboLadderHeight.Text = 6000
        End If
        get_ProjectPath()
        get_LadderFilename()
    End Sub

    Private Sub cboLadderHeight_LostFocus(sender As Object, e As EventArgs) Handles cboLadderHeight.LostFocus
        If cboLadderHeight.Text = Nothing Then
            cboLadderHeight.Text = 1200
        End If
        oLadderHeight = CDbl(cboLadderHeight.Text)
        If oLadderHeight < 1200 Then
            oLadderHeight = 1200
            cboLadderHeight.Text = 1200
        End If
        If oLadderHeight < 2300 Then
            oCageState = "NoCage"
            chkCage.Enabled = False
            chkCage.Checked = False
        End If

        If oLadderHeight >= 2300 Then
            oCageState = "Cage"
            chkCage.Enabled = True
            chkCage.Checked = True

        End If
        If oLadderHeight > 6000 Then
            oLadderHeight = 6000
            cboLadderHeight.Text = 6000
        End If
        get_ProjectPath()
        get_LadderFilename()
    End Sub



    Private Sub cboLadderHeight_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLadderHeight.SelectedIndexChanged
        oLadderHeight = CDbl(cboLadderHeight.Text)
        If oLadderHeight < 1500 Then
            oLadderHeight = 1500
            cboLadderHeight.Text = 1500
        End If
        If oLadderHeight < 2400 Then
            oCageState = "NoCage"
            chkCage.Enabled = False
            chkCage.Checked = False
        End If

        If oLadderHeight >= 2400 Then
            oCageState = "Cage"
            chkCage.Enabled = True
            chkCage.Checked = True

        End If

        If oLadderHeight > 6000 Then
            oLadderHeight = 6000
            cboLadderHeight.Text = 6000
        End If
        get_ProjectPath()
        get_LadderFilename()
    End Sub


    Private Sub get_Filename()
        lblFile.Text = "STAIRS" & oStairsHeight & "x" & oStairsWidth & "x" & oStairsAngle & oHandrailState & ".ipt"
        oFilename = lblFile.Text
    End Sub

    Private Sub get_LadderFilename()
        lblLadderFile.Text = "LADDER" & oLadderHeight & oCageState & ".ipt"
        oLadderFilename = lblLadderFile.Text
    End Sub

    Private Sub cboAngle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAngle.SelectedIndexChanged
        If cboAngle.Text = "37°" Then
            oStairsAngle = 37
        End If
        If cboAngle.Text = "38°" Then
            oStairsAngle = 38
        End If
        If cboAngle.Text = "39°" Then
            oStairsAngle = 39
        End If
        If cboAngle.Text = "40°" Then
            oStairsAngle = 40
        End If
        If cboAngle.Text = "41°" Then
            oStairsAngle = 41
        End If
        If cboAngle.Text = "42°" Then
            oStairsAngle = 42
        End If
        If cboAngle.Text = "43°" Then
            oStairsAngle = 43
        End If
        If cboAngle.Text = "44°" Then
            oStairsAngle = 44
        End If
        If cboAngle.Text = "45°" Then
            oStairsAngle = 45
        End If
        If cboWidth.Text = "46°" Then
            oStairsAngle = 46
        End If
        If cboAngle.Text = "47°" Then
            oStairsAngle = 47
        End If
        If cboAngle.Text = "48°" Then
            oStairsAngle = 48
        End If
        get_ProjectPath()
        get_Filename()
    End Sub

    Private Sub cboWidth_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboWidth.SelectedIndexChanged
        If cboWidth.Text = "600" Then
            oStairsWidth = 600
        End If
        If cboWidth.Text = "650" Then
            oStairsWidth = 650
        End If
        If cboWidth.Text = "700" Then
            oStairsWidth = 700
        End If
        If cboWidth.Text = "750" Then
            oStairsWidth = 750
        End If
        If cboWidth.Text = "800" Then
            oStairsWidth = 800
        End If
        If cboWidth.Text = "850" Then
            oStairsWidth = 850
        End If
        If cboWidth.Text = "900" Then
            oStairsWidth = 900
        End If
        If cboWidth.Text = "950" Then
            oStairsWidth = 950
        End If
        If cboWidth.Text = "1000" Then
            oStairsWidth = 1000
        End If
        get_ProjectPath()
        get_Filename()
    End Sub



    Private Sub get_Railing1Filename()
        lblRailFile.Text = "RAILING" & oRailingState & oRailLng1 & ".ipt"
        oRailFilename = lblRailFile.Text
    End Sub
    Private Sub get_Railing2Filename()
        lblRailFile.Text = "RAILING" & oRailingState & oRailLng1 & "x" & oRailLng2 & ".ipt"
        oRailFilename = lblRailFile.Text
    End Sub
    Private Sub get_Railing3Filename()
        lblRailFile.Text = "RAILING" & oRailingState & oRailLng1 & "x" & oRailLng2 & "x" & oRailLng3 & ".ipt"
        oRailFilename = lblRailFile.Text
    End Sub
    Private Sub get_Railing4Filename()
        lblRailFile.Text = "RAILING" & oRailingState & oRailLng1 & "x" & oRailLng2 & "x" & oRailLng3 & "x" & oRailLng4 & ".ipt"
        oRailFilename = lblRailFile.Text
    End Sub
    Private Sub get_Railing5Filename()
        lblRailFile.Text = "RAILING" & oRailingState & oRailLng1 & "x" & oRailLng2 & "x" & oRailLng3 & "x" & oRailLng4 & "x" & oRailLng5 & ".ipt"
        oRailFilename = lblRailFile.Text
    End Sub
    Private Sub get_RailingRFilename()
        lblRailFile.Text = "RAILING" & oRailingState & oRailDia & "x" & oRailAngle & ".ipt"
        oRailFilename = lblRailFile.Text
    End Sub
    Private Sub CboRailLng1_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboRailLng1.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboRailLng1.Text = Nothing Then
                CboRailLng1.Text = "300"
            End If
            If CboRailLng1.Text > 6000 Then
                oRailLng1 = 6000
                CboRailLng1.Text = 6000
            End If
            oRailLng1 = CDbl(CboRailLng1.Text)
            If oRailingState = "1Side" Then
                get_Railing1Filename()
            End If
            If oRailingState = "2Sides" Then
                get_Railing2Filename()
            End If
            If oRailingState = "3Sides" Then
                get_Railing3Filename()
            End If
            If oRailingState = "4Sides" Then
                get_Railing4Filename()
            End If
            If oRailingState = "5Sides" Then
                get_Railing5Filename()
            End If
        End If
    End Sub
    Private Sub CboRailLng2_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboRailLng2.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboRailLng2.Text = Nothing Then
                CboRailLng2.Text = "300"
            End If
            If CDbl(CboRailLng2.Text) > 6000 Then
                oRailLng2 = 6000
                CboRailLng2.Text = 6000
            End If
            oRailLng2 = CDbl(CboRailLng2.Text)
            If oRailingState = "1Side" Then
                get_Railing1Filename()
            End If
            If oRailingState = "2Sides" Then
                get_Railing2Filename()
            End If
            If oRailingState = "3Sides" Then
                get_Railing3Filename()
            End If
            If oRailingState = "4Sides" Then
                get_Railing4Filename()
            End If
            If oRailingState = "5Sides" Then
                get_Railing5Filename()
            End If
        End If
    End Sub
    Private Sub CboRailLng3_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboRailLng3.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboRailLng3.Text = Nothing Then
                CboRailLng3.Text = "300"
            End If
            If CDbl(CboRailLng3.Text) > 6000 Then
                oRailLng3 = 6000
                CboRailLng3.Text = 6000
            End If
            oRailLng3 = CDbl(CboRailLng3.Text)
            If oRailingState = "1Side" Then
                get_Railing1Filename()
            End If
            If oRailingState = "2Sides" Then
                get_Railing2Filename()
            End If
            If oRailingState = "3Sides" Then
                get_Railing3Filename()
            End If
            If oRailingState = "4Sides" Then
                get_Railing4Filename()
            End If
            If oRailingState = "5Sides" Then
                get_Railing5Filename()
            End If
        End If
    End Sub
    Private Sub CboRailLng4_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboRailLng4.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboRailLng4.Text = Nothing Then
                CboRailLng4.Text = "300"
            End If
            If CDbl(CboRailLng4.Text) > 6000 Then
                oRailLng4 = 6000
                CboRailLng4.Text = 6000
            End If
            oRailLng4 = CDbl(CboRailLng4.Text)
            If oRailingState = "1Side" Then
                get_Railing1Filename()
            End If
            If oRailingState = "2Sides" Then
                get_Railing2Filename()
            End If
            If oRailingState = "3Sides" Then
                get_Railing3Filename()
            End If
            If oRailingState = "4Sides" Then
                get_Railing4Filename()
            End If
            If oRailingState = "5Sides" Then
                get_Railing5Filename()
            End If
        End If
    End Sub
    Private Sub CboRailLng5_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboRailLng5.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboRailLng5.Text = Nothing Then
                CboRailLng5.Text = "300"
            End If
            If CDbl(CboRailLng5.Text) > 6000 Then
                oRailLng5 = 6000
                CboRailLng5.Text = 6000
            End If
            oRailLng5 = CDbl(CboRailLng5.Text)
            If oRailingState = "1Side" Then
                get_Railing1Filename()
            End If
            If oRailingState = "2Sides" Then
                get_Railing2Filename()
            End If
            If oRailingState = "3Sides" Then
                get_Railing3Filename()
            End If
            If oRailingState = "4Sides" Then
                get_Railing4Filename()
            End If
            If oRailingState = "5Sides" Then
                get_Railing5Filename()
            End If
        End If
    End Sub
    Private Sub CboRailDia_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboRailDia.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboRailDia.Text = Nothing Then
                CboRailDia.Text = "1200"
            End If
            If CboRailDia.Text > 25000 Then
                oRailDia = 25000
                CboRailDia.Text = 25000
            End If
            oRailDia = CDbl(CboRailDia.Text)
            get_RailingRFilename()
        End If
    End Sub
    Private Sub CboRailAngle_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboRailAngle.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboRailAngle.Text = Nothing Then
                CboRailAngle.Text = "90"
            End If
            If CboRailAngle.Text > 350 Then
                CboRailAngle.Text = 350
            End If
            If CboRailAngle.Text < 30 Then
                CboRailAngle.Text = 30
            End If
            oRailAngle = CDbl(CboRailAngle.Text)
            get_RailingRFilename()
        End If
    End Sub
    Private Sub CboPlatformAngle_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboPlatformAngle.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboPlatformAngle.Text = Nothing Then
                CboPlatformAngle.Text = "300"
            End If
            If CboPlatformAngle.Text > 350 Then
                CboPlatformAngle.Text = 350
            End If
            If CboPlatformAngle.Text < 30 Then
                CboPlatformAngle.Text = 30
            End If
            oPlatformAngle = CDbl(CboPlatformAngle.Text)
            get_PlatformFilename()
        End If
    End Sub
    Private Sub CboPlatformDiameter_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboPlatformDiameter.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboPlatformDiameter.Text = Nothing Then
                CboPlatformDiameter.Text = "1200"
            End If
            If CboPlatformDiameter.Text > 25000 Then
                CboPlatformDiameter.Text = 25000
            End If
            If CboPlatformDiameter.Text < 1200 Then
                CboPlatformDiameter.Text = 1200
            End If
            oPlatformDiameter = CDbl(CboPlatformDiameter.Text)
            get_PlatformFilename()
        End If
    End Sub

    Private Sub CboPlatformHeight_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboPlatformHeight.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboPlatformHeight.Text = Nothing Then
                CboPlatformHeight.Text = 600
            End If
            oPlatformHeight = CDbl(CboPlatformHeight.Text)
            If oPlatformHeight < 600 Then
                oPlatformHeight = 600
                CboPlatformHeight.Text = 600
            End If
            If oPlatformHeight > 9000 Then
                oPlatformHeight = 9000
                CboPlatformHeight.Text = 9000
            End If
            get_PlatformFilename()
            get_ProjectPath()
        End If
    End Sub
    Private Sub CboPlatformLenght_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboPlatformLenght.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboPlatformLenght.Text = Nothing Then
                CboPlatformLenght.Text = 800
            End If
            oPlatformLenght = CDbl(CboPlatformLenght.Text)
            If oPlatformLenght < 800 Then
                oPlatformLenght = 800
                CboPlatformLenght.Text = 800
            End If
            If oPlatformLenght > 9000 Then
                oPlatformLenght = 9000
                CboPlatformLenght.Text = 9000
            End If
            get_PlatformFilename()
            get_ProjectPath()
        End If
    End Sub
    Private Sub CboPlatformWidth_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CboPlatformWidth.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If CboPlatformWidth.Text = Nothing Then
                CboPlatformWidth.Text = 90
            End If
            oPlatformWidth = CDbl(CboPlatformWidth.Text)
            If oPlatformWidth < 800 Then
                oPlatformWidth = 800
                CboPlatformWidth.Text = 800
            End If
            If oPlatformWidth > 9000 Then
                oPlatformWidth = 9000
                CboPlatformWidth.Text = 9000
            End If
            If oPlatformState = "CURVED" And oPlatformWidth > oPlatformDiameter / 2 - 500 Then
                oPlatformWidth = 880
                CboPlatformWidth.Text = 880
            End If
            get_PlatformFilename()
            get_ProjectPath()
        End If
    End Sub

    Private Sub RbnRailLng1_CheckedChanged(sender As Object, e As EventArgs) Handles RbnRailLng1.CheckedChanged
        If RbnRailLng1.Checked = True Then
            PicRailR.Visible = False
            PicRail01.Visible = True
            PicRail02.Visible = False
            PicRail03.Visible = False
            PicRail04.Visible = False
            PicRail05.Visible = False
            oRailingState = "1Side"
            CboRailLng1.Visible = True
            CboRailLng2.Visible = False
            CboRailLng3.Visible = False
            CboRailLng4.Visible = False
            CboRailLng5.Visible = False
            CboRailDia.Visible = False
            CboRailAngle.Visible = False
            get_Railing1Filename()
        End If
    End Sub
    Private Sub RbnRailLng2_CheckedChanged(sender As Object, e As EventArgs) Handles RbnRailLng2.CheckedChanged
        If RbnRailLng2.Checked = True Then
            PicRailR.Visible = False
            PicRail01.Visible = False
            PicRail02.Visible = True
            PicRail03.Visible = False
            PicRail04.Visible = False
            PicRail05.Visible = False
            oRailingState = "2Sides"
            CboRailLng1.Visible = True
            CboRailLng2.Visible = True
            CboRailLng3.Visible = False
            CboRailLng4.Visible = False
            CboRailLng5.Visible = False
            CboRailDia.Visible = False
            CboRailAngle.Visible = False
            get_Railing2Filename()
        End If
    End Sub
    Private Sub RbnRailLng3_CheckedChanged(sender As Object, e As EventArgs) Handles RbnRailLng3.CheckedChanged
        If RbnRailLng3.Checked = True Then
            PicRailR.Visible = False
            PicRail01.Visible = False
            PicRail02.Visible = False
            PicRail03.Visible = True
            PicRail04.Visible = False
            PicRail05.Visible = False
            oRailingState = "3Sides"
            CboRailLng1.Visible = True
            CboRailLng2.Visible = True
            CboRailLng3.Visible = True
            CboRailLng4.Visible = False
            CboRailLng5.Visible = False
            CboRailDia.Visible = False
            CboRailAngle.Visible = False
            get_Railing4Filename()
        End If
    End Sub
    Private Sub RbnRailLng4_CheckedChanged(sender As Object, e As EventArgs) Handles RbnRailLng4.CheckedChanged
        If RbnRailLng4.Checked = True Then
            PicRailR.Visible = False
            PicRail01.Visible = False
            PicRail02.Visible = False
            PicRail03.Visible = False
            PicRail04.Visible = True
            PicRail05.Visible = False
            oRailingState = "4Sides"
            CboRailLng1.Visible = True
            CboRailLng2.Visible = True
            CboRailLng3.Visible = True
            CboRailLng4.Visible = True
            CboRailLng5.Visible = False
            CboRailDia.Visible = False
            CboRailAngle.Visible = False
            get_Railing4Filename()
        End If
    End Sub
    Private Sub RbnRailLng5_CheckedChanged(sender As Object, e As EventArgs) Handles RbnRailLng5.CheckedChanged
        If RbnRailLng5.Checked = True Then
            PicRailR.Visible = False
            PicRail01.Visible = False
            PicRail02.Visible = False
            PicRail03.Visible = False
            PicRail04.Visible = False
            PicRail05.Visible = True
            oRailingState = "5Sides"
            CboRailLng1.Visible = True
            CboRailLng2.Visible = True
            CboRailLng3.Visible = True
            CboRailLng4.Visible = True
            CboRailLng5.Visible = True
            CboRailDia.Visible = False
            CboRailAngle.Visible = False
            get_Railing5Filename()
        End If
    End Sub
    Private Sub RbnRailDia_CheckedChanged(sender As Object, e As EventArgs) Handles RbnRailDia.CheckedChanged
        If RbnRailDia.Checked = True Then
            PicRailR.Visible = True
            PicRail01.Visible = False
            PicRail02.Visible = False
            PicRail03.Visible = False
            PicRail04.Visible = False
            PicRail05.Visible = False
            oRailingState = "Round"
            CboRailLng1.Visible = False
            CboRailLng2.Visible = False
            CboRailLng3.Visible = False
            CboRailLng4.Visible = False
            CboRailLng5.Visible = False
            CboRailDia.Visible = True
            CboRailAngle.Visible = True
            get_RailingRFilename()
        End If
    End Sub

    Private Sub CboRailLng1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboRailLng1.SelectedIndexChanged, btnPlaceRailing.MouseHover
        oRailLng1 = CDbl(CboRailLng1.Text)

        RailingRule()

    End Sub
    Private Sub CboRailLng2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboRailLng2.SelectedIndexChanged, btnPlaceRailing.MouseHover
        oRailLng2 = CDbl(CboRailLng2.Text)

        RailingRule()

    End Sub
    Private Sub CboRailLng3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboRailLng3.SelectedIndexChanged, btnPlaceRailing.MouseHover
        oRailLng3 = CDbl(CboRailLng3.Text)
        RailingRule()

    End Sub
    Private Sub CboRailLng4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboRailLng4.SelectedIndexChanged, btnPlaceRailing.MouseHover
        oRailLng4 = CDbl(CboRailLng4.Text)

        RailingRule()

    End Sub
    Private Sub CboRailLng5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboRailLng5.SelectedIndexChanged, btnPlaceRailing.MouseHover
        oRailLng5 = CDbl(CboRailLng5.Text)

        RailingRule()

    End Sub
    Private Sub CboRailDia_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboRailDia.SelectedIndexChanged, btnPlaceRailing.MouseHover
        oRailDia = CDbl(CboRailDia.Text)

        RailingRule()

    End Sub
    Private Sub CboRailAngle_MouseLeave(sender As Object, e As EventArgs) Handles CboRailAngle.MouseLeave
        If CboRailAngle.Text = Nothing Then
            CboRailAngle.Text = 90
        End If
        oRailAngle = CDbl(CboRailAngle.Text)

        RailingRule()

    End Sub

    Private Sub CboRailAngle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboRailAngle.SelectedIndexChanged, btnPlaceRailing.MouseHover
        oRailAngle = CDbl(CboRailAngle.Text)

        RailingRule()

    End Sub


    Private Sub CboRailLng1_LostFocus(sender As Object, e As EventArgs) Handles CboRailLng1.LostFocus
        If CboRailLng1.Text = Nothing Then
            CboRailLng1.Text = 300
        End If
        oRailLng1 = CDbl(CboRailLng1.Text)

        RailingRule()

    End Sub


    Private Sub CboRailLng1_MouseLeave(sender As Object, e As EventArgs) Handles CboRailLng1.MouseLeave
        If CboRailLng1.Text = Nothing Then
            CboRailLng1.Text = 300
        End If
        oRailLng1 = CDbl(CboRailLng1.Text)

        RailingRule()

    End Sub
    Private Sub CboRailLng2_MouseLeave(sender As Object, e As EventArgs) Handles CboRailLng2.MouseLeave
        If CboRailLng2.Text = Nothing Then
            CboRailLng2.Text = 300
        End If
        oRailLng2 = CDbl(CboRailLng2.Text)

        RailingRule()

    End Sub

    Private Sub CboRailLng2_LostFocus(sender As Object, e As EventArgs) Handles CboRailLng2.LostFocus
        If CboRailLng2.Text = Nothing Then
            CboRailLng2.Text = 300
        End If
        oRailLng2 = CDbl(CboRailLng2.Text)

        RailingRule()

    End Sub





    Private Sub CboRailLng3_LostFocus(sender As Object, e As EventArgs) Handles CboRailLng3.LostFocus
        If CboRailLng3.Text = Nothing Then
            CboRailLng3.Text = 300
        End If
        oRailLng3 = CDbl(CboRailLng3.Text)

        RailingRule()

    End Sub


    Private Sub CboRailLng3_MouseLeave(sender As Object, e As EventArgs) Handles CboRailLng3.MouseLeave
        If CboRailLng3.Text = Nothing Then
            CboRailLng3.Text = 300
        End If
        oRailLng3 = CDbl(CboRailLng3.Text)

        RailingRule()

    End Sub




    Private Sub CboRailLng4_LostFocus(sender As Object, e As EventArgs) Handles CboRailLng4.LostFocus
        If CboRailLng4.Text = Nothing Then
            CboRailLng4.Text = 300
        End If
        oRailLng4 = CDbl(CboRailLng4.Text)

        RailingRule()

    End Sub




    Private Sub CboRailLng4_MouseLeave(sender As Object, e As EventArgs) Handles CboRailLng4.MouseLeave
        If CboRailLng4.Text = Nothing Then
            CboRailLng4.Text = 300
        End If
        oRailLng4 = CDbl(CboRailLng4.Text)

        RailingRule()

    End Sub



    Private Sub CboRailLng5_LostFocus(sender As Object, e As EventArgs) Handles CboRailLng5.LostFocus
        If CboRailLng5.Text = Nothing Then
            CboRailLng5.Text = 300
        End If
        oRailLng5 = CDbl(CboRailLng5.Text)

        RailingRule()

    End Sub





    Private Sub CboRailLng5_MouseLeave(sender As Object, e As EventArgs) Handles CboRailLng5.MouseLeave
        If CboRailLng5.Text = Nothing Then
            CboRailLng5.Text = 300
        End If
        oRailLng5 = CDbl(CboRailLng5.Text)

        RailingRule()

    End Sub





    Private Sub CboRailDia_LostFocus(sender As Object, e As EventArgs) Handles CboRailDia.LostFocus
        If CboRailDia.Text = Nothing Then
            CboRailDia.Text = 1200
        End If
        oRailDia = CDbl(CboRailDia.Text)

        RailingRule()

    End Sub



    Private Sub CboRailDia_MouseLeave(sender As Object, e As EventArgs) Handles CboRailDia.MouseLeave
        If CboRailDia.Text = Nothing Then
            CboRailDia.Text = 1200
        End If
        oRailDia = CDbl(CboRailDia.Text)

        RailingRule()


    End Sub


    '------------------------------------------------------------------
    '
    '
    '
    '
    '------------------------------PLATFORMS---------------------------
    '
    '
    '
    '
    '------------------------------------------------------------------
    Dim oPlatformLenght As Double
    Dim oPlatformWidth As Double
    Dim oPlatformFilename As String
    Dim oPlatformState As String
    Dim oPlatformHeight As Double
    Dim oPlatformDiameter As Double
    Dim oPlatformAngle As Double




    Private Sub RbnPlatformRect_CheckedChanged(sender As Object, e As EventArgs) Handles RbnPlatformRect.CheckedChanged
        If RbnPlatformRect.Checked = True Then
            lblPlatformDiameter.Visible = False
            CboPlatformDiameter.Visible = False
            lblPlatformAngle.Visible = False
            CboPlatformAngle.Visible = False
            lblPlatformHeight.Visible = True
            CboPlatformHeight.Visible = True
            lblPlatformLegNumber.Visible = False
            CboPlatformLegNumber.Visible = False
            lblPlatformLenght.Visible = True
            CboPlatformLenght.Visible = True
            lblPlatformWidth.Visible = True
            CboPlatformWidth.Visible = True
            CbxFoot.Visible = True
            CbxLegs.Visible = True
            If CbxLegs.Checked = True And CbxFoot.Checked = True Then
                PicPlatformRectLegsPlates.Visible = True
                PicPlatformRectLegs.Visible = False
                PicPlatformRect.Visible = False
                PicPlatformCircleLegsPlates.Visible = False
                PicPlatformCircleLegs.Visible = False
                PicPlatformCircle.Visible = False
                PicPlatformCurved.Visible = False
            End If
            If CbxLegs.Checked = True And CbxFoot.Checked = False Then
                PicPlatformRectLegsPlates.Visible = False
                PicPlatformRectLegs.Visible = True
                PicPlatformRect.Visible = False
                PicPlatformCircleLegsPlates.Visible = False
                PicPlatformCircleLegs.Visible = False
                PicPlatformCircle.Visible = False
                PicPlatformCurved.Visible = False
            End If
            If CbxLegs.Checked = False And CbxFoot.Checked = False Then
                PicPlatformRectLegsPlates.Visible = False
                PicPlatformRectLegs.Visible = False
                PicPlatformRect.Visible = True
                PicPlatformCircleLegsPlates.Visible = False
                PicPlatformCircleLegs.Visible = False
                PicPlatformCircle.Visible = False
                PicPlatformCurved.Visible = False
            End If
        End If
        If RbnPlatformRect.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = True Then
            oPlatformState = "RECT_LEGS_FOOT"
        End If
        If RbnPlatformRect.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = False Then
            oPlatformState = "RECT_LEGS"
        End If
        If RbnPlatformRect.Checked = True And CbxLegs.Checked = False Then
            oPlatformState = "RECT"
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub

    Private Sub CbxLegs_CheckedChanged(sender As Object, e As EventArgs) Handles CbxLegs.CheckedChanged
        If CbxLegs.Checked = True Then
            CbxFoot.Enabled = True
        End If
        If CbxLegs.Checked = False Then
            CbxFoot.Checked = False
            CbxFoot.Enabled = False
        End If
        If CbxLegs.Checked = True And RbnPlatformRect.Checked = True Then
            PicPlatformRectLegsPlates.Visible = False
            PicPlatformRectLegs.Visible = True
            PicPlatformRect.Visible = False
            PicPlatformCircleLegsPlates.Visible = False
            PicPlatformCircleLegs.Visible = False
            PicPlatformCircle.Visible = False
            PicPlatformCurved.Visible = False
        End If
        If CbxLegs.Checked = False And RbnPlatformRect.Checked = True Then
            PicPlatformRectLegsPlates.Visible = False
            PicPlatformRectLegs.Visible = False
            PicPlatformRect.Visible = True
            PicPlatformCircleLegsPlates.Visible = False
            PicPlatformCircleLegs.Visible = False
            PicPlatformCircle.Visible = False
            PicPlatformCurved.Visible = False
        End If
        If CbxLegs.Checked = True And RbnPlatformCirc.Checked = True Then
            PicPlatformRectLegsPlates.Visible = False
            PicPlatformRectLegs.Visible = False
            PicPlatformRect.Visible = False
            PicPlatformCircleLegsPlates.Visible = False
            PicPlatformCircleLegs.Visible = True
            PicPlatformCircle.Visible = False
            PicPlatformCurved.Visible = False
        End If
        If CbxLegs.Checked = False And RbnPlatformCirc.Checked = True Then
            PicPlatformRectLegsPlates.Visible = False
            PicPlatformRectLegs.Visible = False
            PicPlatformRect.Visible = False
            PicPlatformCircleLegsPlates.Visible = False
            PicPlatformCircleLegs.Visible = False
            PicPlatformCircle.Visible = True
            PicPlatformCurved.Visible = False
        End If
        If RbnPlatformRect.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = True Then
            oPlatformState = "RECT_LEGS_FOOT"
        End If
        If RbnPlatformRect.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = False Then
            oPlatformState = "RECT_LEGS"
        End If
        If RbnPlatformRect.Checked = True And CbxLegs.Checked = False Then
            oPlatformState = "RECT"
        End If
        If RbnPlatformCirc.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = True Then
            oPlatformState = "CIRC_LEGS_FOOT"
        End If
        If RbnPlatformCirc.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = False Then
            oPlatformState = "CIRC_LEGS"
        End If
        If RbnPlatformCirc.Checked = True And CbxLegs.Checked = False Then
            oPlatformState = "CIRC"
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub

    Private Sub CbxFoot_CheckedChanged(sender As Object, e As EventArgs) Handles CbxFoot.CheckedChanged
        If CbxLegs.Checked = False Then
            CbxFoot.Checked = False
        End If
        If CbxFoot.Checked = True And RbnPlatformRect.Checked = True Then
            PicPlatformRectLegsPlates.Visible = True
            PicPlatformRectLegs.Visible = False
            PicPlatformRect.Visible = False
            PicPlatformCircleLegsPlates.Visible = False
            PicPlatformCircleLegs.Visible = False
            PicPlatformCircle.Visible = False
            PicPlatformCurved.Visible = False
        End If
        If CbxFoot.Checked = False And RbnPlatformRect.Checked = True Then
            PicPlatformRectLegsPlates.Visible = False
            PicPlatformRectLegs.Visible = True
            PicPlatformRect.Visible = False
            PicPlatformCircleLegsPlates.Visible = False
            PicPlatformCircleLegs.Visible = False
            PicPlatformCircle.Visible = False
            PicPlatformCurved.Visible = False
        End If
        If CbxFoot.Checked = True And RbnPlatformCirc.Checked = True Then
            PicPlatformRectLegsPlates.Visible = False
            PicPlatformRectLegs.Visible = False
            PicPlatformRect.Visible = False
            PicPlatformCircleLegsPlates.Visible = True
            PicPlatformCircleLegs.Visible = False
            PicPlatformCircle.Visible = False
            PicPlatformCurved.Visible = False
        End If
        If CbxFoot.Checked = False And RbnPlatformCirc.Checked = True Then
            PicPlatformRectLegsPlates.Visible = False
            PicPlatformRectLegs.Visible = False
            PicPlatformRect.Visible = False
            PicPlatformCircleLegsPlates.Visible = False
            PicPlatformCircleLegs.Visible = True
            PicPlatformCircle.Visible = False
            PicPlatformCurved.Visible = False
        End If
        If RbnPlatformRect.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = True Then
            oPlatformState = "RECT_LEGS_FOOT"
        End If
        If RbnPlatformRect.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = False Then
            oPlatformState = "RECT_LEGS"
        End If
        If RbnPlatformRect.Checked = True And CbxLegs.Checked = False Then
            oPlatformState = "RECT"
        End If
        If RbnPlatformCirc.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = True Then
            oPlatformState = "CIRC_LEGS_FOOT"
        End If
        If RbnPlatformCirc.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = False Then
            oPlatformState = "CIRC_LEGS"
        End If
        If RbnPlatformCirc.Checked = True And CbxLegs.Checked = False Then
            oPlatformState = "CIRC"
        End If
        If RbnPlatformCurved.Checked = True Then
            oPlatformState = "CURVED"
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub

    Private Sub RbnPlatformCirc_CheckedChanged(sender As Object, e As EventArgs) Handles RbnPlatformCirc.CheckedChanged
        If RbnPlatformCirc.Checked = True Then
            lblPlatformDiameter.Visible = True
            CboPlatformDiameter.Visible = True
            lblPlatformAngle.Visible = False
            CboPlatformAngle.Visible = False
            lblPlatformHeight.Visible = True
            CboPlatformHeight.Visible = True
            lblPlatformLegNumber.Visible = True
            CboPlatformLegNumber.Visible = True
            lblPlatformLenght.Visible = False
            CboPlatformLenght.Visible = False
            lblPlatformWidth.Visible = False
            CboPlatformWidth.Visible = False
            CbxFoot.Visible = True
            CbxLegs.Visible = True
            If CbxLegs.Checked = True And CbxFoot.Checked = True Then
                PicPlatformRectLegsPlates.Visible = False
                PicPlatformRectLegs.Visible = False
                PicPlatformRect.Visible = False
                PicPlatformCircleLegsPlates.Visible = True
                PicPlatformCircleLegs.Visible = False
                PicPlatformCircle.Visible = False
                PicPlatformCurved.Visible = False
            End If
            If CbxLegs.Checked = True And CbxFoot.Checked = False Then
                PicPlatformRectLegsPlates.Visible = False
                PicPlatformRectLegs.Visible = False
                PicPlatformRect.Visible = False
                PicPlatformCircleLegsPlates.Visible = False
                PicPlatformCircleLegs.Visible = True
                PicPlatformCircle.Visible = False
                PicPlatformCurved.Visible = False
            End If
            If CbxLegs.Checked = False And CbxFoot.Checked = False Then
                PicPlatformRectLegsPlates.Visible = False
                PicPlatformRectLegs.Visible = False
                PicPlatformRect.Visible = False
                PicPlatformCircleLegsPlates.Visible = False
                PicPlatformCircleLegs.Visible = False
                PicPlatformCircle.Visible = True
                PicPlatformCurved.Visible = False
            End If
        End If
        If RbnPlatformCirc.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = True Then
            oPlatformState = "CIRC_LEGS_FOOT"
        End If
        If RbnPlatformCirc.Checked = True And CbxLegs.Checked = True And CbxFoot.Checked = False Then
            oPlatformState = "CIRC_LEGS"
        End If
        If RbnPlatformCirc.Checked = True And CbxLegs.Checked = False Then
            oPlatformState = "CIRC"
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub

    Private Sub RbnPlatformCurved_CheckedChanged(sender As Object, e As EventArgs) Handles RbnPlatformCurved.CheckedChanged
        If RbnPlatformCurved.Checked = True Then
            lblPlatformDiameter.Visible = True
            CboPlatformDiameter.Visible = True
            lblPlatformAngle.Visible = True
            CboPlatformAngle.Visible = True
            lblPlatformHeight.Visible = False
            CboPlatformHeight.Visible = False
            lblPlatformLegNumber.Visible = False
            CboPlatformLegNumber.Visible = False
            lblPlatformLenght.Visible = False
            CboPlatformLenght.Visible = False
            lblPlatformWidth.Visible = True
            CboPlatformWidth.Visible = True
            CbxFoot.Visible = False
            CbxLegs.Visible = False
            PicPlatformRectLegsPlates.Visible = False
            PicPlatformRectLegs.Visible = False
            PicPlatformRect.Visible = False
            PicPlatformCircleLegsPlates.Visible = False
            PicPlatformCircleLegs.Visible = False
            PicPlatformCircle.Visible = False
            PicPlatformCurved.Visible = True
        End If
        If RbnPlatformCurved.Checked = True Then
            oPlatformState = "CURVED"
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub
    Private Sub get_PlatformFilename()
        If oPlatformState = "CURVED" Then
            lblPlatformFile.Text = "PLATFORM_" & oPlatformState & "x" & oPlatformDiameter & "x" & oPlatformWidth & "x" & oPlatformAngle & ".ipt"
        End If
        If oPlatformState = "RECT_LEGS_FOOT" Then
            lblPlatformFile.Text = "PLATFORM_" & oPlatformState & "x" & oPlatformLenght & "x" & oPlatformWidth & "x" & oPlatformHeight & ".ipt"
        End If
        If oPlatformState = "RECT_LEGS" Then
            lblPlatformFile.Text = "PLATFORM_" & oPlatformState & "x" & oPlatformLenght & "x" & oPlatformWidth & "x" & oPlatformHeight & ".ipt"
        End If
        If oPlatformState = "RECT" Then
            lblPlatformFile.Text = "PLATFORM_" & oPlatformState & "x" & oPlatformLenght & "x" & oPlatformWidth & "x" & ".ipt"
        End If
        If oPlatformState = "CIRC_LEGS_FOOT" Then
            lblPlatformFile.Text = "PLATFORM_" & oPlatformState & "x" & oPlatformDiameter & "x" & oPlatformHeight & ".ipt"
        End If
        If oPlatformState = "CIRC_LEGS" Then
            lblPlatformFile.Text = "PLATFORM_" & oPlatformState & "x" & oPlatformDiameter & "x" & oPlatformHeight & ".ipt"
        End If
        If oPlatformState = "CIRC" Then
            lblPlatformFile.Text = "PLATFORM_" & oPlatformState & "x" & oPlatformDiameter & ".ipt"
        End If
        oPlatformFilename = lblPlatformFile.Text
    End Sub

    Private Sub CboPlatformLenght_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboPlatformLenght.SelectedIndexChanged, btnPlacePlatform.MouseHover
        oPlatformLenght = CDbl(CboPlatformLenght.Text)
        If oPlatformLenght < 800 Then
            oPlatformLenght = 800
            CboPlatformLenght.Text = 800
        End If
        If oPlatformLenght > 9000 Then
            oPlatformLenght = 9000
            CboPlatformLenght.Text = 9000
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub

    Private Sub CboPlatformLenght_MouseLeave(sender As Object, e As EventArgs) Handles CboPlatformLenght.MouseLeave
        If CboPlatformLenght.Text = Nothing Then
            CboPlatformLenght.Text = 800
        End If
        oPlatformLenght = CDbl(CboPlatformLenght.Text)
        If oPlatformLenght < 800 Then
            oPlatformLenght = 800
            CboPlatformLenght.Text = 800
        End If
        If oPlatformLenght > 9000 Then
            oPlatformLenght = 9000
            CboPlatformLenght.Text = 9000
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub


    Private Sub CboPlatformWidth_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboPlatformWidth.SelectedIndexChanged, btnPlacePlatform.MouseHover
        oPlatformWidth = CDbl(CboPlatformWidth.Text)
        If oPlatformWidth < 800 Then
            oPlatformWidth = 800
            CboPlatformWidth.Text = 800
        End If
        If oPlatformWidth > 9000 Then
            oPlatformWidth = 9000
            CboPlatformWidth.Text = 9000
        End If
        If oPlatformState = "CURVED" And oPlatformWidth > oPlatformDiameter / 2 - 500 Then
            oPlatformWidth = 880
            CboPlatformWidth.Text = 880
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub
    Private Sub CboPlatformWidth_MouseLeave(sender As Object, e As EventArgs) Handles CboPlatformWidth.MouseLeave
        If CboPlatformWidth.Text = Nothing Then
            CboPlatformWidth.Text = 90
        End If
        oPlatformWidth = CDbl(CboPlatformWidth.Text)
        If oPlatformWidth < 800 Then
            oPlatformWidth = 800
            CboPlatformWidth.Text = 800
        End If
        If oPlatformWidth > 9000 Then
            oPlatformWidth = 9000
            CboPlatformWidth.Text = 9000
        End If
        If oPlatformState = "CURVED" And oPlatformWidth > oPlatformDiameter / 2 - 500 Then
            oPlatformWidth = 880
            CboPlatformWidth.Text = 880
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub

    Private Sub CboPlatformDiameter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboPlatformDiameter.SelectedIndexChanged, btnPlacePlatform.MouseHover
        oPlatformDiameter = CDbl(CboPlatformDiameter.Text)
        If oPlatformDiameter < 1200 Then
            oPlatformDiameter = 1200
            CboPlatformDiameter.Text = 1200
        End If
        If oPlatformDiameter > 9000 Then
            oPlatformDiameter = 9000
            CboPlatformDiameter.Text = 9000
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub
    Private Sub CboPlatformDiameter_MouseLeave(sender As Object, e As EventArgs) Handles CboPlatformDiameter.MouseLeave
        If CboPlatformDiameter.Text = Nothing Then
            CboPlatformDiameter.Text = 1200
        End If
        oPlatformDiameter = CDbl(CboPlatformDiameter.Text)
        If oPlatformDiameter < 1200 Then
            oPlatformDiameter = 1200
            CboPlatformDiameter.Text = 1200
        End If
        If oPlatformDiameter > 9000 Then
            oPlatformDiameter = 9000
            CboPlatformDiameter.Text = 9000
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub

    Private Sub CboPlatformAngle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboPlatformAngle.SelectedIndexChanged, btnPlacePlatform.MouseHover
        oPlatformAngle = CDbl(CboPlatformAngle.Text)

        If oPlatformAngle > 350 Then
            oPlatformAngle = 350
            CboPlatformAngle.Text = 350
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub
    Private Sub CboPlatformAngle_MouseLeave(sender As Object, e As EventArgs) Handles CboPlatformAngle.MouseLeave
        If CboPlatformAngle.Text = Nothing Then
            CboPlatformAngle.Text = 90
        End If
        oPlatformAngle = CDbl(CboPlatformAngle.Text)
        If oPlatformAngle > 350 Then
            oPlatformAngle = 350
            CboPlatformAngle.Text = 350
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub
    Private Sub CboPlatformHeight_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboPlatformHeight.SelectedIndexChanged, btnPlacePlatform.MouseHover
        oPlatformHeight = CDbl(CboPlatformHeight.Text)
        If oPlatformHeight < 600 Then
            oPlatformHeight = 600
            CboPlatformHeight.Text = 600
        End If
        If oPlatformHeight > 9000 Then
            oPlatformHeight = 9000
            CboPlatformHeight.Text = 9000
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub
    Private Sub CboPlatformHeight_MouseLeave(sender As Object, e As EventArgs) Handles CboPlatformHeight.MouseLeave
        If CboPlatformHeight.Text = Nothing Then
            CboPlatformHeight.Text = 600
        End If
        oPlatformHeight = CDbl(CboPlatformHeight.Text)
        If oPlatformHeight < 600 Then
            oPlatformHeight = 600
            CboPlatformHeight.Text = 600
        End If
        If oPlatformHeight > 9000 Then
            oPlatformHeight = 9000
            CboPlatformHeight.Text = 9000
        End If
        get_PlatformFilename()
        get_ProjectPath()
    End Sub






    '---------------------------START CURVED STAIRS---------------------------------------------------
    '               ***
    '                       ***
    '                               ***
    '                                        ***
    '                                                  ***
    '
    '---------------------------START CURVED STAIRS---------------------------------------------------
    '
    '



    Private Sub get_CSFilename()
        lblCurvedStairsFile.Text = "STAIRS" & oCSHeight & "x" & oCSWidth & "x" & oCSIntRadius & "x" & oCSAngle & oCSState & ".ipt"
        oCSFilename = lblCurvedStairsFile.Text
    End Sub

    Private Sub cboCurvedStairsHeight_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles cboCurvedStairsHeight.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If cboCurvedStairsHeight.Text = Nothing Then
                cboCurvedStairsHeight.Text = "600"
            End If
            If cboHeight.Text < 600 Then
                oCSHeight = 600
                cboCurvedStairsHeight.Text = 600
            End If
            If cboCurvedStairsHeight.Text > 6000 Then
                oCSHeight = 6000
                cboCurvedStairsHeight.Text = 6000
            End If
            oCSHeight = CDbl(cboCurvedStairsHeight.Text)
        End If
        get_CSFilename()
    End Sub
    Private Sub cboCurvedStairsHeight_MouseLeave(sender As Object, e As EventArgs) Handles cboCurvedStairsHeight.MouseLeave
        If cboCurvedStairsHeight.Text = Nothing Then
            cboCurvedStairsHeight.Text = 600
        End If
        oCSHeight = CDbl(cboCurvedStairsHeight.Text)
        If oCSHeight < 600 Then
            oCSHeight = 600
            cboCurvedStairsHeight.Text = 600
        End If
        If oCSHeight > 6000 Then
            oCSHeight = 6000
            cboCurvedStairsHeight.Text = 6000
        End If
        get_ProjectPath()
        get_CSFilename()
    End Sub
    Private Sub cboCurvedStairsHeight_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCurvedStairsHeight.SelectedIndexChanged
        oCSHeight = CDbl(cboCurvedStairsHeight.Text)
        If oCSHeight < 600 Then
            oCSHeight = 600
            cboCurvedStairsHeight.Text = 600
        End If
        If oCSHeight > 6000 Then
            oCSHeight = 6000
            cboCurvedStairsHeight.Text = 6000
        End If
        get_ProjectPath()
        get_CSFilename()
    End Sub

    Private Sub cboCurvedStairsHeight_LostFocus(sender As Object, e As EventArgs) Handles cboCurvedStairsHeight.LostFocus
        If cboCurvedStairsHeight.Text = Nothing Then
            cboCurvedStairsHeight.Text = 600
        End If
        oCSHeight = CDbl(cboCurvedStairsHeight.Text)
        If oCSHeight < 600 Then
            oCSHeight = 600
            cboCurvedStairsHeight.Text = 600
        End If
        If oCSHeight > 6000 Then
            oCSHeight = 6000
            cboCurvedStairsHeight.Text = 6000
        End If
        get_ProjectPath()
        get_CSFilename()
    End Sub


    Private Sub cboCurvedStairsWidth_MouseLeave(sender As Object, e As EventArgs) Handles cboCurvedStairsWidth.MouseLeave
        If cboCurvedStairsWidth.Text = Nothing Then
            cboCurvedStairsWidth.Text = 600
        End If
        oCSWidth = CDbl(cboCurvedStairsWidth.Text)
        If oCSWidth < 600 Then
            oCSWidth = 600
            cboCurvedStairsWidth.Text = 600
        End If
        If oCSWidth > 1200 Then
            oCSWidth = 1200
            cboCurvedStairsWidth.Text = 1200
        End If
        get_ProjectPath()
        get_CSFilename()
    End Sub
    Private Sub cboCurvedStairsWidth_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCurvedStairsWidth.SelectedIndexChanged
        oCSWidth = CDbl(cboCurvedStairsWidth.Text)
        If oCSWidth < 600 Then
            oCSWidth = 600
            cboCurvedStairsWidth.Text = 600
        End If
        If oCSWidth > 1200 Then
            oCSWidth = 1200
            cboCurvedStairsWidth.Text = 1200
        End If
        get_ProjectPath()
        get_CSFilename()
    End Sub
    Private Sub cboCurvedStairsIntRadius_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles cboCurvedStairsIntRadius.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
        If Asc(e.KeyChar) = 22 Then
            e.KeyChar = ""
        End If
        If Asc(e.KeyChar) = 13 Then
            If cboCurvedStairsIntRadius.Text = Nothing Then
                cboCurvedStairsIntRadius.Text = "600"
            End If
            If cboCurvedStairsIntRadius.Text < 600 Then
                oCSIntRadius = 600
                cboCurvedStairsIntRadius.Text = 600
            End If
            If cboCurvedStairsIntRadius.Text > 15000 Then
                oCSIntRadius = 15000
                cboCurvedStairsIntRadius.Text = 15000
            End If
            oCSIntRadius = CDbl(cboCurvedStairsIntRadius.Text)
        End If
        get_CSFilename()
    End Sub
    Private Sub cboCurvedStairsRadius_MouseLeave(sender As Object, e As EventArgs) Handles cboCurvedStairsIntRadius.MouseLeave
        If cboCurvedStairsIntRadius.Text = Nothing Then
            cboCurvedStairsIntRadius.Text = 600
        End If
        oCSIntRadius = CDbl(cboCurvedStairsIntRadius.Text)
        If oCSIntRadius < 600 Then
            oCSIntRadius = 600
            cboCurvedStairsIntRadius.Text = 600
        End If
        If oCSIntRadius > 15000 Then
            oCSIntRadius = 15000
            cboCurvedStairsIntRadius.Text = 15000
        End If
        get_ProjectPath()
        get_CSFilename()
    End Sub
    Private Sub cboCurvedStairsRadius_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCurvedStairsIntRadius.SelectedIndexChanged
        oCSIntRadius = CDbl(cboCurvedStairsIntRadius.Text)
        If oCSIntRadius < 600 Then
            oCSIntRadius = 600
            cboCurvedStairsIntRadius.Text = 600
        End If
        If oCSIntRadius > 15000 Then
            oCSIntRadius = 15000
            cboCurvedStairsIntRadius.Text = 15000
        End If
        get_ProjectPath()
        get_CSFilename()
    End Sub


    Private Sub cboCurvedStairsIntRadius_LostFocus(sender As Object, e As EventArgs) Handles cboCurvedStairsIntRadius.LostFocus
        If cboCurvedStairsIntRadius.Text = Nothing Then
            cboCurvedStairsIntRadius.Text = 600
        End If
        oCSIntRadius = CDbl(cboCurvedStairsIntRadius.Text)
        If oCSIntRadius < 600 Then
            oCSIntRadius = 600
            cboCurvedStairsIntRadius.Text = 600
        End If
        If oCSIntRadius > 15000 Then
            oCSIntRadius = 15000
            cboCurvedStairsIntRadius.Text = 15000
        End If
        get_ProjectPath()
        get_CSFilename()
    End Sub


    Private Sub cboCurvedStairsAngle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCurvedStairsAngle.SelectedIndexChanged
        If cboCurvedStairsAngle.Text = "37°" Then
            oCSAngle = 37
        End If
        If cboCurvedStairsAngle.Text = "38°" Then
            oCSAngle = 38
        End If
        If cboCurvedStairsAngle.Text = "39°" Then
            oCSAngle = 39
        End If
        If cboCurvedStairsAngle.Text = "40°" Then
            oCSAngle = 40
        End If
        If cboCurvedStairsAngle.Text = "41°" Then
            oCSAngle = 41
        End If
        If cboCurvedStairsAngle.Text = "42°" Then
            oCSAngle = 42
        End If
        If cboCurvedStairsAngle.Text = "43°" Then
            oCSAngle = 43
        End If
        If cboCurvedStairsAngle.Text = "44°" Then
            oCSAngle = 44
        End If
        If cboCurvedStairsAngle.Text = "45°" Then
            oCSAngle = 45
        End If
        If cboCurvedStairsAngle.Text = "46°" Then
            oCSAngle = 46
        End If
        If cboCurvedStairsAngle.Text = "47°" Then
            oCSAngle = 47
        End If
        If cboCurvedStairsAngle.Text = "48°" Then
            oCSAngle = 48
        End If
        get_ProjectPath()
        get_CSFilename()
    End Sub

    Private Sub chkRightHandrailCurvedStairs_CheckedChanged(sender As Object, e As EventArgs) Handles chkRightHandrailCurvedStairs.CheckedChanged
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = True
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Both"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = True
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CCW_Both"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = True
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Left"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = True
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CCW_Left"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = True
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Right"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = True
            picCCW_None.Visible = False
            oCSState = "CCW_Right"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = True
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_None"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = True
            oCSState = "CCW_None"
        End If
        get_CSFilename()
    End Sub

    Private Sub rbnClockWise_CheckedChanged(sender As Object, e As EventArgs) Handles rbnClockWise.CheckedChanged
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = True
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Both"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = True
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CCW_Both"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = True
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Left"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = True
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CCW_Left"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = True
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Right"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = True
            picCCW_None.Visible = False
            oCSState = "CCW_Right"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = True
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_None"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = True
            oCSState = "CCW_None"
        End If
        get_CSFilename()
    End Sub

    Private Sub rbnCounterClockWise_CheckedChanged(sender As Object, e As EventArgs) Handles rbnCounterClockWise.CheckedChanged
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = True
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Both"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = True
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CCW_Both"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = True
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Left"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = True
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CCW_Left"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = True
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Right"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = True
            picCCW_None.Visible = False
            oCSState = "CCW_Right"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = True
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_None"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = True
            oCSState = "CCW_None"
        End If
        get_CSFilename()
    End Sub

    Private Sub chkLeftHandrailCurvedStairs_CheckedChanged(sender As Object, e As EventArgs) Handles chkLeftHandrailCurvedStairs.CheckedChanged
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = True
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Both"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = True
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CCW_Both"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = True
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Left"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 1 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = True
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CCW_Left"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = True
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_Right"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 1 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = True
            picCCW_None.Visible = False
            oCSState = "CCW_Right"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = True
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = False
            oCSState = "CW_None"
        End If
        If chkRightHandrailCurvedStairs.CheckState = 0 AndAlso chkLeftHandrailCurvedStairs.CheckState = 0 AndAlso rbnCounterClockWise.Checked = True Then
            picCW_Both.Visible = False
            picCW_Left.Visible = False
            picCW_Right.Visible = False
            picCW_None.Visible = False
            picCCW_Both.Visible = False
            picCCW_Left.Visible = False
            picCCW_Right.Visible = False
            picCCW_None.Visible = True
            oCSState = "CCW_None"
        End If
        get_CSFilename()
    End Sub





    Private Sub StairsTab_MouseMove(sender As Object, e As MouseEventArgs) Handles StairsTab.MouseMove
        If cboHeight.Text = Nothing Then
            cboHeight.Text = 600
        End If
        oStairsHeight = CDbl(cboHeight.Text)
        If oStairsHeight < 600 Then
            oStairsHeight = 600
            cboHeight.Text = 600
        End If
        If oStairsHeight > 6000 Then
            oStairsHeight = 6000
            cboHeight.Text = 6000
        End If
        get_ProjectPath()
        get_Filename()
    End Sub



    Private Sub CurvedStairsTab_MouseMove(sender As Object, e As MouseEventArgs) Handles CurvedStairsTab.MouseMove
        If cboCurvedStairsHeight.Text = Nothing Then
            cboCurvedStairsHeight.Text = 600
        End If
        oCSHeight = CDbl(cboCurvedStairsHeight.Text)
        If oCSHeight < 600 Then
            oCSHeight = 600
            cboCurvedStairsHeight.Text = 600
        End If
        If oCSHeight > 6000 Then
            oCSHeight = 6000
            cboCurvedStairsHeight.Text = 6000
        End If
        If cboCurvedStairsIntRadius.Text = Nothing Then
            cboCurvedStairsIntRadius.Text = 600
        End If
        oCSIntRadius = CDbl(cboCurvedStairsIntRadius.Text)
        If oCSIntRadius < 600 Then
            oCSIntRadius = 600
            cboCurvedStairsIntRadius.Text = 600
        End If
        If oCSIntRadius > 15000 Then
            oCSIntRadius = 15000
            cboCurvedStairsIntRadius.Text = 15000
        End If
        get_ProjectPath()
        get_CSFilename()

    End Sub

    Private Sub LadderTab_MouseMove(sender As Object, e As MouseEventArgs) Handles LadderTab.MouseMove
        If cboLadderHeight.Text = Nothing Then
            cboLadderHeight.Text = 1200
        End If
        oLadderHeight = CDbl(cboLadderHeight.Text)
        If oLadderHeight < 1200 Then
            oLadderHeight = 1200
            cboLadderHeight.Text = 1200
        End If
        If oLadderHeight < 2300 Then
            oCageState = "NoCage"
            chkCage.Enabled = False
            chkCage.Checked = False
        End If

        If oLadderHeight >= 2300 Then
            oCageState = "Cage"
            chkCage.Enabled = True
            chkCage.Checked = True

        End If
        If oLadderHeight > 6000 Then
            oLadderHeight = 6000
            cboLadderHeight.Text = 6000
        End If
        get_ProjectPath()
        get_LadderFilename()
    End Sub



    Private Sub RailingRule()
        Select Case oRailingState

            Case "1Side"
                If oRailLng1 > 20000 Then
                    oRailLng1 = 20000
                    CboRailLng1.Text = 20000
                End If
                If oRailLng1 < 600 Then
                    oRailLng1 = 600
                    CboRailLng1.Text = 600
                End If

                get_Railing1Filename()

            Case "2Sides"
                If oRailLng1 > 20000 Then
                    oRailLng1 = 20000
                    CboRailLng1.Text = 20000
                End If
                If oRailLng2 > 20000 Then
                    oRailLng2 = 20000
                    CboRailLng2.Text = 20000
                End If
                If oRailLng1 < 300 And oRailLng2 > 600 Then
                    oRailLng1 = 300
                    CboRailLng1.Text = 300
                End If
                If oRailLng1 < 600 And oRailLng2 < 600 Then
                    oRailLng1 = 600
                    CboRailLng1.Text = 600
                End If
                If oRailLng2 < 300 And oRailLng1 > 600 Then
                    oRailLng1 = 300
                    CboRailLng1.Text = 300
                End If
                If oRailLng2 < 600 And oRailLng1 < 600 Then
                    oRailLng2 = 600
                    CboRailLng2.Text = 600
                End If

                get_Railing2Filename()

            Case "3Sides"
                If oRailLng1 > 20000 Then
                    oRailLng1 = 20000
                    CboRailLng1.Text = 20000
                End If
                If oRailLng2 > 20000 Then
                    oRailLng2 = 20000
                    CboRailLng2.Text = 20000
                End If
                If oRailLng3 > 20000 Then
                    oRailLng3 = 20000
                    CboRailLng3.Text = 20000
                End If
                If oRailLng1 < 300 Then
                    oRailLng1 = 300
                    CboRailLng1.Text = 300
                End If
                If oRailLng2 < 600 Then
                    oRailLng2 = 600
                    CboRailLng2.Text = 600
                End If
                If oRailLng3 < 300 Then
                    oRailLng3 = 300
                    CboRailLng3.Text = 300
                End If

                get_Railing3Filename()

            Case "4Sides"

                If oRailLng1 > 20000 Then
                    oRailLng1 = 20000
                    CboRailLng1.Text = 20000
                End If
                If oRailLng2 > 20000 Then
                    oRailLng2 = 20000
                    CboRailLng2.Text = 20000
                End If
                If oRailLng3 > 20000 Then
                    oRailLng3 = 20000
                    CboRailLng3.Text = 20000
                End If
                If oRailLng4 > 20000 Then
                    oRailLng4 = 20000
                    CboRailLng4.Text = 20000
                End If
                If oRailLng1 < 300 Then
                    oRailLng1 = 300
                    CboRailLng1.Text = 300
                End If
                If oRailLng2 < 600 Then
                    oRailLng2 = 600
                    CboRailLng2.Text = 600
                End If
                If oRailLng3 < 600 Then
                    oRailLng3 = 600
                    CboRailLng3.Text = 600
                End If
                If oRailLng4 < 300 Then
                    oRailLng4 = 300
                    CboRailLng4.Text = 300
                End If

                If oRailLng1 + 50 > oRailLng3 And oRailLng4 + 50 > oRailLng2 Then
                    If oRailLng2 >= 1200 Then
                        oRailLng4 = oRailLng2 - 600
                        CboRailLng4.Text = oRailLng4
                    End If
                    If oRailLng2 < 1200 Then
                        oRailLng3 = oRailLng1 + 600
                        CboRailLng3.Text = oRailLng3
                    End If
                End If

                get_Railing4Filename()

            Case "5Sides"

                If oRailLng1 > 20000 Then
                    oRailLng1 = 20000
                    CboRailLng1.Text = 20000
                End If
                If oRailLng2 > 20000 Then
                    oRailLng2 = 20000
                    CboRailLng2.Text = 20000
                End If
                If oRailLng3 > 20000 Then
                    oRailLng3 = 20000
                    CboRailLng3.Text = 20000
                End If
                If oRailLng4 > 20000 Then
                    oRailLng4 = 20000
                    CboRailLng4.Text = 20000
                End If
                If oRailLng5 > 20000 Then
                    oRailLng5 = 20000
                    CboRailLng5.Text = 20000
                End If
                If oRailLng1 < 300 Then
                    oRailLng1 = 300
                    CboRailLng1.Text = 300
                End If
                If oRailLng2 < 600 Then
                    oRailLng2 = 600
                    CboRailLng2.Text = 600
                End If
                If oRailLng3 < 700 Then
                    oRailLng3 = 700
                    CboRailLng3.Text = 700
                End If
                If oRailLng4 < 600 Then
                    oRailLng4 = 600
                    CboRailLng4.Text = 600
                End If
                If oRailLng5 < 300 Then
                    oRailLng5 = 300
                    CboRailLng5.Text = 300
                End If

                If (oRailLng2 + 50) >= oRailLng4 And (oRailLng2 - 50) <= oRailLng4 Then
                    If (oRailLng1 + 50) + (oRailLng5 + 50) > oRailLng3 Then
                        oRailLng3 = oRailLng1 + oRailLng5 + 600
                        CboRailLng3.Text = oRailLng3
                    End If
                End If

                If oRailLng2 >= oRailLng4 And oRailLng5 + 50 > oRailLng3 Then
                    If oRailLng3 >= 1200 Then
                        oRailLng5 = oRailLng3 - 600
                        CboRailLng5.Text = oRailLng5
                    End If
                    If oRailLng3 <= 1200 Then
                        oRailLng3 = oRailLng5 + 600
                        CboRailLng3.Text = oRailLng3
                    End If
                End If

                If oRailLng4 >= oRailLng2 And oRailLng1 + 50 > oRailLng3 Then
                    If oRailLng3 >= 1200 Then
                        oRailLng1 = oRailLng3 - 600
                        CboRailLng1.Text = oRailLng1
                    End If
                    If oRailLng3 <= 1200 Then
                        oRailLng3 = oRailLng1 + 600
                        CboRailLng3.Text = oRailLng3
                    End If
                End If

                get_Railing5Filename()

            Case "Round"
                If oRailDia < 1200 Then
                    oRailDia = 1200
                    CboRailDia.Text = 1200
                End If
                If oRailDia > 30000 Then
                    oRailDia = 30000
                    CboRailDia.Text = 30000
                End If

                If CboRailAngle.Text > 350 Then
                    CboRailAngle.Text = 350
                End If
                If CboRailAngle.Text < 30 Then
                    CboRailAngle.Text = 30
                End If
                CboRailAngle.Text = oRailAngle

                get_RailingRFilename()


        End Select
    End Sub

    Private Sub btnPlaceStairsInAssem_Click(sender As Object, e As EventArgs) Handles btnPlaceStairsInAssem.Click
        PlaceStairs()
    End Sub


    Private Sub PlaceStairs()
        '------------------------------------------------
        '
        '          PLACING STAIRS IN AN ASSEMBLY
        '
        '------------------------------------------------

        Dim oDef As PartComponentDefinition
        Dim oTG As TransientGeometry
        Dim oCenterPlane As WorkPlane
        Dim oProfile As Profile
        Dim asmDoc As AssemblyDocument
        Dim oFloorFace As Face
        Dim oStairsDoc As PartDocument
        Dim attrSets As AttributeSetsEnumerator
        Dim stairsOcc As ComponentOccurrence
        Dim oRailingPlane As WorkPlane
        Dim oRightSidePlane As WorkPlane
        Dim propSet1 As PropertySet
        Dim propSet2 As PropertySet
        Dim propSet3 As PropertySet

        'Check of Inventor nog steeds actief is
        Try
            oInvApp = GetObject(, "Inventor.Application")
        Catch ex As Exception
            MsgBox("Inventor must be running!.", vbExclamation, "")
            btnPlaceStairsInAssem.Enabled = False
            btnPlaceLadder.Enabled = False
            btnPlaceRailing.Enabled = False
            btnPlacePlatform.Enabled = False
            btnPlaceCurvedStairs.Enabled = False
            Exit Sub
        End Try

        'Check of er een assembly open is
        If oInvApp.ActiveDocumentType = Global.Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            GoTo BeginRoutine
        Else MsgBox("Open, Create or Activate an Assembly !", vbExclamation, "")
            GoTo EndRoutine
        End If

BeginRoutine:
        If Dir(lblPath.Text & oFilename) = "" Then
            GoTo StartPart
        Else
            GoTo PlacePart
        End If

StartPart:
        ProgressBar1.Visible = True
        ProgressBar1.Value = 10
        oInvApp.ScreenUpdating = False

        oStairsDoc = oInvApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject)
        oStairsDoc = oInvApp.ActiveDocument
        oDef = oInvApp.ActiveDocument.ComponentDefinition
        oTG = oInvApp.TransientGeometry
        pi = Math.Acos(-1) 'radians van 45° hoek
        oAngle = (oStairsAngle * pi / 180)
        oNumberOfSteps = oStairsHeight / 200
        oSPAA = (((90 - oStairsAngle) / 2) * pi / 180)
        oStepheight = oStairsHeight / oNumberOfSteps
        oStairsHypotenusa = (oStairsHeight - oStepheight) / (Math.Sin(oAngle))
        oNumberOfSpils = Math.Round((oStairsHypotenusa / 1000), 0)

        If oNumberOfSpils = 1 Then
            oNumberOfSpils = 2
        End If


        propSet1 = oStairsDoc.PropertySets.Item("Inventor Summary Information")
        propSet1.ItemByPropId(2).Value = "Steel Stairs"
        propSet1.ItemByPropId(4).Value = "Marc Crauwels"
        propSet1.ItemByPropId(6).Value = "This Steel Staircase is generated with the Pantarein Water Panta Stairs application. This software is part of the Pantarein Water Design Suite."

        propSet2 = oStairsDoc.PropertySets.Item("Inventor Document Summary Information")
        propSet2.ItemByPropId(2).Value = "Plant Design Mechanical"
        propSet2.ItemByPropId(15).Value = "Pantarein Water BVBA"

        propSet3 = oStairsDoc.PropertySets.Item("Design Tracking Properties")
        propSet3.ItemByPropId(41).Value = "Marc Crauwels"
        propSet3.ItemByPropId(29).Value = "Steel Stairs"
        propSet3.ItemByPropId(23).Value = "http://www.pantareinwater.be/en"


        '----------Adjust Orientation XY is floor Z is up----------

        Dim v As Inventor.View
        v = oInvApp.ActiveView
        Dim c As Camera
        c = oInvApp.ActiveView.Camera
        Dim t2eDist As Double
        t2eDist = c.Target.DistanceTo(c.Eye)
        Dim t2e As Vector
        t2e = oTG.CreateVector(0, -t2eDist, 0)
        Dim newEye As Point
        newEye = c.Target.Copy
        Call newEye.TranslateBy(t2e)
        c.Eye = newEye
        c.UpVector = oTG.CreateUnitVector(0, 0, 1)
        c.ApplyWithoutTransition()
        v.SetCurrentAsFront()

        '-----------------SidePlane---------------------

        oCenterPlane = oDef.WorkPlanes.Item(1)
        oRightSidePlane = oDef.WorkPlanes.AddByPlaneAndOffset(oCenterPlane, (oStairsWidth + 150) / 2 / 10, False)
        '    oRightSidePlane.Name = "RightStairsSide"

        '----------------Sketch side----------------------

        Dim oSideSketch As PlanarSketch
        oSideSketch = oDef.Sketches.Add(oRightSidePlane)
        Dim oWp01 As SketchPoints
        oWp01 = oSideSketch.SketchPoints

        ProgressBar1.Value = 30

        '------------Sketch points side----------------------

        Call oWp01.Add(oTG.CreatePoint2d(0, 0), False)
        Call oWp01.Add(oTG.CreatePoint2d(0, (oStepheight / 10)), False)
        Call oWp01.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)), (oStairsHeight / 10)), False) 'X coordinaat = Traphoogte-profielhoogte gedeeld door 10 voor mm naar cm gedeeld door de tangens van de traphoek
        Call oWp01.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 23, (oStairsHeight / 10)), False)
        Call oWp01.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 23, (oStairsHeight / 10) - 20), False)
        Call oWp01.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 23 - (((Math.Atan(oAngle)) * 200) / 10), (oStairsHeight / 10) - 20), False)
        Call oWp01.Add(oTG.CreatePoint2d(20, (oStepheight - (Math.Tan(oSPAA)) * 200) / 10), False)
        Call oWp01.Add(oTG.CreatePoint2d(20, 0), False)
        '------------Sketch points grouded for later constraints-------------------

        Call oSideSketch.GeometricConstraints.AddGround(oWp01(1))
        Call oSideSketch.GeometricConstraints.AddGround(oWp01(2))
        Call oSideSketch.GeometricConstraints.AddGround(oWp01(3))
        Call oSideSketch.GeometricConstraints.AddGround(oWp01(4))
        Call oSideSketch.GeometricConstraints.AddGround(oWp01(5))
        Call oSideSketch.GeometricConstraints.AddGround(oWp01(7))
        Call oSideSketch.GeometricConstraints.AddGround(oWp01(8))

        '-------------Draw lines----------------------------------------------------

        Dim oLn01 As SketchLine
        oLn01 = oSideSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
        Dim oLn02 As SketchLine
        oLn02 = oSideSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
        Dim oLn03 As SketchLine
        oLn03 = oSideSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
        Dim oLn04 As SketchLine
        oLn04 = oSideSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(5))
        Dim oLn05 As SketchLine
        oLn05 = oSideSketch.SketchLines.AddByTwoPoints(oWp01(5), oWp01(6))
        Dim oLn06 As SketchLine
        oLn06 = oSideSketch.SketchLines.AddByTwoPoints(oWp01(6), oWp01(7))
        Dim oLn07 As SketchLine
        oLn07 = oSideSketch.SketchLines.AddByTwoPoints(oWp01(7), oWp01(8))
        Dim oLn08 As SketchLine
        oLn08 = oSideSketch.SketchLines.AddByTwoPoints(oWp01(8), oWp01(1))

        Call oSideSketch.GeometricConstraints.AddParallel(oLn03, oLn05)
        Call oSideSketch.GeometricConstraints.AddParallel(oLn02, oLn06)


        '--------------Extrusion of the sidebeam--------------------------------------

        oProfile = oSideSketch.Profiles.AddForSolid
        Dim oExtrudeDef As ExtrudeDefinition
        Dim kJoinOperation As Object = Nothing
        oExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kJoinOperation)
        Call oExtrudeDef.SetDistanceExtent(7.5, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        Dim oExtrude As ExtrudeFeature
        oExtrude = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)

        'start---------Maak een AttributeSet "ConnectFace" ----------

        Dim entity As Face
        entity = oDef.SurfaceBodies.Item(1).Faces.Item(1)
        Dim attribSets As AttributeSets
        attribSets = entity.AttributeSets
        Dim attribSet As AttributeSet
        attribSet = attribSets.Add("ConnectFace")

        'end---------Maak een AttributeSet "ConnectFace" ----------

        Dim oSideSketch2 As PlanarSketch
        oSideSketch2 = oDef.Sketches.Add(oRightSidePlane)
        Dim oWp02 As SketchPoints
        oWp02 = oSideSketch2.SketchPoints
        Call oWp02.Add(oTG.CreatePoint2d(1, 1), False)
        Call oWp02.Add(oTG.CreatePoint2d(1, (oStepheight / 10) - (Math.Atan(oSPAA) * 1)), False)
        Call oWp02.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + (Math.Atan(oAngle / 2)), ((oStairsHeight - 10) / 10)), False)
        Call oWp02.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 22, ((oStairsHeight - 10) / 10)), False)
        Call oWp02.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 22, ((oStairsHeight - 10) / 10) - 18), False)
        Call oWp02.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 23 - (((Math.Tan(oAngle)) * 200) / 10) - (Math.Tan(oAngle / 2)), ((oStairsHeight - 10) / 10) - 18), False)
        Call oWp02.Add(oTG.CreatePoint2d(19, ((oStepheight - (Math.Tan(oSPAA)) * 200) / 10) + (Math.Tan(oSPAA))), False)
        Call oWp02.Add(oTG.CreatePoint2d(19, 1), False)

        Call oSideSketch.GeometricConstraints.AddGround(oWp02(1))
        Call oSideSketch.GeometricConstraints.AddGround(oWp02(2))
        Call oSideSketch.GeometricConstraints.AddGround(oWp02(3))
        Call oSideSketch.GeometricConstraints.AddGround(oWp02(4))
        Call oSideSketch.GeometricConstraints.AddGround(oWp02(5))
        Call oSideSketch.GeometricConstraints.AddGround(oWp02(7))
        Call oSideSketch.GeometricConstraints.AddGround(oWp02(8))

        Dim oLn09 As SketchLine
        oLn09 = oSideSketch2.SketchLines.AddByTwoPoints(oWp02(1), oWp02(2))
        Dim oLn10 As SketchLine
        oLn10 = oSideSketch2.SketchLines.AddByTwoPoints(oWp02(2), oWp02(3))
        Dim oLn11 As SketchLine
        oLn11 = oSideSketch2.SketchLines.AddByTwoPoints(oWp02(3), oWp02(4))
        Dim oLn12 As SketchLine
        oLn12 = oSideSketch2.SketchLines.AddByTwoPoints(oWp02(4), oWp02(5))
        Dim oLn13 As SketchLine
        oLn13 = oSideSketch2.SketchLines.AddByTwoPoints(oWp02(5), oWp02(6))
        Dim oLn14 As SketchLine
        oLn14 = oSideSketch2.SketchLines.AddByTwoPoints(oWp02(6), oWp02(7))
        Dim oLn15 As SketchLine
        oLn15 = oSideSketch2.SketchLines.AddByTwoPoints(oWp02(7), oWp02(8))
        Dim oLn16 As SketchLine
        oLn16 = oSideSketch2.SketchLines.AddByTwoPoints(oWp02(8), oWp02(1))

        Call oSideSketch.GeometricConstraints.AddParallel(oLn11, oLn13)
        Call oSideSketch.GeometricConstraints.AddParallel(oLn10, oLn14)

        'Extrusion sidebeam

        Dim oProfile2 As Profile
        oProfile2 = oSideSketch2.Profiles.AddForSolid
        Dim oExtrudeDef2 As ExtrudeDefinition
        Dim kCutOperation As Object = Nothing
        oExtrudeDef2 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile2, PartFeatureOperationEnum.kCutOperation)
        Call oExtrudeDef2.SetDistanceExtent(6.7, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        Dim oExtrude2 As ExtrudeFeature
        oExtrude2 = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef2)

        oInvApp.ScreenUpdating = True
        ProgressBar1.Value = 60
        oInvApp.ScreenUpdating = False

        'Creating the steps

        Dim oStepSketch As PlanarSketch
        oStepSketch = oDef.Sketches.Add(oCenterPlane)
        Dim StPt01 As SketchPoints
        StPt01 = oStepSketch.SketchPoints
        Call StPt01.Add(oTG.CreatePoint2d(0, (oStepheight / 10)), False)
        Call StPt01.Add(oTG.CreatePoint2d(23, (oStepheight / 10)), False)
        Call StPt01.Add(oTG.CreatePoint2d(23, (oStepheight / 10) - 3), False)
        Call StPt01.Add(oTG.CreatePoint2d(20, (oStepheight / 10) - 6), False)
        Call StPt01.Add(oTG.CreatePoint2d(0, (oStepheight / 10) - 6), False)

        Dim oStLn As SketchLine
        oStLn = oStepSketch.SketchLines.AddByTwoPoints(StPt01(1), StPt01(2))
        oStLn = oStepSketch.SketchLines.AddByTwoPoints(StPt01(2), StPt01(3))
        oStLn = oStepSketch.SketchLines.AddByTwoPoints(StPt01(3), StPt01(4))
        oStLn = oStepSketch.SketchLines.AddByTwoPoints(StPt01(4), StPt01(5))
        oStLn = oStepSketch.SketchLines.AddByTwoPoints(StPt01(5), StPt01(1))

        'Extrusion Step1

        Dim oStProf As Profile
        oStProf = oStepSketch.Profiles.AddForSolid
        Dim oStExtDef As ExtrudeDefinition
        Dim kNewBodyOperation As Object = Nothing
        oStExtDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oStProf, PartFeatureOperationEnum.kNewBodyOperation)
        Call oStExtDef.SetDistanceExtent((oStairsWidth / 10), PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
        Dim oStExt As ExtrudeFeature
        oStExt = oDef.Features.ExtrudeFeatures.Add(oStExtDef)

        Dim oStPatCol As ObjectCollection
        Dim CreateObjectCollection As ObjectCollection = Nothing

        oStPatCol = (oStExt.Definition.AffectedBodies)

        Dim oStPatAxi As WorkPlane
        oStPatAxi = oDef.WorkPlanes.AddByLinePlaneAndAngle(oDef.WorkAxes(1), oDef.WorkPlanes(2), oAngle)
        Dim oStPattern As RectangularPatternFeature
        Dim kDefault As PatternSpacingTypeEnum = Nothing
        oStPattern = oDef.Features.RectangularPatternFeatures.Add(oStPatCol, oStPatAxi, True, oNumberOfSteps, (oStairsHypotenusa / (oNumberOfSteps - 1)) / 10, kDefault)


        If oHandrailState = "Left" Then
            GoTo startLeftRail
        End If
        If oHandrailState = "None" Then
            GoTo startLeftRail
        End If

        oRailingPlane = oDef.WorkPlanes.Item(1)
        oRailingPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oCenterPlane, (oStairsWidth + 70) / 2 / 10, False)


        Dim oRailingSketch As PlanarSketch
        oRailingSketch = oDef.Sketches.Add(oRailingPlane)
        Dim oRp01 As SketchPoints
        oRp01 = oRailingSketch.SketchPoints
        Call oRp01.Add(oTG.CreatePoint2d(2, 55), False)
        Call oRp01.Add(oTG.CreatePoint2d(2, 105), False)
        Call oRp01.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 21, (Math.Tan(oAngle) * ((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 21)) + 105), False)
        Call oRp01.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 21, (Math.Tan(oAngle) * ((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 21)) + 55), False)
        Dim oRLn01 As SketchLine
        oRLn01 = oRailingSketch.SketchLines.AddByTwoPoints(oRp01(1), oRp01(2))
        Dim oRLn02 As SketchLine
        oRLn02 = oRailingSketch.SketchLines.AddByTwoPoints(oRp01(2), oRp01(3))
        Dim oRLn03 As SketchLine
        oRLn03 = oRailingSketch.SketchLines.AddByTwoPoints(oRp01(3), oRp01(4))
        Dim oRLn04 As SketchLine
        oRLn04 = oRailingSketch.SketchLines.AddByTwoPoints(oRp01(4), oRp01(1))

        Call oRailingSketch.SketchArcs.AddByFillet(oRLn01, oRLn02, 7.5, oRLn01.StartSketchPoint.Geometry, oRLn02.EndSketchPoint.Geometry)
        Call oRailingSketch.SketchArcs.AddByFillet(oRLn02, oRLn03, 7.5, oRLn02.StartSketchPoint.Geometry, oRLn03.EndSketchPoint.Geometry)
        Call oRailingSketch.SketchArcs.AddByFillet(oRLn03, oRLn04, 7.5, oRLn03.StartSketchPoint.Geometry, oRLn04.EndSketchPoint.Geometry)
        Call oRailingSketch.SketchArcs.AddByFillet(oRLn04, oRLn01, 7.5, oRLn04.StartSketchPoint.Geometry, oRLn01.EndSketchPoint.Geometry)

        Dim oPath As Path
        oPath = oDef.Features.CreatePath(oRLn01)

        'Create a work plane in the middle of the 2D sketch.

        Dim oGroundPlane As WorkPlane
        oGroundPlane = oDef.WorkPlanes.Item(3)
        Dim oRailDiaPlane As WorkPlane
        oRailDiaPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oGroundPlane, 80, False)
        Dim oRailingSketch2 As PlanarSketch
        oRailingSketch2 = oDef.Sketches.Add(oRailDiaPlane)
        Dim oRailDiaCircle As SketchCircle
        oRailDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d((oStairsWidth / 2 / 10) + 3.5, 2), 2)
        Dim oRailIntDiaCircle As SketchCircle
        oRailIntDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d((oStairsWidth / 2 / 10) + 3.5, 2), 1.8)
        Dim oRailProfile As Profile
        oRailProfile = oRailingSketch2.Profiles.AddForSolid
        Dim oSweep As SweepFeature
        oSweep = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile, oPath, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)
        '        oSweep.Name = "RightHandRailing"

        oInvApp.ScreenUpdating = True
        ProgressBar1.Value = 80
        oInvApp.ScreenUpdating = False


        'Create a SpilR

        Dim oRailingSketch3 As PlanarSketch
        oRailingSketch3 = oDef.Sketches.Add(oRailDiaPlane)
        Dim oSpilDiaCircle As SketchCircle
        oSpilDiaCircle = oRailingSketch3.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d((oStairsWidth / 2 / 10) + 3.5, 20), 1.685)
        Dim oSpilProfile As Profile
        oSpilProfile = oRailingSketch3.Profiles.AddForSolid
        Dim oExtrudeDef3 As ExtrudeDefinition
        oExtrudeDef3 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSpilProfile, PartFeatureOperationEnum.kNewBodyOperation)
        Call oExtrudeDef3.SetFromToExtent(oSweep.Faces.Item(6), True, oExtrude.Faces.Item(7), True)
        Dim oExtrude3 As ExtrudeFeature
        oExtrude3 = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef3)
        Dim RS As RenderStyle
        Try
            RS = oStairsDoc.RenderStyles.Item("StairsRailingStyle")
        Catch ex As Exception
            RS = oStairsDoc.RenderStyles.Add("StairsRailingStyle")
            RS.Reflectivity = 3
            RS.SetDiffuseColor(255, 255, 0)  'Yellow
        End Try
        'assign new render style to part feature
        'Dim oRailFeature As PartFeature = oExtrude
        Call oSweep.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
        Call oExtrude3.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)

        '-------------------------------------------------
        oInvApp.ScreenUpdating = True
        ProgressBar1.Value = 10
        oInvApp.ScreenUpdating = False
        '-------------------------------------------------

        Dim oPatCol As ObjectCollection
        Dim CreateObjectCollection2 As ObjectCollection = Nothing
        oPatCol = (oExtrude3.Definition.AffectedBodies)
        Dim oPatAxi As WorkPlane
        oPatAxi = oDef.WorkPlanes.AddByLinePlaneAndAngle(oDef.WorkAxes(1), oDef.WorkPlanes(2), oAngle)
        Dim oPattern As RectangularPatternFeature
        oPattern = oDef.Features.RectangularPatternFeatures.Add(oPatCol, oPatAxi, True, oNumberOfSpils, ((oStairsHypotenusa - 500) / (oNumberOfSpils - 1)) / 10, kDefault)
        oPatAxi.Visible = False
        oRailDiaPlane.Visible = False
        oRailingPlane.Visible = False
        '---------------------------------------------------------


startLeftRail:


        'mirror Stairsbeam

        Dim oCol As ObjectCollection
        Dim CreateObjectCollection3 As ObjectCollection = Nothing

        oCol = (oExtrude.Definition.AffectedBodies)
        oCol = (oExtrude2.Definition.AffectedBodies)

        Dim oPlane As WorkPlane
        oPlane = oStairsDoc.ComponentDefinition.WorkPlanes(1)
        Dim oMirror As MirrorFeature
        oMirror = oStairsDoc.ComponentDefinition.Features.MirrorFeatures.Add(oCol, oCenterPlane, False, PatternComputeTypeEnum.kOptimizedCompute)
        ProgressBar1.Value = 20

        '------------------------------------------------------------
        oInvApp.ScreenUpdating = True
        ProgressBar1.Value = 80
        oInvApp.ScreenUpdating = False
        '------------------------------------------------------------

        If oHandrailState = "Right" Then
            GoTo StartRightRail
        End If
        If oHandrailState = "None" Then
            GoTo StartRightRail
        End If

        'Mid Plane Railing aanmaken

        Dim oRailingPlane4 As WorkPlane
        oRailingPlane4 = oDef.WorkPlanes.Item(1)
        oRailingPlane4 = oDef.WorkPlanes.AddByPlaneAndOffset(oCenterPlane, (oStairsWidth + 70) / 2 / -10, False)

        'Railing

        Dim oRailingSketch4 As PlanarSketch
        ' Dim oRailingPlane2 As WorkPlane
        oRailingSketch4 = oDef.Sketches.Add(oRailingPlane4)
        Dim oRp02 As SketchPoints
        oRp02 = oRailingSketch4.SketchPoints
        Call oRp02.Add(oTG.CreatePoint2d(2, 55), False)
        Call oRp02.Add(oTG.CreatePoint2d(2, 105), False)
        Call oRp02.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 21, (Math.Tan(oAngle) * ((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 21)) + 105), False)
        Call oRp02.Add(oTG.CreatePoint2d((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 21, (Math.Tan(oAngle) * ((((oStairsHeight / 10) - (oStepheight / 10)) / Math.Tan(oAngle)) + 21)) + 55), False)
        Dim oRLn05 As SketchLine
        oRLn05 = oRailingSketch4.SketchLines.AddByTwoPoints(oRp02(1), oRp02(2))
        Dim oRLn06 As SketchLine
        oRLn06 = oRailingSketch4.SketchLines.AddByTwoPoints(oRp02(2), oRp02(3))
        Dim oRLn07 As SketchLine
        oRLn07 = oRailingSketch4.SketchLines.AddByTwoPoints(oRp02(3), oRp02(4))
        Dim oRLn08 As SketchLine
        oRLn08 = oRailingSketch4.SketchLines.AddByTwoPoints(oRp02(4), oRp02(1))

        Call oRailingSketch4.SketchArcs.AddByFillet(oRLn05, oRLn06, 7.5, oRLn05.StartSketchPoint.Geometry, oRLn06.EndSketchPoint.Geometry)
        Call oRailingSketch4.SketchArcs.AddByFillet(oRLn06, oRLn07, 7.5, oRLn06.StartSketchPoint.Geometry, oRLn07.EndSketchPoint.Geometry)
        Call oRailingSketch4.SketchArcs.AddByFillet(oRLn07, oRLn08, 7.5, oRLn07.StartSketchPoint.Geometry, oRLn08.EndSketchPoint.Geometry)
        Call oRailingSketch4.SketchArcs.AddByFillet(oRLn08, oRLn05, 7.5, oRLn08.StartSketchPoint.Geometry, oRLn05.EndSketchPoint.Geometry)

        '-------------------------------------------------
        ' ProgressBar1.Value = 60
        '-------------------------------------------------

        Dim oPath2 As Path
        oPath2 = oDef.Features.CreatePath(oRLn05)

        'Create a work plane in the middle of the 2D sketch.

        Dim oGroundPlane2 As WorkPlane
        oGroundPlane2 = oDef.WorkPlanes.Item(3)
        Dim oRailDiaPlane2 As WorkPlane
        oRailDiaPlane2 = oDef.WorkPlanes.AddByPlaneAndOffset(oGroundPlane2, 80, False)
        Dim oRailingSketch5 As PlanarSketch
        oRailingSketch5 = oDef.Sketches.Add(oRailDiaPlane2)
        Dim oRailDiaCircle2 As SketchCircle
        oRailDiaCircle2 = oRailingSketch5.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d((oStairsWidth / -2 / 10) - 3.5, 2), 2)
        Dim oRailIntDiaCircle2 As SketchCircle
        oRailIntDiaCircle2 = oRailingSketch5.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d((oStairsWidth / -2 / 10) - 3.5, 2), 1.8)

        Dim oRailProfile2 As Profile
        oRailProfile2 = oRailingSketch5.Profiles.AddForSolid

        Dim oSweep2 As SweepFeature
        oSweep2 = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile2, oPath2, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)

        '------------------------------------------------------------
        oInvApp.ScreenUpdating = True
        ProgressBar1.Value = 10
        oInvApp.ScreenUpdating = False
        '------------------------------------------------------------

        Dim oRailingSketch6 As PlanarSketch
        oRailingSketch6 = oDef.Sketches.Add(oRailDiaPlane2)
        Dim oSpilDiaCircle2 As SketchCircle
        oSpilDiaCircle2 = oRailingSketch6.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d((oStairsWidth / -2 / 10) - 3.5, 20), 1.685)
        Dim oSpilProfile2 As Profile
        oSpilProfile2 = oRailingSketch6.Profiles.AddForSolid
        Dim oExtrudeDef4 As ExtrudeDefinition
        oExtrudeDef4 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSpilProfile2, PartFeatureOperationEnum.kNewBodyOperation)
        Call oExtrudeDef4.SetFromToExtent(oSweep2.Faces.Item(6), True, oMirror.Faces.Item(7), True)
        Dim oExtrude4 As ExtrudeFeature
        oExtrude4 = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef4)


        Try
            RS = oStairsDoc.RenderStyles.Item("StairsRailingStyle")
        Catch ex As Exception
            RS = oStairsDoc.RenderStyles.Add("StairsRailingStyle")
            RS.Reflectivity = 3
            RS.SetDiffuseColor(255, 255, 0)  'Yellow
        End Try

        Call oSweep2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
        Call oExtrude4.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
        Dim oPatCol2 As ObjectCollection
        Dim CreateObjectCollection4 As ObjectCollection = Nothing
        oPatCol2 = (oExtrude4.Definition.AffectedBodies)
        Dim oPatAxi2 As WorkPlane
        oPatAxi2 = oDef.WorkPlanes.AddByLinePlaneAndAngle(oDef.WorkAxes(1), oDef.WorkPlanes(2), oAngle)
        Dim oPattern2 As RectangularPatternFeature
        oPattern2 = oDef.Features.RectangularPatternFeatures.Add(oPatCol2, oPatAxi2, True, oNumberOfSpils, ((oStairsHypotenusa - 500) / (oNumberOfSpils - 1)) / 10, kDefault)
        oPatAxi2.Visible = False
        oRailDiaPlane2.Visible = False
        oRailingPlane4.Visible = False
        oRailingSketch5.Visible = False

        '------------------------------------------------------------

StartRightRail:

        oInvApp.ScreenUpdating = True
        ProgressBar1.Value = 90
        oInvApp.ScreenUpdating = False

        oCenterPlane.Visible = False
        oRightSidePlane.Visible = False
        oStPatAxi.Visible = False
        oStairsDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits


        '-------------------------------------------------
        oInvApp.ScreenUpdating = True
        ProgressBar1.Value = 100
        oInvApp.ScreenUpdating = False


        '---------Nieuw part saven in juiste directory

        Try
            oStairsDoc.SaveAs(lblPath.Text & oFilename, False)
        Catch ex As Exception
            MsgBox("Could not save the part." & vbCrLf & "Try again or check your credentials", vbOKOnly + "4064", "Warning")
            oInvApp.ScreenUpdating = True
            GoTo EndRoutine
        End Try

        oInvApp.ScreenUpdating = True
        ProgressBar1.Value = 100
        oInvApp.ScreenUpdating = False

        oStairsDoc.Close()
        oInvApp.ScreenUpdating = True

        '--------------Place the part interactive--------------------------------------
PlacePart:
        ProgressBar1.Value = 0
        ' Get the active assembly.

        asmDoc = oInvApp.ActiveDocument

        ' If assembly empty place stairs grounded at origin.

        If asmDoc.ComponentDefinition.ImmediateReferencedDefinitions.Count < 1 Then
            Dim trans As Matrix
            trans = oInvApp.TransientGeometry.CreateMatrix
            stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblPath.Text & oFilename), trans)
            stairsOcc.Grounded = True
            GoTo EndRoutine
        End If

        'Else place stairs on a floor surface.

        Hide()
        oFloorFace = oInvApp.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, "Select a floor to place the stairs.")
        If oFloorFace Is Nothing Then
            GoTo EndRoutine
        End If

        stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblPath.Text & oFilename), oInvApp.TransientGeometry.CreateMatrix)
        oStairsDoc = stairsOcc.Definition.Document

        Try
            attrSets = oStairsDoc.AttributeManager.FindAttributeSets("ConnectFace")
            Dim oStairsCon As Face
            oStairsCon = attrSets.Item(1).Parent.Parent
            Dim oStairsFaceProxy As FaceProxy
            oStairsFaceProxy = Nothing
            Call stairsOcc.CreateGeometryProxy(oStairsCon, oStairsFaceProxy)
            Call asmDoc.ComponentDefinition.Constraints.AddMateConstraint(oFloorFace, oStairsFaceProxy, 0)
        Catch ex As Exception
            'Dim trans2 As Matrix
            'trans2 = oInvApp.TransientGeometry.CreateMatrix
            'stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblPath.Text & oFilename), trans2)
            'GoTo EndRoutine
        End Try

EndRoutine:
        Me.Show()
    End Sub

    '------------------------------------------------
    '
    '          PLACING Curved STAIRS IN AN ASSEMBLY
    '
    '------------------------------------------------

    Private Sub btnPlaceCurvedStairs_Click(sender As Object, e As EventArgs) Handles btnPlaceCurvedStairs.Click
        PlaceCurvedStairs()
    End Sub

    Private Sub PlaceCurvedStairs()

        Dim oDef As PartComponentDefinition
        Dim oTG As TransientGeometry
        Dim asmDoc As AssemblyDocument
        Dim oFloorFace As Face
        Dim oStairsDoc As PartDocument
        Dim attrSets As AttributeSetsEnumerator
        Dim stairsOcc As ComponentOccurrence
        Dim propSet1 As PropertySet
        Dim propSet2 As PropertySet
        Dim propSet3 As PropertySet

        'Check of Inventor nog steeds actief is
        Try
            Me.oInvApp = GetObject(, "Inventor.Application")
        Catch ex As Exception
            MsgBox("Inventor must be running!.", vbExclamation, "")
            btnPlaceStairsInAssem.Enabled = False
            btnPlaceLadder.Enabled = False
            btnPlaceRailing.Enabled = False
            btnPlacePlatform.Enabled = False
            btnPlaceCurvedStairs.Enabled = False
            Exit Sub
        End Try

        'Check of er een assembly open is
        If Me.oInvApp.ActiveDocumentType = Global.Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            GoTo BeginRoutine
        Else MsgBox("Open, Create or Activate an Assembly !", vbExclamation, "")
            GoTo EndRoutine
        End If

BeginRoutine:
        If Dir(lblCurvedStairsPath.Text & oCSFilename) = "" Then
            GoTo StartPart
        Else
            GoTo PlacePart
        End If
StartPart:

        oInvApp = GetObject(, "Inventor.Application")

        Dim oCircumIntStairs As Double
        Dim oCircumMidStairs As Double
        Dim oCircumOutStairs As Double
        Dim oProjetSpiralMidStairs As Double
        Dim oLenghtSpiralMidStairs As Double
        Dim oLenghtSpiralOutStairs As Double
        Dim oLenghtSpiralIntStairs As Double
        Dim oStairRevAngle As Double
        Dim oStairRev As Double
        ProgressBar5.Value = 5
        oInvApp.ScreenUpdating = False



        oStairsDoc = oInvApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject)
        oStairsDoc = oInvApp.ActiveDocument
        oDef = oInvApp.ActiveDocument.ComponentDefinition
        oTG = oInvApp.TransientGeometry
        pi = Math.Acos(-1) 'radians van 45° hoek
        oAngle = (oCSAngle * pi / 180)
        oNumberOfSteps = Math.Ceiling(oCSHeight / 200)
        oStepheight = oCSHeight / oNumberOfSteps
        oCircumIntStairs = 2 * oCSIntRadius * pi
        oCircumMidStairs = 2 * (oCSIntRadius + 8 + (oCSWidth / 2)) * pi
        oCircumOutStairs = 2 * (oCSIntRadius + 16 + oCSWidth) * pi
        oProjetSpiralMidStairs = (oCSHeight - oStepheight) / (Math.Tan(oAngle))
        oLenghtSpiralMidStairs = Math.Sqrt(((oCSHeight - oStepheight) * (oCSHeight - oStepheight)) + (oProjetSpiralMidStairs * oProjetSpiralMidStairs))
        oStairRevAngle = (360 / (oCircumMidStairs / oProjetSpiralMidStairs)) * pi / 180 'in rad
        oStairRev = (360 / (oCircumMidStairs / oProjetSpiralMidStairs)) / 360
        Dim oRatio As Double
        oRatio = (360 / (oCircumMidStairs / oProjetSpiralMidStairs)) / 360
        oLenghtSpiralOutStairs = Math.Sqrt(((oCircumOutStairs * oRatio) * (oCircumOutStairs * oRatio)) + ((oCSHeight - oStepheight) * (oCSHeight - oStepheight)))
        oLenghtSpiralIntStairs = Math.Sqrt(((oCircumIntStairs * oRatio) * (oCircumIntStairs * oRatio)) + ((oCSHeight - oStepheight) * (oCSHeight - oStepheight)))

        propSet1 = oStairsDoc.PropertySets.Item("Inventor Summary Information")
        propSet1.ItemByPropId(2).Value = "Steel Curved Stairs"
        propSet1.ItemByPropId(4).Value = "Marc Crauwels"
        propSet1.ItemByPropId(6).Value = "This Steel Staircase is generated with the Pantarein Water Panta Stairs application. This software is part of the Pantarein Water Design Suite."

        propSet2 = oStairsDoc.PropertySets.Item("Inventor Document Summary Information")
        propSet2.ItemByPropId(2).Value = "Plant Design Mechanical"
        propSet2.ItemByPropId(15).Value = "Pantarein Water BVBA"

        propSet3 = oStairsDoc.PropertySets.Item("Design Tracking Properties")
        propSet3.ItemByPropId(41).Value = "Marc Crauwels"
        propSet3.ItemByPropId(29).Value = "Steel Curved Stairs"
        propSet3.ItemByPropId(23).Value = "http://www.pantareinwater.be/en"

        propSet1 = Nothing
        propSet2 = Nothing
        propSet3 = Nothing
        '-------------------------------------------------------------------------------------------------------------------------------------------------

        '----------Adjust Orientation XY is floor Z is up----------

        '-------------------------------------------------------------------------------------------------------------------------------------------------
        Dim v As Inventor.View
        v = oInvApp.ActiveView
        Dim c As Camera
        c = oInvApp.ActiveView.Camera
        Dim t2eDist As Double
        t2eDist = c.Target.DistanceTo(c.Eye)
        Dim t2e As Vector
        t2e = oTG.CreateVector(0, -t2eDist, 0)
        Dim newEye As Point
        newEye = c.Target.Copy
        Call newEye.TranslateBy(t2e)
        c.Eye = newEye
        c.UpVector = oTG.CreateUnitVector(0, 0, 1)
        c.ApplyWithoutTransition()
        v.SetCurrentAsFront()

        If oCSState = "CCW_None" Or oCSState = "CCW_Right" Or oCSState = "CCW_Left" Or oCSState = "CCW_Both" Then

            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 10
            oInvApp.ScreenUpdating = False



            '----------------Workplane Tankside ----------------------
            Dim oYplane As WorkPlane
            Dim oIntSidePlane As WorkPlane
            oYplane = oDef.WorkPlanes.Item(1)
            oIntSidePlane = oDef.WorkPlanes.AddByPlaneAndOffset(oYplane, oCSIntRadius / 10, False)

            oIntSidePlane.Visible = False
            '----------------Sketch Tankside ----------------------
            Dim oTankSideSketch As PlanarSketch
            oTankSideSketch = oDef.Sketches.Add(oIntSidePlane)
            Dim oWp01 As SketchPoints
            oWp01 = oTankSideSketch.SketchPoints

            '------------Sketch points side----------------------
            Call oWp01.Add(oTG.CreatePoint2d(0, 0), False)
            Call oWp01.Add(oTG.CreatePoint2d(0, (oStepheight / 10)), False)
            Call oWp01.Add(oTG.CreatePoint2d(Math.Cos(((90 - oCSAngle) * pi / 180) / 2) * 26, ((oStepheight / 10) - (Math.Sin(((90 - oCSAngle) * pi / 180) / 2) * 26))), False)
            Call oWp01.Add(oTG.CreatePoint2d(Math.Cos(((90 - oCSAngle) * pi / 180) / 2) * 26, 0), False)
            '-------------Draw lines----------------------------------------------------
            Dim oLn01 As SketchLine
            oLn01 = oTankSideSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
            Dim oLn02 As SketchLine
            oLn02 = oTankSideSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
            Dim oLn03 As SketchLine
            oLn03 = oTankSideSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
            Dim oLn04 As SketchLine
            oLn04 = oTankSideSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(1))
            '--------------Extrusion of the TankSideGroundplate--------------------------------------
            Dim oGPlateProf As Profile
            oGPlateProf = oTankSideSketch.Profiles.AddForSolid
            Dim oGPlateExtrudeDef As ExtrudeDefinition
            Dim kJoinOperation As Object = Nothing
            oGPlateExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oGPlateProf, PartFeatureOperationEnum.kJoinOperation)
            Call oGPlateExtrudeDef.SetDistanceExtent(0.8, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oGPlateExtrude As ExtrudeFeature
            oGPlateExtrude = oDef.Features.ExtrudeFeatures.Add(oGPlateExtrudeDef)
            'oGPlateExtrude.Name = "Tank Side Ground Plate"

            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 15
            oInvApp.ScreenUpdating = False

            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(3)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------



            '--------------Pattern of the TankSideGroundplate--------------------------------------
            Dim oGPlateCol As ObjectCollection
            oGPlateCol = (oGPlateExtrude.Definition.AffectedBodies)
            Dim oGPlatePattern As RectangularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing
            oGPlatePattern = oDef.Features.RectangularPatternFeatures.Add(oGPlateCol, oDef.WorkPlanes.Item(1), True, 2, oCSWidth / 10 + 0.8, kDefault)
            Dim oSpiralSketch As PlanarSketch
            oSpiralSketch = oDef.Sketches.Add(oGPlateExtrude.SurfaceBody.Faces.Item(1))
            Dim oLn05 As SketchEntity
            oLn05 = oSpiralSketch.AddByProjectingEntity(oGPlateExtrude.SurfaceBody.Faces.Item(1).Edges.Item(1))
            Dim oLn06 As SketchEntity
            oLn06 = oSpiralSketch.AddByProjectingEntity(oGPlateExtrude.SurfaceBody.Faces.Item(1).Edges.Item(2))
            Dim oLn07 As SketchEntity
            oLn07 = oSpiralSketch.AddByProjectingEntity(oGPlateExtrude.SurfaceBody.Faces.Item(1).Edges.Item(3))
            Dim oLn08 As SketchEntity
            oLn08 = oSpiralSketch.AddByProjectingEntity(oGPlateExtrude.SurfaceBody.Faces.Item(1).Edges.Item(4))
            Dim oLn09 As SketchEntity
            oLn09 = oSpiralSketch.AddByProjectingEntity(oGPlatePattern.PatternElements.Item(2).Faces.Item(1).Edges.Item(1))
            Dim oLn10 As SketchEntity
            oLn10 = oSpiralSketch.AddByProjectingEntity(oGPlatePattern.PatternElements.Item(2).Faces.Item(1).Edges.Item(2))
            Dim oLn11 As SketchEntity
            oLn11 = oSpiralSketch.AddByProjectingEntity(oGPlatePattern.PatternElements.Item(2).Faces.Item(1).Edges.Item(3))
            Dim oLn12 As SketchEntity
            oLn12 = oSpiralSketch.AddByProjectingEntity(oGPlatePattern.PatternElements.Item(2).Faces.Item(1).Edges.Item(4))
            Dim oIntSpiralProf As Profile
            oIntSpiralProf = oSpiralSketch.Profiles.AddForSolid
            Dim oSpiralIntSide As CoilFeature
            oSpiralIntSide = oDef.Features.CoilFeatures.AddByRevolutionAndHeight(oIntSpiralProf, oDef.WorkAxes.Item(3), oStairRev, (oCSHeight - oStepheight) / 10, PartFeatureOperationEnum.kJoinOperation, False,)

            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 20
            oInvApp.ScreenUpdating = False

            '--------------Cut of the Spirals--------------------------------------

            Dim oTopStepFrontPlane As WorkPlane
            Dim oCenterline As WorkAxis
            oCenterline = oDef.WorkAxes(3)
            oTopStepFrontPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oCenterline, oDef.WorkPlanes(2), oStairRevAngle, False)
            'oTopStepFrontPlane.Name = "Top Step Front Plane"
            oTopStepFrontPlane.Visible = False
            Dim oCutSketch As PlanarSketch
            oCutSketch = oDef.Sketches.Add(oTopStepFrontPlane)
            Dim oWp04 As SketchPoints
            oWp04 = oCutSketch.SketchPoints
            Call oWp04.Add(oTG.CreatePoint2d(oCSHeight / 10, oCSIntRadius / 10 - 5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oCSHeight / 10, oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.8 + 5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.8 + 5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, oCSIntRadius / 10 - 5), False)
            Dim oLn25 As SketchLine
            oLn25 = oCutSketch.SketchLines.AddByTwoPoints(oWp04(1), oWp04(2))
            Dim oLn26 As SketchLine
            oLn26 = oCutSketch.SketchLines.AddByTwoPoints(oWp04(2), oWp04(3))
            Dim oLn27 As SketchLine
            oLn27 = oCutSketch.SketchLines.AddByTwoPoints(oWp04(3), oWp04(4))
            Dim oLn28 As SketchLine
            oLn28 = oCutSketch.SketchLines.AddByTwoPoints(oWp04(4), oWp04(1))
            ProgressBar5.Value = 40
            '--------------Cut of Stairs End--------------------------------------

            Dim oCutProf As Profile
            oCutProf = oCutSketch.Profiles.AddForSolid
            Dim oCutExtrudeDef As ExtrudeDefinition
            oCutExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oCutProf, PartFeatureOperationEnum.kCutOperation)
            Call oCutExtrudeDef.SetDistanceExtent(26, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oCutExtrude As ExtrudeFeature
            oCutExtrude = oDef.Features.ExtrudeFeatures.Add(oCutExtrudeDef)

            '------------Sketch points Stairs End----------------------

            Dim oStairEndSketch As PlanarSketch
            oStairEndSketch = oDef.Sketches.Add(oTopStepFrontPlane)
            Dim oWp03 As SketchPoints
            oWp03 = oStairEndSketch.SketchPoints
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10, oCSIntRadius / 10), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 4, oCSIntRadius / 10 + 0.8), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, oCSIntRadius / 10 + 0.8), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, oCSIntRadius / 10), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 4, oCSIntRadius / 10 + oCSWidth / 10 + 0.8), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10, oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.8), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.8), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, oCSIntRadius / 10 + oCSWidth / 10 + 0.8), False)

            '-------------Draw lines----------------------------------------------------

            Dim oLn17 As SketchLine
            oLn17 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(1), oWp03(6))
            Dim oLn18 As SketchLine
            oLn18 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(6), oWp03(7))
            Dim oLn19 As SketchLine
            oLn19 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(7), oWp03(8))
            Dim oLn20 As SketchLine
            oLn20 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(8), oWp03(5))
            Dim oLn21 As SketchLine
            oLn21 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(5), oWp03(2))
            Dim oLn22 As SketchLine
            oLn22 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(2), oWp03(3))
            Dim oLn23 As SketchLine
            oLn23 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(3), oWp03(4))
            Dim oLn24 As SketchLine
            oLn24 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(4), oWp03(1))
            ProgressBar5.Value = 50


            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 25
            oInvApp.ScreenUpdating = False

            '--------------Extrusion of Stairs End--------------------------------------

            Dim oEndProf As Profile
            oEndProf = oStairEndSketch.Profiles.AddForSolid
            Dim oENdExtrudeDef As ExtrudeDefinition
            oENdExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oEndProf, PartFeatureOperationEnum.kNewBodyOperation)
            Call oENdExtrudeDef.SetDistanceExtent(26, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oEndExtrude As ExtrudeFeature
            oEndExtrude = oDef.Features.ExtrudeFeatures.Add(oENdExtrudeDef)

            '----------------Sketch Step 1 ---------------------

            Dim oStep1Plane As WorkPlane
            oStep1Plane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(3), oStepheight / 10, False)
            Dim oStep1Sketch As PlanarSketch
            oStep1Sketch = oDef.Sketches.Add(oStep1Plane)
            Dim oWp02 As SketchPoints
            oWp02 = oStep1Sketch.SketchPoints
            Call oWp02.Add(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8, 0), False)
            Call oWp02.Add(oTG.CreatePoint2d(oCSIntRadius / 10, 26), False)
            Call oWp02.Add(oTG.CreatePoint2d(oCSIntRadius / 10 + oCSWidth / 10, 26), False)
            Call oWp02.Add(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8 + oCSWidth / 10, 0), False)
            Dim oLn13 As SketchLine
            oLn13 = oStep1Sketch.SketchLines.AddByTwoPoints(oWp02(1), oWp02(2))
            Dim oLn14 As SketchLine
            oLn14 = oStep1Sketch.SketchLines.AddByTwoPoints(oWp02(2), oWp02(3))
            Dim oLn15 As SketchLine
            oLn15 = oStep1Sketch.SketchLines.AddByTwoPoints(oWp02(3), oWp02(4))
            Dim oLn16 As SketchLine
            oLn16 = oStep1Sketch.SketchLines.AddByTwoPoints(oWp02(4), oWp02(1))

            '--------------Extrusion of Step 1--------------------------------------

            Dim oStep1Prof As Profile
            oStep1Prof = oStep1Sketch.Profiles.AddForSolid
            Dim oStep1ExtrudeDef As ExtrudeDefinition
            oStep1ExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oStep1Prof, PartFeatureOperationEnum.kNewBodyOperation)
            Call oStep1ExtrudeDef.SetDistanceExtent(4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oStep1Extrude As ExtrudeFeature
            oStep1Extrude = oDef.Features.ExtrudeFeatures.Add(oStep1ExtrudeDef)
            'oStep1Extrude.Name = "First Step"
            oStep1Plane.Visible = False

            '--------------Pattern of the Steps--------------------------------------

            Dim oStepsCol As ObjectCollection
            oStepsCol = (oStep1Extrude.Definition.AffectedBodies)
            Dim oStepsPattern As RectangularPatternFeature
            oStepsPattern = oDef.Features.RectangularPatternFeatures.Add(oStepsCol, oSpiralIntSide.SurfaceBody.Edges.Item(40), True, oNumberOfSteps - 1, oLenghtSpiralOutStairs / 10 / (oNumberOfSteps - 1), kDefault,,,,,,,,, PatternOrientationEnum.kAdjustToDirection1)


            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 35
            oInvApp.ScreenUpdating = False


            If oCSState = "CCW_None" Then GoTo CCWend
            If oCSState = "CCW_Left" Then GoTo CCWLeftStart


            '----------------Sketch First Outside Spil ----------------------
            '----------------Spil on 15cm from stairs front------------------

            Dim oSpil1Plane As WorkPlane
            oSpil1Plane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(3), oStepheight / 10, False)
            Dim oSpil1Sketch As PlanarSketch
            oSpil1Sketch = oDef.Sketches.Add(oSpil1Plane)
            Dim oCirc01 As SketchCircle
            oCirc01 = oSpil1Sketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.4, 15), 3.37 / 2)

            '--------------Extrusion of First Outside Spil--------------------------------------

            Dim oSpil1Prof As Profile
            oSpil1Prof = oSpil1Sketch.Profiles.AddForSolid
            Dim oSpil1ExtrudeDef As ExtrudeDefinition
            oSpil1ExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSpil1Prof, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSpil1ExtrudeDef.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oSpil1Extrude As ExtrudeFeature
            oSpil1Extrude = oDef.Features.ExtrudeFeatures.Add(oSpil1ExtrudeDef)
            'oSpil1Extrude.Name = "First Outside Spil"
            oSpil1Plane.Visible = False

            '--------------Upper outer railing--------------------------------------

            Dim oPnt1 As Point
            oPnt1 = oTG.CreatePoint(0, 15, oStepheight / 10 + 105)
            Dim oWorkPoint1 As WorkPoint
            oWorkPoint1 = oDef.WorkPoints.AddFixed(oPnt1, False)
            Dim oPnt2 As Point
            oPnt2 = oTG.CreatePoint(oCSIntRadius / 10, 15, oStepheight / 10 + 105)
            Dim oWorkPoint2 As WorkPoint
            oWorkPoint2 = oDef.WorkPoints.AddFixed(oPnt2, False)
            Dim oWorkAxis1 As WorkAxis
            oWorkAxis1 = oDef.WorkAxes.AddByTwoPoints(oWorkPoint1, oWorkPoint2, False)
            Dim oUpprRailWorkPlane As WorkPlane
            oUpprRailWorkPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oWorkAxis1, oDef.WorkPlanes(3), (180 - oStairsAngle) * pi / 180)
            Dim oUpprRailSketch As PlanarSketch
            oUpprRailSketch = oDef.Sketches.Add(oUpprRailWorkPlane)
            Dim oCirc02 As SketchCircle
            oCirc02 = oUpprRailSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.4, 0), 4.21 / 2)
            Dim oUpOutRailProf As Profile
            oUpOutRailProf = oUpprRailSketch.Profiles.AddForSolid
            Dim oSpiralUpOutRail As CoilFeature
            oSpiralUpOutRail = oDef.Features.CoilFeatures.AddByRevolutionAndHeight(oUpOutRailProf, oDef.WorkAxes.Item(3), oStairRev, (oCSHeight - oStepheight) / 10, PartFeatureOperationEnum.kJoinOperation, False,)
            'oSpiralUpOutRail.Name = "Upper Outside Railing"
            oWorkPoint1.Visible = False
            oWorkPoint2.Visible = False
            oWorkAxis1.Visible = False
            oUpprRailWorkPlane.Visible = False

            '--------------Lower outer railing--------------------------------------

            Dim oPnt3 As Point
            oPnt3 = oTG.CreatePoint(0, 15, oStepheight / 10 + 55)
            Dim oWorkPoint3 As WorkPoint
            oWorkPoint3 = oDef.WorkPoints.AddFixed(oPnt3, False)
            Dim oPnt4 As Point
            oPnt4 = oTG.CreatePoint(oCSIntRadius / 10, 15, oStepheight / 10 + 55)
            Dim oWorkPoint4 As WorkPoint
            oWorkPoint4 = oDef.WorkPoints.AddFixed(oPnt4, False)
            Dim oWorkAxis2 As WorkAxis
            oWorkAxis2 = oDef.WorkAxes.AddByTwoPoints(oWorkPoint3, oWorkPoint4, False)
            Dim oLowprRailWorkPlane As WorkPlane
            oLowprRailWorkPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oWorkAxis2, oDef.WorkPlanes(3), (180 - oStairsAngle) * pi / 180)
            Dim oLowerRailSketch As PlanarSketch
            oLowerRailSketch = oDef.Sketches.Add(oLowprRailWorkPlane)
            Dim oCirc03 As SketchCircle
            oCirc03 = oLowerRailSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.4, 0), 4.21 / 2)
            Dim oLowOutRailProf As Profile
            oLowOutRailProf = oLowerRailSketch.Profiles.AddForSolid
            Dim oSpiralLowOutRail As CoilFeature
            oSpiralLowOutRail = oDef.Features.CoilFeatures.AddByRevolutionAndHeight(oLowOutRailProf, oDef.WorkAxes.Item(3), oStairRev, (oCSHeight - oStepheight) / 10, PartFeatureOperationEnum.kJoinOperation, False,)
            'oSpiralLowOutRail.Name = "Lower Outside Railing"
            oWorkPoint3.Visible = False
            oWorkPoint4.Visible = False
            oWorkAxis2.Visible = False
            oLowprRailWorkPlane.Visible = False
            Dim oOuterSpilCol As ObjectCollection
            oOuterSpilCol = (oSpil1Extrude.Definition.AffectedBodies)
            oNumberOfSpils = oNumberOfSteps / 4
            If oNumberOfSpils <= 1 Then
                oNumberOfSpils = 2
            End If
            Dim oOuterSpilPattern As RectangularPatternFeature
            oOuterSpilPattern = oDef.Features.RectangularPatternFeatures.Add(oOuterSpilCol, oSpiralIntSide.SurfaceBody.Edges.Item(40), True, oNumberOfSpils, oLenghtSpiralOutStairs / 10 / (oNumberOfSpils - 1), kDefault,,,,,,,,, PatternOrientationEnum.kAdjustToDirection1)

            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 50
            oInvApp.ScreenUpdating = False


            '----------------Sketch Outside Lower Railing Elbow ---------------------

            Dim oLowerOutsideElbowPlane As WorkPlane
            oLowerOutsideElbowPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(1), oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.4, False)
            Dim oLowerOutsideElbowSketch As PlanarSketch
            oLowerOutsideElbowSketch = oDef.Sketches.Add(oLowerOutsideElbowPlane)
            Dim oWp05 As SketchPoints
            oWp05 = oLowerOutsideElbowSketch.SketchPoints
            Call oWp05.Add(oTG.CreatePoint2d(15, 105 + oStepheight / 10), False)
            Call oWp05.Add(oTG.CreatePoint2d(0, 105 + oStepheight / 10 - (Math.Tan(oAngle * ((oCSIntRadius + 0.4) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15)), False)
            Call oWp05.Add(oTG.CreatePoint2d(0, 55 + oStepheight / 10 - (Math.Tan(oAngle * ((oCSIntRadius + 0.4) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15)), False)
            Call oWp05.Add(oTG.CreatePoint2d(15, 55 + oStepheight / 10), False)
            Dim oLn30 As SketchLine
            oLn30 = oLowerOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp05(1), oWp05(2))
            Dim oLn31 As SketchLine
            oLn31 = oLowerOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp05(2), oWp05(3))
            Dim oLn32 As SketchLine
            oLn32 = oLowerOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp05(3), oWp05(4))
            Dim oArc4 As SketchArc
            oArc4 = oLowerOutsideElbowSketch.SketchArcs.AddByFillet(oLn30, oLn31, 7.5, oLn30.Geometry.MidPoint, oLn31.Geometry.MidPoint)
            Dim oArc5 As SketchArc
            oArc5 = oLowerOutsideElbowSketch.SketchArcs.AddByFillet(oLn31, oLn32, 7.5, oLn31.Geometry.MidPoint, oLn32.Geometry.MidPoint)
            Dim oPath1 As Path
            oPath1 = oDef.Features.CreatePath(oLn32)
            Dim oRailProfile1 As Profile
            oRailProfile1 = oUpprRailSketch.Profiles.AddForSolid
            Dim oSweep As SweepFeature
            oSweep = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile1, oPath1, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)
            'oSweep.Name = "Outer Lower Railing Elbow"
            oLowerOutsideElbowPlane.Visible = False
            ProgressBar5.Value = 70
            '----------------Upper Outside Railing Elbow ----------------------

            Dim oUpperOutsideElbowPlane As WorkPlane
            oUpperOutsideElbowPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oLn20, oTopStepFrontPlane, 270 * pi / 180, False)
            Dim oUpperOutsideElbowSketch As PlanarSketch
            oUpperOutsideElbowSketch = oDef.Sketches.Add(oUpperOutsideElbowPlane)
            Dim oWp06 As SketchPoints
            oWp06 = oUpperOutsideElbowSketch.SketchPoints
            Call oWp06.Add(oTG.CreatePoint2d(oStepheight / 10 + 105, -15), False)
            Call oWp06.Add(oTG.CreatePoint2d(oStepheight / 10 + 105 + (Math.Tan(oAngle * ((oCSIntRadius + 0.4) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15), -30), False)
            Call oWp06.Add(oTG.CreatePoint2d(oStepheight / 10 + 55 + (Math.Tan(oAngle * ((oCSIntRadius + 0.4) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15), -30), False)
            Call oWp06.Add(oTG.CreatePoint2d(oStepheight / 10 + 55, -15), False)
            Dim oLn33 As SketchLine
            oLn33 = oUpperOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp06(1), oWp06(2))
            Dim oLn34 As SketchLine
            oLn34 = oUpperOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp06(2), oWp06(3))
            Dim oLn35 As SketchLine
            oLn35 = oUpperOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp06(3), oWp06(4))
            Dim oArc6 As SketchArc
            oArc6 = oUpperOutsideElbowSketch.SketchArcs.AddByFillet(oLn33, oLn34, 7.5, oLn33.Geometry.MidPoint, oLn34.Geometry.MidPoint)
            Dim oArc7 As SketchArc
            oArc7 = oUpperOutsideElbowSketch.SketchArcs.AddByFillet(oLn34, oLn35, 7.5, oLn34.Geometry.MidPoint, oLn35.Geometry.MidPoint)
            Dim oHelpPln1 As WorkPlane
            oHelpPln1 = oDef.WorkPlanes.AddByPlaneAndOffset(oTopStepFrontPlane, 15, False)
            Dim oHelpSketch As PlanarSketch
            oHelpSketch = oDef.Sketches.Add(oHelpPln1)
            Dim oHelpPnt1 As SketchPoints
            oHelpPnt1 = oHelpSketch.SketchPoints
            Call oHelpPnt1.Add(oTG.CreatePoint2d(oCSHeight / 10 + 105, 0), False)
            Call oHelpPnt1.Add(oTG.CreatePoint2d(oCSHeight / 10 + 105, oCSIntRadius / 10), False)
            Dim oHelpLn1 As SketchLine
            oHelpLn1 = oHelpSketch.SketchLines.AddByTwoPoints(oHelpPnt1(1), oHelpPnt1(2))
            Dim oHelpPln2 As WorkPlane
            oHelpPln2 = oDef.WorkPlanes.AddByLinePlaneAndAngle(oHelpLn1, oHelpPln1, oAngle, False)
            Dim oUpOutRailElbSketch As PlanarSketch
            oUpOutRailElbSketch = oDef.Sketches.Add(oHelpPln2)
            Dim oCirc10 As SketchCircle
            oCirc10 = oUpOutRailElbSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.4, 0), 4.21 / 2)
            Dim oPath2 As Path
            oPath2 = oDef.Features.CreatePath(oLn34)
            Dim oRailProfile2 As Profile
            oRailProfile2 = oUpOutRailElbSketch.Profiles.AddForSolid
            Dim oSweep2 As SweepFeature
            oSweep2 = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile2, oPath2, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)

            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 70
            oInvApp.ScreenUpdating = False



            oUpperOutsideElbowPlane.Visible = False
            oHelpPln1.Visible = False
            oHelpPln2.Visible = False
            oHelpSketch.Visible = False

            Dim RS As RenderStyle
            Try
                RS = oStairsDoc.RenderStyles.Item("StairsRailingStyle")
            Catch ex As Exception
                RS = oStairsDoc.RenderStyles.Add("StairsRailingStyle")
                RS.Reflectivity = 3
                RS.SetDiffuseColor(255, 255, 0)  'Yellow
            End Try
            'assign new render style to part feature
            'Dim oRailFeature As PartFeature = oExtrude
            Call oSweep.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSweep2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oOuterSpilPattern.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpiralLowOutRail.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpiralUpOutRail.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpil1Extrude.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)



            ProgressBar5.Value = 95


            If oCSState = "CCW_Right" Then GoTo CCWend




CCWLeftStart:

            '----------------Sketch First Inside Spil ----------------------
            '----------------Spil on 15cm from stairs front------------------

            Dim oSpil2Plane As WorkPlane
            oSpil2Plane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(3), oStepheight / 10, False)
            Dim oSpil2Sketch As PlanarSketch
            oSpil2Sketch = oDef.Sketches.Add(oSpil2Plane)
            Dim oCirc04 As SketchCircle
            oCirc04 = oSpil2Sketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.4, 15), 3.37 / 2)

            '--------------Extrusion of First Inside Spil--------------------------------------

            Dim oSpil2Prof As Profile
            oSpil2Prof = oSpil2Sketch.Profiles.AddForSolid
            Dim oSpil2ExtrudeDef As ExtrudeDefinition
            oSpil2ExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSpil2Prof, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSpil2ExtrudeDef.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oSpil2Extrude As ExtrudeFeature
            oSpil2Extrude = oDef.Features.ExtrudeFeatures.Add(oSpil2ExtrudeDef)

            oSpil2Plane.Visible = False

            '--------------Upper Inside railing--------------------------------------

            Dim oPnt5 As Point
            oPnt5 = oTG.CreatePoint(0, 15, oStepheight / 10 + 105)
            Dim oWorkPoint5 As WorkPoint
            oWorkPoint5 = oDef.WorkPoints.AddFixed(oPnt5, False)
            Dim oPnt6 As Point
            oPnt6 = oTG.CreatePoint(oCSIntRadius / 10, 15, oStepheight / 10 + 105)
            Dim oWorkPoint6 As WorkPoint
            oWorkPoint6 = oDef.WorkPoints.AddFixed(oPnt6, False)
            Dim oWorkAxis3 As WorkAxis
            oWorkAxis3 = oDef.WorkAxes.AddByTwoPoints(oWorkPoint5, oWorkPoint6, False)
            Dim oUpprRailWorkPlane2 As WorkPlane
            oUpprRailWorkPlane2 = oDef.WorkPlanes.AddByLinePlaneAndAngle(oWorkAxis3, oDef.WorkPlanes(3), (180 - oStairsAngle) * pi / 180)
            Dim oUpprRailSketch2 As PlanarSketch
            oUpprRailSketch2 = oDef.Sketches.Add(oUpprRailWorkPlane2)
            Dim oCirc05 As SketchCircle
            oCirc05 = oUpprRailSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.4, 0), 4.21 / 2)
            Dim oUpOutRailProf2 As Profile
            oUpOutRailProf2 = oUpprRailSketch2.Profiles.AddForSolid
            Dim oSpiralUpOutRail2 As CoilFeature
            oSpiralUpOutRail2 = oDef.Features.CoilFeatures.AddByRevolutionAndHeight(oUpOutRailProf2, oDef.WorkAxes.Item(3), oStairRev, (oCSHeight - oStepheight) / 10, PartFeatureOperationEnum.kJoinOperation, False,)


            oWorkPoint5.Visible = False
            oWorkPoint6.Visible = False
            oWorkAxis3.Visible = False
            oUpprRailWorkPlane2.Visible = False

            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 80
            oInvApp.ScreenUpdating = False



            '--------------Lower Inside railing--------------------------------------

            Dim oPnt7 As Point
            oPnt7 = oTG.CreatePoint(0, 15, oStepheight / 10 + 55)
            Dim oWorkPoint7 As WorkPoint
            oWorkPoint7 = oDef.WorkPoints.AddFixed(oPnt7, False)
            Dim oPnt8 As Point
            oPnt8 = oTG.CreatePoint(oCSIntRadius / 10, 15, oStepheight / 10 + 55)
            Dim oWorkPoint8 As WorkPoint
            oWorkPoint8 = oDef.WorkPoints.AddFixed(oPnt8, False)
            Dim oWorkAxis4 As WorkAxis
            oWorkAxis4 = oDef.WorkAxes.AddByTwoPoints(oWorkPoint7, oWorkPoint8, False)
            Dim oLowprRailWorkPlane2 As WorkPlane
            oLowprRailWorkPlane2 = oDef.WorkPlanes.AddByLinePlaneAndAngle(oWorkAxis4, oDef.WorkPlanes(3), (180 - oStairsAngle) * pi / 180)
            Dim oLowerRailSketch2 As PlanarSketch
            oLowerRailSketch2 = oDef.Sketches.Add(oLowprRailWorkPlane2)
            Dim oCirc06 As SketchCircle
            oCirc06 = oLowerRailSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.4, 0), 4.21 / 2)
            Dim oLowOutRailProf2 As Profile
            oLowOutRailProf2 = oLowerRailSketch2.Profiles.AddForSolid
            Dim oSpiralLowOutRail2 As CoilFeature
            oSpiralLowOutRail2 = oDef.Features.CoilFeatures.AddByRevolutionAndHeight(oLowOutRailProf2, oDef.WorkAxes.Item(3), oStairRev, (oCSHeight - oStepheight) / 10, PartFeatureOperationEnum.kJoinOperation, False,)
            'oSpiralLowOutRail2.Name = "Lower Inside Railing"
            oWorkPoint7.Visible = False
            oWorkPoint8.Visible = False
            oWorkAxis4.Visible = False
            oLowprRailWorkPlane2.Visible = False
            Dim oInnerSpilCol As ObjectCollection
            oInnerSpilCol = (oSpil2Extrude.Definition.AffectedBodies)
            oNumberOfSpils = oNumberOfSteps / 4
            If oNumberOfSpils <= 1 Then
                oNumberOfSpils = 2
            End If
            Dim oInnerSpilPattern As RectangularPatternFeature
            oInnerSpilPattern = oDef.Features.RectangularPatternFeatures.Add(oInnerSpilCol, oSpiralIntSide.SurfaceBody.Edges.Item(17), True, oNumberOfSpils, oLenghtSpiralIntStairs / 10 / (oNumberOfSpils - 1), kDefault,,,,,,,,, PatternOrientationEnum.kAdjustToDirection1)

            '----------------Intside Lower Railing Elbow ----------------------

            Dim oLowerItsideElbowPlane As WorkPlane
            oLowerItsideElbowPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(1), oCSIntRadius / 10 + 0.4, False)
            Dim oLowerInsideElbowSketch As PlanarSketch
            oLowerInsideElbowSketch = oDef.Sketches.Add(oLowerItsideElbowPlane)
            Dim oWp11 As SketchPoints
            oWp11 = oLowerInsideElbowSketch.SketchPoints
            Call oWp11.Add(oTG.CreatePoint2d(15, 105 + oStepheight / 10), False)
            Call oWp11.Add(oTG.CreatePoint2d(0, 105 + oStepheight / 10 - (Math.Tan(oAngle * ((oCSIntRadius + 1.2 + oCSWidth) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15)), False)
            Call oWp11.Add(oTG.CreatePoint2d(0, 55 + oStepheight / 10 - (Math.Tan(oAngle * ((oCSIntRadius + 1.2 + oCSWidth) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15)), False)
            Call oWp11.Add(oTG.CreatePoint2d(15, 55 + oStepheight / 10), False)
            Dim oLn40 As SketchLine
            oLn40 = oLowerInsideElbowSketch.SketchLines.AddByTwoPoints(oWp11(1), oWp11(2))
            Dim oLn41 As SketchLine
            oLn41 = oLowerInsideElbowSketch.SketchLines.AddByTwoPoints(oWp11(2), oWp11(3))
            Dim oLn42 As SketchLine
            oLn42 = oLowerInsideElbowSketch.SketchLines.AddByTwoPoints(oWp11(3), oWp11(4))
            Dim oArc8 As SketchArc
            oArc8 = oLowerInsideElbowSketch.SketchArcs.AddByFillet(oLn40, oLn41, 7.5, oLn40.Geometry.MidPoint, oLn41.Geometry.MidPoint)
            Dim oArc9 As SketchArc
            oArc9 = oLowerInsideElbowSketch.SketchArcs.AddByFillet(oLn41, oLn42, 7.5, oLn41.Geometry.MidPoint, oLn42.Geometry.MidPoint)
            Dim oPath3 As Path
            oPath3 = oDef.Features.CreatePath(oLn42)
            Dim oRailProfile3 As Profile
            oRailProfile3 = oUpprRailSketch2.Profiles.AddForSolid
            Dim oSweep3 As SweepFeature
            oSweep3 = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile3, oPath3, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)


            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 95
            oInvApp.ScreenUpdating = False



            oLowerItsideElbowPlane.Visible = False

            '----------------Upper Inside Railing Elbow ----------------------

            Dim oUpperInsideElbowPlane As WorkPlane
            oUpperInsideElbowPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oLn22, oTopStepFrontPlane, 270 * pi / 180, False)
            Dim oUpperInsideElbowSketch As PlanarSketch
            oUpperInsideElbowSketch = oDef.Sketches.Add(oUpperInsideElbowPlane)
            Dim oWp07 As SketchPoints
            oWp07 = oUpperInsideElbowSketch.SketchPoints
            Call oWp07.Add(oTG.CreatePoint2d(-4 - 55, -15), False)
            Call oWp07.Add(oTG.CreatePoint2d(-4 - 55 - (Math.Tan(oAngle * ((oCSIntRadius + 1.2 + oCSWidth) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15), -30), False)
            Call oWp07.Add(oTG.CreatePoint2d(-4 - 105 - (Math.Tan(oAngle * ((oCSIntRadius + 1.2 + oCSWidth) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15), -30), False)
            Call oWp07.Add(oTG.CreatePoint2d(-4 - 105, -15), False)
            Dim oLn53 As SketchLine
            oLn53 = oUpperInsideElbowSketch.SketchLines.AddByTwoPoints(oWp07(1), oWp07(2))
            Dim oLn54 As SketchLine
            oLn54 = oUpperInsideElbowSketch.SketchLines.AddByTwoPoints(oWp07(2), oWp07(3))
            Dim oLn55 As SketchLine
            oLn55 = oUpperInsideElbowSketch.SketchLines.AddByTwoPoints(oWp07(3), oWp07(4))
            Dim oArc10 As SketchArc
            oArc10 = oUpperInsideElbowSketch.SketchArcs.AddByFillet(oLn53, oLn54, 7.5, oLn53.Geometry.MidPoint, oLn54.Geometry.MidPoint)
            Dim oArc11 As SketchArc
            oArc11 = oUpperInsideElbowSketch.SketchArcs.AddByFillet(oLn54, oLn55, 7.5, oLn54.Geometry.MidPoint, oLn55.Geometry.MidPoint)
            Dim oHelpPln3 As WorkPlane
            oHelpPln3 = oDef.WorkPlanes.AddByPlaneAndOffset(oTopStepFrontPlane, 15, False)
            Dim oHelpSketch2 As PlanarSketch
            oHelpSketch2 = oDef.Sketches.Add(oHelpPln3)
            Dim oHelpPnt2 As SketchPoints
            oHelpPnt2 = oHelpSketch2.SketchPoints
            Call oHelpPnt2.Add(oTG.CreatePoint2d(oCSHeight / 10 + 105, 0), False)
            Call oHelpPnt2.Add(oTG.CreatePoint2d(oCSHeight / 10 + 105, oCSIntRadius / 10), False)
            Dim oHelpLn2 As SketchLine
            oHelpLn2 = oHelpSketch2.SketchLines.AddByTwoPoints(oHelpPnt2(1), oHelpPnt2(2))
            Dim oHelpPln4 As WorkPlane
            oHelpPln4 = oDef.WorkPlanes.AddByLinePlaneAndAngle(oHelpLn2, oHelpPln3, oAngle, False)
            Dim oUpInRailElbSketch As PlanarSketch
            oUpInRailElbSketch = oDef.Sketches.Add(oHelpPln4)
            Dim oCirc11 As SketchCircle
            oCirc11 = oUpInRailElbSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.4, 0), 4.21 / 2)
            Dim oPath4 As Path
            oPath4 = oDef.Features.CreatePath(oLn54)
            Dim oRailProfile4 As Profile
            oRailProfile4 = oUpInRailElbSketch.Profiles.AddForSolid
            Dim oSweep4 As SweepFeature
            oSweep4 = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile4, oPath4, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)


            Try
                RS = oStairsDoc.RenderStyles.Item("StairsRailingStyle")
            Catch ex As Exception
                RS = oStairsDoc.RenderStyles.Add("StairsRailingStyle")
                RS.Reflectivity = 3
                RS.SetDiffuseColor(255, 255, 0)  'Yellow
            End Try
            'assign new render style to part feature
            'Dim oRailFeature As PartFeature = oExtrude
            Call oSweep3.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSweep4.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oInnerSpilPattern.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpiralLowOutRail2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpiralUpOutRail2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpil2Extrude.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)


            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 100
            oInvApp.ScreenUpdating = False

            oUpperInsideElbowPlane.Visible = False
            oHelpPln3.Visible = False
            oHelpPln4.Visible = False
            oHelpSketch2.Visible = False

CCWend:
        End If



        If oCSState = "CW_None" Or oCSState = "CW_Right" Or oCSState = "CW_Left" Or oCSState = "CW_Both" Then


            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 10
            oInvApp.ScreenUpdating = False


            '----------------Workplane Tankside ----------------------
            Dim oYplane As WorkPlane
            Dim oIntSidePlane As WorkPlane
            oYplane = oDef.WorkPlanes.Item(1)
            oIntSidePlane = oDef.WorkPlanes.AddByPlaneAndOffset(oYplane, oCSIntRadius / 10, False)
            oIntSidePlane.Visible = False
            '----------------Sketch Tankside ----------------------
            Dim oTankSideSketch As PlanarSketch
            oTankSideSketch = oDef.Sketches.Add(oIntSidePlane)
            Dim oWp01 As SketchPoints
            oWp01 = oTankSideSketch.SketchPoints
            ProgressBar5.Value = 30
            '------------Sketch points side----------------------
            Call oWp01.Add(oTG.CreatePoint2d(0, 0), False)
            Call oWp01.Add(oTG.CreatePoint2d(0, (oStepheight / 10)), False)
            Call oWp01.Add(oTG.CreatePoint2d(Math.Cos(((90 - oCSAngle) * pi / 180) / 2) * -26, ((oStepheight / 10) - (Math.Sin(((90 - oCSAngle) * pi / 180) / 2) * 26))), False)
            Call oWp01.Add(oTG.CreatePoint2d(Math.Cos(((90 - oCSAngle) * pi / 180) / 2) * -26, 0), False)
            '-------------Draw lines----------------------------------------------------
            Dim oLn01 As SketchLine
            oLn01 = oTankSideSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
            Dim oLn02 As SketchLine
            oLn02 = oTankSideSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
            Dim oLn03 As SketchLine
            oLn03 = oTankSideSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
            Dim oLn04 As SketchLine
            oLn04 = oTankSideSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(1))
            '--------------Extrusion of the TankSideGroundplate--------------------------------------
            Dim oGPlateProf As Profile
            oGPlateProf = oTankSideSketch.Profiles.AddForSolid
            Dim oGPlateExtrudeDef As ExtrudeDefinition
            Dim kJoinOperation As Object = Nothing
            oGPlateExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oGPlateProf, PartFeatureOperationEnum.kJoinOperation)
            Call oGPlateExtrudeDef.SetDistanceExtent(0.8, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oGPlateExtrude As ExtrudeFeature
            oGPlateExtrude = oDef.Features.ExtrudeFeatures.Add(oGPlateExtrudeDef)


            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(1)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------


            '--------------Pattern of the TankSideGroundplate--------------------------------------
            Dim oGPlateCol As ObjectCollection
            oGPlateCol = (oGPlateExtrude.Definition.AffectedBodies)
            Dim oGPlatePattern As RectangularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing
            oGPlatePattern = oDef.Features.RectangularPatternFeatures.Add(oGPlateCol, oDef.WorkPlanes.Item(1), True, 2, oCSWidth / 10 + 0.8, kDefault)

            Dim oSpiralSketch As PlanarSketch
            oSpiralSketch = oDef.Sketches.Add(oGPlateExtrude.SurfaceBody.Faces.Item(3))
            Dim oLn05 As SketchEntity
            oLn05 = oSpiralSketch.AddByProjectingEntity(oGPlateExtrude.SurfaceBody.Faces.Item(3).Edges.Item(1))
            Dim oLn06 As SketchEntity
            oLn06 = oSpiralSketch.AddByProjectingEntity(oGPlateExtrude.SurfaceBody.Faces.Item(3).Edges.Item(2))
            Dim oLn07 As SketchEntity
            oLn07 = oSpiralSketch.AddByProjectingEntity(oGPlateExtrude.SurfaceBody.Faces.Item(3).Edges.Item(3))
            Dim oLn08 As SketchEntity
            oLn08 = oSpiralSketch.AddByProjectingEntity(oGPlateExtrude.SurfaceBody.Faces.Item(3).Edges.Item(4))
            Dim oLn09 As SketchEntity
            oLn09 = oSpiralSketch.AddByProjectingEntity(oGPlatePattern.PatternElements.Item(2).Faces.Item(3).Edges.Item(1))
            Dim oLn10 As SketchEntity
            oLn10 = oSpiralSketch.AddByProjectingEntity(oGPlatePattern.PatternElements.Item(2).Faces.Item(3).Edges.Item(2))
            Dim oLn11 As SketchEntity
            oLn11 = oSpiralSketch.AddByProjectingEntity(oGPlatePattern.PatternElements.Item(2).Faces.Item(3).Edges.Item(3))
            Dim oLn12 As SketchEntity
            oLn12 = oSpiralSketch.AddByProjectingEntity(oGPlatePattern.PatternElements.Item(2).Faces.Item(3).Edges.Item(4))
            Dim oIntSpiralProf As Profile
            oIntSpiralProf = oSpiralSketch.Profiles.AddForSolid
            Dim oSpiralIntSide As CoilFeature
            oSpiralIntSide = oDef.Features.CoilFeatures.AddByRevolutionAndHeight(oIntSpiralProf, oDef.WorkAxes.Item(3), oStairRev, (oCSHeight - oStepheight) / 10, PartFeatureOperationEnum.kJoinOperation, False, True, 0,,,,,,)


            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 20
            oInvApp.ScreenUpdating = False



            '--------------Cut of the Spirals--------------------------------------

            Dim oTopStepFrontPlane As WorkPlane
            Dim oCenterline As WorkAxis
            oCenterline = oDef.WorkAxes(3)
            oTopStepFrontPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oCenterline, oDef.WorkPlanes(2), pi - oStairRevAngle, False)
            'oTopStepFrontPlane.Name = "Top Step Front Plane"
            oTopStepFrontPlane.Visible = False
            Dim oCutSketch As PlanarSketch
            oCutSketch = oDef.Sketches.Add(oTopStepFrontPlane)
            Dim oWp04 As SketchPoints
            oWp04 = oCutSketch.SketchPoints
            Call oWp04.Add(oTG.CreatePoint2d(oCSHeight / 10, -oCSIntRadius / 10 + 5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oCSHeight / 10, -oCSIntRadius / 10 - 0.8 - oCSWidth / 10 - 0.8 - 5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, -oCSIntRadius / 10 - 0.8 - oCSWidth / 10 - 0.8 - 5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, -oCSIntRadius / 10 + 5), False)
            Dim oLn25 As SketchLine
            oLn25 = oCutSketch.SketchLines.AddByTwoPoints(oWp04(1), oWp04(2))
            Dim oLn26 As SketchLine
            oLn26 = oCutSketch.SketchLines.AddByTwoPoints(oWp04(2), oWp04(3))
            Dim oLn27 As SketchLine
            oLn27 = oCutSketch.SketchLines.AddByTwoPoints(oWp04(3), oWp04(4))
            Dim oLn28 As SketchLine
            oLn28 = oCutSketch.SketchLines.AddByTwoPoints(oWp04(4), oWp04(1))

            '--------------Cut of Stairs End--------------------------------------

            Dim oCutProf As Profile
            oCutProf = oCutSketch.Profiles.AddForSolid
            Dim oCutExtrudeDef As ExtrudeDefinition
            oCutExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oCutProf, PartFeatureOperationEnum.kCutOperation)
            Call oCutExtrudeDef.SetDistanceExtent(26, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oCutExtrude As ExtrudeFeature
            oCutExtrude = oDef.Features.ExtrudeFeatures.Add(oCutExtrudeDef)

            '------------Sketch points Stairs End----------------------

            Dim oStairEndSketch As PlanarSketch
            oStairEndSketch = oDef.Sketches.Add(oTopStepFrontPlane)
            Dim oWp03 As SketchPoints
            oWp03 = oStairEndSketch.SketchPoints
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10, -oCSIntRadius / 10), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 4, -oCSIntRadius / 10 - 0.8), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, -oCSIntRadius / 10 - 0.8), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, -oCSIntRadius / 10), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 4, -oCSIntRadius / 10 - oCSWidth / 10 - 0.8), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10, -oCSIntRadius / 10 - 0.8 - oCSWidth / 10 - 0.8), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, -oCSIntRadius / 10 - 0.8 - oCSWidth / 10 - 0.8), False)
            Call oWp03.Add(oTG.CreatePoint2d(oCSHeight / 10 - 20, -oCSIntRadius / 10 - oCSWidth / 10 - 0.8), False)

            '-------------Draw lines----------------------------------------------------

            Dim oLn17 As SketchLine
            oLn17 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(1), oWp03(6))
            Dim oLn18 As SketchLine
            oLn18 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(6), oWp03(7))
            Dim oLn19 As SketchLine
            oLn19 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(7), oWp03(8))
            Dim oLn20 As SketchLine
            oLn20 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(8), oWp03(5))
            Dim oLn21 As SketchLine
            oLn21 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(5), oWp03(2))
            Dim oLn22 As SketchLine
            oLn22 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(2), oWp03(3))
            Dim oLn23 As SketchLine
            oLn23 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(3), oWp03(4))
            Dim oLn24 As SketchLine
            oLn24 = oStairEndSketch.SketchLines.AddByTwoPoints(oWp03(4), oWp03(1))
            ProgressBar5.Value = 50

            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 30
            oInvApp.ScreenUpdating = False

            '--------------Extrusion of Stairs End--------------------------------------

            Dim oEndProf As Profile
            oEndProf = oStairEndSketch.Profiles.AddForSolid
            Dim oENdExtrudeDef As ExtrudeDefinition
            oENdExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oEndProf, PartFeatureOperationEnum.kNewBodyOperation)
            Call oENdExtrudeDef.SetDistanceExtent(26, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oEndExtrude As ExtrudeFeature
            oEndExtrude = oDef.Features.ExtrudeFeatures.Add(oENdExtrudeDef)

            '----------------Sketch Step 1 ---------------------

            Dim oStep1Plane As WorkPlane
            oStep1Plane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(3), oStepheight / 10, False)
            Dim oStep1Sketch As PlanarSketch
            oStep1Sketch = oDef.Sketches.Add(oStep1Plane)
            Dim oWp02 As SketchPoints
            oWp02 = oStep1Sketch.SketchPoints
            Call oWp02.Add(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8, 0), False)
            Call oWp02.Add(oTG.CreatePoint2d(oCSIntRadius / 10, -26), False)
            Call oWp02.Add(oTG.CreatePoint2d(oCSIntRadius / 10 + oCSWidth / 10, -26), False)
            Call oWp02.Add(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8 + oCSWidth / 10, 0), False)
            Dim oLn13 As SketchLine
            oLn13 = oStep1Sketch.SketchLines.AddByTwoPoints(oWp02(1), oWp02(2))
            Dim oLn14 As SketchLine
            oLn14 = oStep1Sketch.SketchLines.AddByTwoPoints(oWp02(2), oWp02(3))
            Dim oLn15 As SketchLine
            oLn15 = oStep1Sketch.SketchLines.AddByTwoPoints(oWp02(3), oWp02(4))
            Dim oLn16 As SketchLine
            oLn16 = oStep1Sketch.SketchLines.AddByTwoPoints(oWp02(4), oWp02(1))

            '--------------Extrusion of Step 1--------------------------------------

            Dim oStep1Prof As Profile
            oStep1Prof = oStep1Sketch.Profiles.AddForSolid
            Dim oStep1ExtrudeDef As ExtrudeDefinition
            oStep1ExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oStep1Prof, PartFeatureOperationEnum.kNewBodyOperation)
            Call oStep1ExtrudeDef.SetDistanceExtent(4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oStep1Extrude As ExtrudeFeature
            oStep1Extrude = oDef.Features.ExtrudeFeatures.Add(oStep1ExtrudeDef)
            'oStep1Extrude.Name = "First Step"
            oStep1Plane.Visible = False

            '--------------Pattern of the Steps--------------------------------------

            Dim oStepsCol As ObjectCollection
            oStepsCol = (oStep1Extrude.Definition.AffectedBodies)
            Dim oStepsPattern As RectangularPatternFeature
            oStepsPattern = oDef.Features.RectangularPatternFeatures.Add(oStepsCol, oSpiralIntSide.SurfaceBody.Edges.Item(40), True, oNumberOfSteps - 1, oLenghtSpiralOutStairs / 10 / (oNumberOfSteps - 1), kDefault,,,,,,,,, PatternOrientationEnum.kAdjustToDirection1)

            If oCSState = "CW_None" Then GoTo CWend
            If oCSState = "CW_Right" Then GoTo CWRightStart



            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 40
            oInvApp.ScreenUpdating = False


            '----------------Sketch First Outside Spil ----------------------
            '----------------Spil on 15cm from stairs front------------------

            Dim oSpil1Plane As WorkPlane
            oSpil1Plane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(3), oStepheight / 10, False)
            Dim oSpil1Sketch As PlanarSketch
            oSpil1Sketch = oDef.Sketches.Add(oSpil1Plane)
            Dim oCirc01 As SketchCircle
            oCirc01 = oSpil1Sketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.4, -15), 3.37 / 2)

            '--------------Extrusion of First Outside Spil--------------------------------------

            Dim oSpil1Prof As Profile
            oSpil1Prof = oSpil1Sketch.Profiles.AddForSolid
            Dim oSpil1ExtrudeDef As ExtrudeDefinition
            oSpil1ExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSpil1Prof, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSpil1ExtrudeDef.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oSpil1Extrude As ExtrudeFeature
            oSpil1Extrude = oDef.Features.ExtrudeFeatures.Add(oSpil1ExtrudeDef)
            oSpil1Plane.Visible = False

            '--------------Upper outer railing--------------------------------------

            Dim oPnt1 As Point
            oPnt1 = oTG.CreatePoint(0, -15, oStepheight / 10 + 105)
            Dim oWorkPoint1 As WorkPoint
            oWorkPoint1 = oDef.WorkPoints.AddFixed(oPnt1, False)
            Dim oPnt2 As Point
            oPnt2 = oTG.CreatePoint(oCSIntRadius / 10, -15, oStepheight / 10 + 105)
            Dim oWorkPoint2 As WorkPoint
            oWorkPoint2 = oDef.WorkPoints.AddFixed(oPnt2, False)
            Dim oWorkAxis1 As WorkAxis
            oWorkAxis1 = oDef.WorkAxes.AddByTwoPoints(oWorkPoint1, oWorkPoint2, False)
            Dim oUpprRailWorkPlane As WorkPlane
            oUpprRailWorkPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oWorkAxis1, oDef.WorkPlanes(3), oAngle)
            Dim oUpprRailSketch As PlanarSketch
            oUpprRailSketch = oDef.Sketches.Add(oUpprRailWorkPlane)
            Dim oCirc02 As SketchCircle
            oCirc02 = oUpprRailSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.4, 0), 4.21 / 2)
            Dim oUpOutRailProf As Profile
            oUpOutRailProf = oUpprRailSketch.Profiles.AddForSolid
            Dim oSpiralUpOutRail As CoilFeature
            oSpiralUpOutRail = oDef.Features.CoilFeatures.AddByRevolutionAndHeight(oUpOutRailProf, oDef.WorkAxes.Item(3), oStairRev, (oCSHeight - oStepheight) / 10, PartFeatureOperationEnum.kJoinOperation, False, True, 0,,,,,,)
            oWorkPoint1.Visible = False
            oWorkPoint2.Visible = False
            oWorkAxis1.Visible = False
            oUpprRailWorkPlane.Visible = False

            '--------------Lower outer railing--------------------------------------

            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 50
            oInvApp.ScreenUpdating = False


            Dim oPnt3 As Point
            oPnt3 = oTG.CreatePoint(0, -15, oStepheight / 10 + 55)
            Dim oWorkPoint3 As WorkPoint
            oWorkPoint3 = oDef.WorkPoints.AddFixed(oPnt3, False)
            Dim oPnt4 As Point
            oPnt4 = oTG.CreatePoint(oCSIntRadius / 10, -15, oStepheight / 10 + 55)
            Dim oWorkPoint4 As WorkPoint
            oWorkPoint4 = oDef.WorkPoints.AddFixed(oPnt4, False)
            Dim oWorkAxis2 As WorkAxis
            oWorkAxis2 = oDef.WorkAxes.AddByTwoPoints(oWorkPoint3, oWorkPoint4, False)
            Dim oLowprRailWorkPlane As WorkPlane
            oLowprRailWorkPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oWorkAxis2, oDef.WorkPlanes(3), oAngle)
            Dim oLowerRailSketch As PlanarSketch
            oLowerRailSketch = oDef.Sketches.Add(oLowprRailWorkPlane)
            Dim oCirc03 As SketchCircle
            oCirc03 = oLowerRailSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.4, 0), 4.21 / 2)
            Dim oLowOutRailProf As Profile
            oLowOutRailProf = oLowerRailSketch.Profiles.AddForSolid
            Dim oSpiralLowOutRail As CoilFeature
            oSpiralLowOutRail = oDef.Features.CoilFeatures.AddByRevolutionAndHeight(oLowOutRailProf, oDef.WorkAxes.Item(3), oStairRev, (oCSHeight - oStepheight) / 10, PartFeatureOperationEnum.kJoinOperation, False, True, 0,,,,,,)
            oWorkPoint3.Visible = False
            oWorkPoint4.Visible = False
            oWorkAxis2.Visible = False
            oLowprRailWorkPlane.Visible = False
            Dim oOuterSpilCol As ObjectCollection
            oOuterSpilCol = (oSpil1Extrude.Definition.AffectedBodies)
            oNumberOfSpils = oNumberOfSteps / 4
            If oNumberOfSpils <= 1 Then
                oNumberOfSpils = 2
            End If
            Dim oOuterSpilPattern As RectangularPatternFeature
            oOuterSpilPattern = oDef.Features.RectangularPatternFeatures.Add(oOuterSpilCol, oSpiralIntSide.SurfaceBody.Edges.Item(40), True, oNumberOfSpils, oLenghtSpiralOutStairs / 10 / (oNumberOfSpils - 1), kDefault,,,,,,,,, PatternOrientationEnum.kAdjustToDirection1)

            '----------------Sketch Outside Lower Railing Elbow ---------------------

            Dim oLowerOutsideElbowPlane As WorkPlane
            oLowerOutsideElbowPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(1), oCSIntRadius / 10 + 0.8 + oCSWidth / 10 + 0.4, False)
            Dim oLowerOutsideElbowSketch As PlanarSketch
            oLowerOutsideElbowSketch = oDef.Sketches.Add(oLowerOutsideElbowPlane)
            Dim oWp05 As SketchPoints
            oWp05 = oLowerOutsideElbowSketch.SketchPoints
            Call oWp05.Add(oTG.CreatePoint2d(-15, 105 + oStepheight / 10), False)
            Call oWp05.Add(oTG.CreatePoint2d(0, 105 + oStepheight / 10 - (Math.Tan(oAngle * ((oCSIntRadius + 0.4) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15)), False)
            Call oWp05.Add(oTG.CreatePoint2d(0, 55 + oStepheight / 10 - (Math.Tan(oAngle * ((oCSIntRadius + 0.4) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15)), False)
            Call oWp05.Add(oTG.CreatePoint2d(-15, 55 + oStepheight / 10), False)
            Dim oLn30 As SketchLine
            oLn30 = oLowerOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp05(1), oWp05(2))
            Dim oLn31 As SketchLine
            oLn31 = oLowerOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp05(2), oWp05(3))
            Dim oLn32 As SketchLine
            oLn32 = oLowerOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp05(3), oWp05(4))
            Dim oArc4 As SketchArc
            oArc4 = oLowerOutsideElbowSketch.SketchArcs.AddByFillet(oLn30, oLn31, 7.5, oLn30.Geometry.MidPoint, oLn31.Geometry.MidPoint)
            Dim oArc5 As SketchArc
            oArc5 = oLowerOutsideElbowSketch.SketchArcs.AddByFillet(oLn31, oLn32, 7.5, oLn31.Geometry.MidPoint, oLn32.Geometry.MidPoint)
            Dim oPath1 As Path
            oPath1 = oDef.Features.CreatePath(oLn32)
            Dim oRailProfile1 As Profile
            oRailProfile1 = oUpprRailSketch.Profiles.AddForSolid
            Dim oSweep As SweepFeature
            oSweep = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile1, oPath1, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)
            oLowerOutsideElbowPlane.Visible = False



            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 60
            oInvApp.ScreenUpdating = False



            '----------------Upper Outside Railing Elbow ----------------------

            Dim oUpperOutsideElbowPlane As WorkPlane
            oUpperOutsideElbowPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oLn20, oTopStepFrontPlane, 90 * pi / 180, False)
            Dim oUpperOutsideElbowSketch As PlanarSketch
            oUpperOutsideElbowSketch = oDef.Sketches.Add(oUpperOutsideElbowPlane)
            Dim oWp06 As SketchPoints
            oWp06 = oUpperOutsideElbowSketch.SketchPoints
            Call oWp06.Add(oTG.CreatePoint2d(oStepheight / 10 + 105, 15), False)
            Call oWp06.Add(oTG.CreatePoint2d(oStepheight / 10 + 105 + (Math.Tan(oAngle * ((oCSIntRadius + 0.4) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15), 30), False)
            Call oWp06.Add(oTG.CreatePoint2d(oStepheight / 10 + 55 + (Math.Tan(oAngle * ((oCSIntRadius + 0.4) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15), 30), False)
            Call oWp06.Add(oTG.CreatePoint2d(oStepheight / 10 + 55, 15), False)
            Dim oLn33 As SketchLine
            oLn33 = oUpperOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp06(1), oWp06(2))
            Dim oLn34 As SketchLine
            oLn34 = oUpperOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp06(2), oWp06(3))
            Dim oLn35 As SketchLine
            oLn35 = oUpperOutsideElbowSketch.SketchLines.AddByTwoPoints(oWp06(3), oWp06(4))
            Dim oArc6 As SketchArc
            oArc6 = oUpperOutsideElbowSketch.SketchArcs.AddByFillet(oLn33, oLn34, 7.5, oLn33.Geometry.MidPoint, oLn34.Geometry.MidPoint)
            Dim oArc7 As SketchArc
            oArc7 = oUpperOutsideElbowSketch.SketchArcs.AddByFillet(oLn34, oLn35, 7.5, oLn34.Geometry.MidPoint, oLn35.Geometry.MidPoint)
            Dim oHelpPln1 As WorkPlane
            oHelpPln1 = oDef.WorkPlanes.AddByPlaneAndOffset(oTopStepFrontPlane, 15, False)
            Dim oHelpSketch As PlanarSketch
            oHelpSketch = oDef.Sketches.Add(oHelpPln1)
            Dim oHelpPnt1 As SketchPoints
            oHelpPnt1 = oHelpSketch.SketchPoints
            Call oHelpPnt1.Add(oTG.CreatePoint2d(oCSHeight / 10 + 105, 0), False)
            Call oHelpPnt1.Add(oTG.CreatePoint2d(oCSHeight / 10 + 105, oCSIntRadius / 10), False)
            Dim oHelpLn1 As SketchLine
            oHelpLn1 = oHelpSketch.SketchLines.AddByTwoPoints(oHelpPnt1(1), oHelpPnt1(2))
            Dim oHelpPln2 As WorkPlane
            oHelpPln2 = oDef.WorkPlanes.AddByLinePlaneAndAngle(oHelpLn1, oHelpPln1, oAngle, False)
            Dim oUpOutRailElbSketch As PlanarSketch
            oUpOutRailElbSketch = oDef.Sketches.Add(oHelpPln2)
            Dim oCirc10 As SketchCircle
            oCirc10 = oUpOutRailElbSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(-oCSIntRadius / 10 - 0.8 - oCSWidth / 10 - 0.4, 0), 4.21 / 2)
            Dim oPath2 As Path
            oPath2 = oDef.Features.CreatePath(oLn34)
            Dim oRailProfile2 As Profile
            oRailProfile2 = oUpOutRailElbSketch.Profiles.AddForSolid
            Dim oSweep2 As SweepFeature
            oSweep2 = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile2, oPath2, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)

            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 70
            oInvApp.ScreenUpdating = False

            oUpperOutsideElbowPlane.Visible = False
            oHelpPln1.Visible = False
            oHelpPln2.Visible = False
            oHelpSketch.Visible = False
            Dim RS As RenderStyle
            Try
                RS = oStairsDoc.RenderStyles.Item("StairsRailingStyle")
            Catch ex As Exception
                RS = oStairsDoc.RenderStyles.Add("StairsRailingStyle")
                RS.Reflectivity = 3
                RS.SetDiffuseColor(255, 255, 0)  'Yellow
            End Try

            Call oSweep.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSweep2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oOuterSpilPattern.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpiralLowOutRail.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpiralUpOutRail.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpil1Extrude.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            ProgressBar5.Value = 90

            If oCSState = "CW_Left" Then GoTo CWend

CWRightStart:

            '----------------Sketch First Inside Spil ----------------------
            '----------------Spil on 15cm from stairs front------------------

            Dim oSpil2Plane As WorkPlane
            oSpil2Plane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(3), oStepheight / 10, False)
            Dim oSpil2Sketch As PlanarSketch
            oSpil2Sketch = oDef.Sketches.Add(oSpil2Plane)
            Dim oCirc04 As SketchCircle
            oCirc04 = oSpil2Sketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.4, -15), 3.37 / 2)

            '--------------Extrusion of First Inside Spil--------------------------------------

            Dim oSpil2Prof As Profile
            oSpil2Prof = oSpil2Sketch.Profiles.AddForSolid
            Dim oSpil2ExtrudeDef As ExtrudeDefinition
            oSpil2ExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSpil2Prof, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSpil2ExtrudeDef.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oSpil2Extrude As ExtrudeFeature
            oSpil2Extrude = oDef.Features.ExtrudeFeatures.Add(oSpil2ExtrudeDef)
            oSpil2Plane.Visible = False


            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 80
            oInvApp.ScreenUpdating = False


            '--------------Upper Inside railing--------------------------------------

            Dim oPnt5 As Point
            oPnt5 = oTG.CreatePoint(0, -15, oStepheight / 10 + 105)
            Dim oWorkPoint5 As WorkPoint
            oWorkPoint5 = oDef.WorkPoints.AddFixed(oPnt5, False)
            Dim oPnt6 As Point
            oPnt6 = oTG.CreatePoint(oCSIntRadius / 10, -15, oStepheight / 10 + 105)
            Dim oWorkPoint6 As WorkPoint
            oWorkPoint6 = oDef.WorkPoints.AddFixed(oPnt6, False)
            Dim oWorkAxis3 As WorkAxis
            oWorkAxis3 = oDef.WorkAxes.AddByTwoPoints(oWorkPoint5, oWorkPoint6, False)
            Dim oUpprRailWorkPlane2 As WorkPlane
            oUpprRailWorkPlane2 = oDef.WorkPlanes.AddByLinePlaneAndAngle(oWorkAxis3, oDef.WorkPlanes(3), oAngle)
            Dim oUpprRailSketch2 As PlanarSketch
            oUpprRailSketch2 = oDef.Sketches.Add(oUpprRailWorkPlane2)
            Dim oCirc05 As SketchCircle
            oCirc05 = oUpprRailSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.4, 0), 4.21 / 2)
            Dim oUpOutRailProf2 As Profile
            oUpOutRailProf2 = oUpprRailSketch2.Profiles.AddForSolid
            Dim oSpiralUpOutRail2 As CoilFeature
            oSpiralUpOutRail2 = oDef.Features.CoilFeatures.AddByRevolutionAndHeight(oUpOutRailProf2, oDef.WorkAxes.Item(3), oStairRev, (oCSHeight - oStepheight) / 10, PartFeatureOperationEnum.kJoinOperation, False, True, 0,,,,,,)

            oWorkPoint5.Visible = False
            oWorkPoint6.Visible = False
            oWorkAxis3.Visible = False
            oUpprRailWorkPlane2.Visible = False

            '--------------Lower Inside railing--------------------------------------

            Dim oPnt7 As Point
            oPnt7 = oTG.CreatePoint(0, -15, oStepheight / 10 + 55)
            Dim oWorkPoint7 As WorkPoint
            oWorkPoint7 = oDef.WorkPoints.AddFixed(oPnt7, False)
            Dim oPnt8 As Point
            oPnt8 = oTG.CreatePoint(oCSIntRadius / 10, -15, oStepheight / 10 + 55)
            Dim oWorkPoint8 As WorkPoint
            oWorkPoint8 = oDef.WorkPoints.AddFixed(oPnt8, False)
            Dim oWorkAxis4 As WorkAxis
            oWorkAxis4 = oDef.WorkAxes.AddByTwoPoints(oWorkPoint7, oWorkPoint8, False)
            Dim oLowprRailWorkPlane2 As WorkPlane
            oLowprRailWorkPlane2 = oDef.WorkPlanes.AddByLinePlaneAndAngle(oWorkAxis4, oDef.WorkPlanes(3), oAngle)
            Dim oLowerRailSketch2 As PlanarSketch
            oLowerRailSketch2 = oDef.Sketches.Add(oLowprRailWorkPlane2)
            Dim oCirc06 As SketchCircle
            oCirc06 = oLowerRailSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oCSIntRadius / 10 + 0.4, 0), 4.21 / 2)
            Dim oLowOutRailProf2 As Profile
            oLowOutRailProf2 = oLowerRailSketch2.Profiles.AddForSolid
            Dim oSpiralLowOutRail2 As CoilFeature
            oSpiralLowOutRail2 = oDef.Features.CoilFeatures.AddByRevolutionAndHeight(oLowOutRailProf2, oDef.WorkAxes.Item(3), oStairRev, (oCSHeight - oStepheight) / 10, PartFeatureOperationEnum.kJoinOperation, False, True, 0,,,,,,)
            'oSpiralLowOutRail2.Name = "Lower Inside Railing"
            oWorkPoint7.Visible = False
            oWorkPoint8.Visible = False
            oWorkAxis4.Visible = False
            oLowprRailWorkPlane2.Visible = False
            Dim oInnerSpilCol As ObjectCollection
            oInnerSpilCol = (oSpil2Extrude.Definition.AffectedBodies)
            oNumberOfSpils = oNumberOfSteps / 4
            If oNumberOfSpils <= 1 Then
                oNumberOfSpils = 2
            End If
            Dim oInnerSpilPattern As RectangularPatternFeature
            oInnerSpilPattern = oDef.Features.RectangularPatternFeatures.Add(oInnerSpilCol, oSpiralIntSide.SurfaceBody.Edges.Item(17), True, oNumberOfSpils, oLenghtSpiralIntStairs / 10 / (oNumberOfSpils - 1), kDefault,,,,,,,,, PatternOrientationEnum.kAdjustToDirection1)

            '----------------Intside Lower Railing Elbow ----------------------

            Dim oLowerItsideElbowPlane As WorkPlane
            oLowerItsideElbowPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(1), oCSIntRadius / 10 + 0.4, False)
            Dim oLowerInsideElbowSketch As PlanarSketch
            oLowerInsideElbowSketch = oDef.Sketches.Add(oLowerItsideElbowPlane)
            Dim oWp11 As SketchPoints
            oWp11 = oLowerInsideElbowSketch.SketchPoints
            Call oWp11.Add(oTG.CreatePoint2d(-15, 105 + oStepheight / 10), False)
            Call oWp11.Add(oTG.CreatePoint2d(0, 105 + oStepheight / 10 - (Math.Tan(oAngle * ((oCSIntRadius + 1.2 + oCSWidth) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15)), False)
            Call oWp11.Add(oTG.CreatePoint2d(0, 55 + oStepheight / 10 - (Math.Tan(oAngle * ((oCSIntRadius + 1.2 + oCSWidth) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15)), False)
            Call oWp11.Add(oTG.CreatePoint2d(-15, 55 + oStepheight / 10), False)
            Dim oLn40 As SketchLine
            oLn40 = oLowerInsideElbowSketch.SketchLines.AddByTwoPoints(oWp11(1), oWp11(2))
            Dim oLn41 As SketchLine
            oLn41 = oLowerInsideElbowSketch.SketchLines.AddByTwoPoints(oWp11(2), oWp11(3))
            Dim oLn42 As SketchLine
            oLn42 = oLowerInsideElbowSketch.SketchLines.AddByTwoPoints(oWp11(3), oWp11(4))
            Dim oArc8 As SketchArc
            oArc8 = oLowerInsideElbowSketch.SketchArcs.AddByFillet(oLn40, oLn41, 7.5, oLn40.Geometry.MidPoint, oLn41.Geometry.MidPoint)
            Dim oArc9 As SketchArc
            oArc9 = oLowerInsideElbowSketch.SketchArcs.AddByFillet(oLn41, oLn42, 7.5, oLn41.Geometry.MidPoint, oLn42.Geometry.MidPoint)
            Dim oPath3 As Path
            oPath3 = oDef.Features.CreatePath(oLn42)
            Dim oRailProfile3 As Profile
            oRailProfile3 = oUpprRailSketch2.Profiles.AddForSolid
            Dim oSweep3 As SweepFeature
            oSweep3 = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile3, oPath3, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)



            oInvApp.ScreenUpdating = True
            ProgressBar5.Value = 90
            oInvApp.ScreenUpdating = False




            oLowerItsideElbowPlane.Visible = False

            '----------------Upper Inside Railing Elbow ----------------------

            Dim oUpperInsideElbowPlane As WorkPlane
            oUpperInsideElbowPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oLn22, oTopStepFrontPlane, 90 * pi / 180, False)
            Dim oUpperInsideElbowSketch As PlanarSketch
            oUpperInsideElbowSketch = oDef.Sketches.Add(oUpperInsideElbowPlane)
            Dim oWp07 As SketchPoints
            oWp07 = oUpperInsideElbowSketch.SketchPoints
            Call oWp07.Add(oTG.CreatePoint2d(-4 - 55, 15), False)
            Call oWp07.Add(oTG.CreatePoint2d(-4 - 55 - (Math.Tan(oAngle * ((oCSIntRadius + 1.2 + oCSWidth) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15), 30), False)
            Call oWp07.Add(oTG.CreatePoint2d(-4 - 105 - (Math.Tan(oAngle * ((oCSIntRadius + 1.2 + oCSWidth) / (oCSIntRadius + 0.8 + oCSWidth / 2))) * 15), 30), False)
            Call oWp07.Add(oTG.CreatePoint2d(-4 - 105, 15), False)
            Dim oLn53 As SketchLine
            oLn53 = oUpperInsideElbowSketch.SketchLines.AddByTwoPoints(oWp07(1), oWp07(2))
            Dim oLn54 As SketchLine
            oLn54 = oUpperInsideElbowSketch.SketchLines.AddByTwoPoints(oWp07(2), oWp07(3))
            Dim oLn55 As SketchLine
            oLn55 = oUpperInsideElbowSketch.SketchLines.AddByTwoPoints(oWp07(3), oWp07(4))
            Dim oArc10 As SketchArc
            oArc10 = oUpperInsideElbowSketch.SketchArcs.AddByFillet(oLn53, oLn54, 7.5, oLn53.Geometry.MidPoint, oLn54.Geometry.MidPoint)
            Dim oArc11 As SketchArc
            oArc11 = oUpperInsideElbowSketch.SketchArcs.AddByFillet(oLn54, oLn55, 7.5, oLn54.Geometry.MidPoint, oLn55.Geometry.MidPoint)
            Dim oHelpPln3 As WorkPlane
            oHelpPln3 = oDef.WorkPlanes.AddByPlaneAndOffset(oTopStepFrontPlane, 15, False)
            Dim oHelpSketch2 As PlanarSketch
            oHelpSketch2 = oDef.Sketches.Add(oHelpPln3)
            Dim oHelpPnt2 As SketchPoints
            oHelpPnt2 = oHelpSketch2.SketchPoints
            Call oHelpPnt2.Add(oTG.CreatePoint2d(oCSHeight / 10 + 105, 0), False)
            Call oHelpPnt2.Add(oTG.CreatePoint2d(oCSHeight / 10 + 105, oCSIntRadius / 10), False)
            Dim oHelpLn2 As SketchLine
            oHelpLn2 = oHelpSketch2.SketchLines.AddByTwoPoints(oHelpPnt2(1), oHelpPnt2(2))
            Dim oHelpPln4 As WorkPlane
            oHelpPln4 = oDef.WorkPlanes.AddByLinePlaneAndAngle(oHelpLn2, oHelpPln3, oAngle, False)
            Dim oUpInRailElbSketch As PlanarSketch
            oUpInRailElbSketch = oDef.Sketches.Add(oHelpPln4)
            Dim oCirc11 As SketchCircle
            oCirc11 = oUpInRailElbSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(-oCSIntRadius / 10 - 0.4, 0), 4.21 / 2)
            Dim oPath4 As Path
            oPath4 = oDef.Features.CreatePath(oLn54)
            Dim oRailProfile4 As Profile
            oRailProfile4 = oUpInRailElbSketch.Profiles.AddForSolid
            Dim oSweep4 As SweepFeature
            oSweep4 = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile4, oPath4, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)


            oUpperInsideElbowPlane.Visible = False
            oHelpPln3.Visible = False
            oHelpPln4.Visible = False
            oHelpSketch2.Visible = False
            Try
                RS = oStairsDoc.RenderStyles.Item("StairsRailingStyle")
            Catch ex As Exception
                RS = oStairsDoc.RenderStyles.Add("StairsRailingStyle")
                RS.Reflectivity = 3
                RS.SetDiffuseColor(255, 255, 0)  'Yellow
            End Try

            Call oSweep3.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSweep4.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oInnerSpilPattern.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpiralLowOutRail2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpiralUpOutRail2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSpil2Extrude.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
CWend:

        End If

        oInvApp.ScreenUpdating = True
        ProgressBar5.Value = 100
        oInvApp.ScreenUpdating = False



        oInvApp.CommandManager.ControlDefinitions.Item("AppViewCubeHomeCmd").Execute()
        oInvApp.ActiveView.GoHome()
        oStairsDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits



        Try
            oStairsDoc.SaveAs(lblCurvedStairsPath.Text & oCSFilename, False)
        Catch ex As Exception
            MsgBox("Could not save the part." & vbCrLf & "Try again or check your credentials", vbOKOnly + "4064", "Warning")
            Me.oInvApp.ScreenUpdating = True
            GoTo EndRoutine
        End Try

        Me.oInvApp.ScreenUpdating = True
        ProgressBar5.Value = 100
        Me.oInvApp.ScreenUpdating = False

        oStairsDoc.Close()
        Me.oInvApp.ScreenUpdating = True

        '--------------Place the part interactive--------------------------------------
PlacePart:
        ProgressBar5.Value = 0
        ' Get the active assembly.

        asmDoc = Me.oInvApp.ActiveDocument

        ' If assembly empty place stairs grounded at origin.

        If asmDoc.ComponentDefinition.ImmediateReferencedDefinitions.Count < 1 Then
            Dim trans As Matrix
            trans = Me.oInvApp.TransientGeometry.CreateMatrix
            stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblCurvedStairsPath.Text & oCSFilename), trans)
            stairsOcc.Grounded = True
            GoTo EndRoutine
        End If

        'Else place stairs on a floor surface.

        Hide()
        oFloorFace = Me.oInvApp.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, "Select a floor to place the stairs.")
        If oFloorFace Is Nothing Then
            GoTo EndRoutine
        End If

        stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblCurvedStairsPath.Text & oCSFilename), Me.oInvApp.TransientGeometry.CreateMatrix)
        oStairsDoc = stairsOcc.Definition.Document

        Try
            attrSets = oStairsDoc.AttributeManager.FindAttributeSets("ConnectFace")
            Dim oStairsCon As Face
            oStairsCon = attrSets.Item(1).Parent.Parent
            Dim oStairsFaceProxy As FaceProxy
            oStairsFaceProxy = Nothing
            Call stairsOcc.CreateGeometryProxy(oStairsCon, oStairsFaceProxy)
            Call asmDoc.ComponentDefinition.Constraints.AddMateConstraint(oFloorFace, oStairsFaceProxy, 0)
        Catch ex As Exception

        End Try

EndRoutine:
        Me.Show()
    End Sub


    '-----------------------------------------------------------------------------------------------------------------------------------------
    '-----------------------------------------------------------------------------------------------------------------------------------------

    '                        PLACING A LADDER IN AN ASSEMBLY

    '-----------------------------------------------------------------------------------------------------------------------------------------
    '-----------------------------------------------------------------------------------------------------------------------------------------



    Private Sub btnPlaceLadder_Click(sender As Object, e As EventArgs) Handles btnPlaceLadder.Click
        PlaceLadderInAssembly()
    End Sub

    Private Sub PlaceLadderInAssembly()



        Dim oDef As PartComponentDefinition
        Dim oTG As TransientGeometry
        Dim asmDoc As AssemblyDocument
        Dim oFloorFace As Face
        Dim oStairsDoc As PartDocument
        Dim attrSets As AttributeSetsEnumerator
        Dim stairsOcc As ComponentOccurrence
        Dim propSet1 As PropertySet
        Dim propSet2 As PropertySet
        Dim propSet3 As PropertySet

        'Check of Inventor nog steeds actief is
        Try
            Me.oInvApp = GetObject(, "Inventor.Application")
        Catch ex As Exception
            MsgBox("Inventor must be running!.", vbExclamation, "")
            btnPlaceStairsInAssem.Enabled = False
            btnPlaceLadder.Enabled = False
            btnPlaceRailing.Enabled = False
            btnPlacePlatform.Enabled = False
            btnPlaceCurvedStairs.Enabled = False
            Exit Sub
        End Try

        'Check of er een assembly open is
        If Me.oInvApp.ActiveDocumentType = Global.Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            GoTo BeginRoutine
        Else MsgBox("Open, Create or Activate an Assembly !", vbExclamation, "")
            GoTo EndRoutine
        End If

BeginRoutine:
        If Dir(lblLadderPath.Text & oLadderFilename) = "" Then
            GoTo StartPart
        Else
            GoTo PlacePart
        End If
StartPart:

        oInvApp = GetObject(, "Inventor.Application")




        '-----------------------------------------------------------------------------------------

        Try
            oInvApp = GetObject(, "Inventor.Application")
        Catch ex As Exception
            MsgBox("Inventor must be running.")
            btnPlaceStairsInAssem.Enabled = False
            btnPlaceLadder.Enabled = False
            btnPlaceRailing.Enabled = False
            btnPlacePlatform.Enabled = False
            btnPlaceCurvedStairs.Enabled = False
            Exit Sub
        End Try
        Dim invApp As Inventor.Application
        Dim oLadderMidPlane As WorkPlane
        Dim oCenterPlane As WorkPlane

        oInvApp.ScreenUpdating = False

        invApp = GetObject(, "Inventor.Application")
        oStairsDoc = invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject)
        oStairsDoc = invApp.ActiveDocument
        oDef = invApp.ActiveDocument.ComponentDefinition
        oTG = invApp.TransientGeometry
        pi = Math.Acos(-1) 'radians van 45° hoek
        oAngle = (oStairsAngle * pi / 180)
        oNumberOfRungs = oLadderHeight / 280
        oRungheight = oLadderHeight / oNumberOfRungs + 1
        oStairsHypotenusa = (oStairsHeight - oStepheight) / (Math.Sin(oAngle))
        oNumberOfHoops = Math.Round(((oLadderHeight - 2150) / 1200), 0)

        oInvApp.ScreenUpdating = True
        ProgressBar2.Value = 10
        oInvApp.ScreenUpdating = False



        '-----------------------------------------------------------------------------Setting Properties------------------------------------------------- 
        ' Get the design tracking property set.

        propSet1 = oStairsDoc.PropertySets.Item("Inventor Summary Information")
        propSet1.ItemByPropId(2).Value = "Steel Ladder"
        propSet1.ItemByPropId(4).Value = "Marc Crauwels"
        propSet1.ItemByPropId(6).Value = "This Steel Ladder is generated with the Pantarein Water Panta Stairs application. This software is part of the Pantarein Water Design Suite."

        propSet2 = oStairsDoc.PropertySets.Item("Inventor Document Summary Information")
        propSet2.ItemByPropId(2).Value = "Plant Design Mechanical"
        propSet2.ItemByPropId(15).Value = "Pantarein Water BVBA"

        propSet3 = oStairsDoc.PropertySets.Item("Design Tracking Properties")
        propSet3.ItemByPropId(41).Value = "Marc Crauwels"
        propSet3.ItemByPropId(29).Value = "Steel Ladder"
        propSet3.ItemByPropId(23).Value = "http://www.pantareinwater.be/en"




        '-------------------------------------------------------------------------------------------------------------------------------------------------
        '----------Adjust Orientation XY is floor Z is up----------

        Dim v As Inventor.View
        v = invApp.ActiveView
        Dim c As Camera
        c = invApp.ActiveView.Camera
        Dim t2eDist As Double
        t2eDist = c.Target.DistanceTo(c.Eye)
        Dim t2e As Vector
        t2e = oTG.CreateVector(0, -t2eDist, 0)
        Dim newEye As Point
        newEye = c.Target.Copy
        Call newEye.TranslateBy(t2e)
        c.Eye = newEye
        c.UpVector = oTG.CreateUnitVector(0, 0, 1)
        c.ApplyWithoutTransition()
        v.SetCurrentAsFront()
        '-------------------------------------------------------------------------------------------------------------------------------------------------
        '-----------------Ladder Mid Plane---------------------

        oCenterPlane = oDef.WorkPlanes.Item(2)
        oLadderMidPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oCenterPlane, -20, False)
        'oLadderMidPlane.Name = "Beam Middle Plane"


        '----------------Sketch Left Side----------------------


        Dim oLsideSketch As PlanarSketch
        oLsideSketch = oDef.Sketches.Add(oLadderMidPlane)

        Dim oWp02 As SketchPoints
        oWp02 = oLsideSketch.SketchPoints

        Call oWp02.Add(oTG.CreatePoint2d(22.5, 5), False)
        Call oWp02.Add(oTG.CreatePoint2d(22.5, oLadderHeight / 10 - 15), False)
        Call oWp02.Add(oTG.CreatePoint2d(22.5 + 15, oLadderHeight / 10), False)
        Call oWp02.Add(oTG.CreatePoint2d(22.5 + 15, oLadderHeight / 10 + 120), False)
        Call oWp02.Add(oTG.CreatePoint2d(22.5 + 19.5, oLadderHeight / 10 + 120), False)
        Call oWp02.Add(oTG.CreatePoint2d(22.5 + 15 + 4.5, (oLadderHeight / 10) - (Math.Tan(45 / 180 * pi / 2) * 4.5)), False)
        Call oWp02.Add(oTG.CreatePoint2d(22.5 + 4.5, (oLadderHeight / 10) - 15 - (Math.Tan(45 / 180 * pi / 2) * 4.5)), False)
        Call oWp02.Add(oTG.CreatePoint2d(22.5 + 4.5, 5), False)

        Dim oLn07 As SketchLine
        oLn07 = oLsideSketch.SketchLines.AddByTwoPoints(oWp02(1), oWp02(2))
        Dim oLn08 As SketchLine
        oLn08 = oLsideSketch.SketchLines.AddByTwoPoints(oWp02(2), oWp02(3))
        Dim oLn09 As SketchLine
        oLn09 = oLsideSketch.SketchLines.AddByTwoPoints(oWp02(3), oWp02(4))
        Dim oLn10 As SketchLine
        oLn10 = oLsideSketch.SketchLines.AddByTwoPoints(oWp02(4), oWp02(5))
        Dim oLn11 As SketchLine
        oLn11 = oLsideSketch.SketchLines.AddByTwoPoints(oWp02(5), oWp02(6))
        Dim oLn12 As SketchLine
        oLn12 = oLsideSketch.SketchLines.AddByTwoPoints(oWp02(6), oWp02(7))
        Dim oLn13 As SketchLine
        oLn13 = oLsideSketch.SketchLines.AddByTwoPoints(oWp02(7), oWp02(8))
        Dim oLn14 As SketchLine
        oLn14 = oLsideSketch.SketchLines.AddByTwoPoints(oWp02(8), oWp02(1))


        oInvApp.ScreenUpdating = True
        ProgressBar2.Value = 20
        oInvApp.ScreenUpdating = False


        Dim oProfile2 As Profile
        oProfile2 = oLsideSketch.Profiles.AddForSolid
        Dim oExtrudeDef2 As ExtrudeDefinition
        Dim kJoinOperation As Object = Nothing
        oExtrudeDef2 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile2, PartFeatureOperationEnum.kJoinOperation)
        Call oExtrudeDef2.SetDistanceExtent(8, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
        Dim oExtrude2 As ExtrudeFeature
        oExtrude2 = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef2)
        ' oExtrude2.Name = "Left Side Ladder Beam"
        '--------------------------------------------------------------------------







        '----------------Sketch Left Side Cut----------------------


        Dim oLsideCutSketch As PlanarSketch
        oLsideCutSketch = oDef.Sketches.Add(oLadderMidPlane)

        Dim oWp03 As SketchPoints
        oWp03 = oLsideCutSketch.SketchPoints

        Call oWp03.Add(oTG.CreatePoint2d(22.5 + 0.6, 5), False)
        Call oWp03.Add(oTG.CreatePoint2d(22.5 + 0.6, oLadderHeight / 10 - 15), False)
        Call oWp03.Add(oTG.CreatePoint2d(22.5 + 15 + 0.6, oLadderHeight / 10 - (Math.Tan(45 / 180 * pi / 2) * 0.6)), False)
        Call oWp03.Add(oTG.CreatePoint2d(22.5 + 15 + 0.6, oLadderHeight + 1200 / 10 - (Math.Tan(45 / 180 * pi / 2) * 0.6)), False)
        Call oWp03.Add(oTG.CreatePoint2d(22.5 + 19.5, oLadderHeight / 10 + 120), False)
        Call oWp03.Add(oTG.CreatePoint2d(22.5 + 15 + 4.5, (oLadderHeight / 10) - (Math.Tan(45 / 180 * pi / 2) * 4.5)), False)
        Call oWp03.Add(oTG.CreatePoint2d(22.5 + 4.5, (oLadderHeight / 10) - 15 - (Math.Tan(45 / 180 * pi / 2) * 4.5)), False)
        Call oWp03.Add(oTG.CreatePoint2d(22.5 + 4.5, 5), False)

        Dim oLn15 As SketchLine
        oLn15 = oLsideCutSketch.SketchLines.AddByTwoPoints(oWp03(1), oWp03(2))
        Dim oLn16 As SketchLine
        oLn16 = oLsideCutSketch.SketchLines.AddByTwoPoints(oWp03(2), oWp03(3))
        Dim oLn17 As SketchLine
        oLn17 = oLsideCutSketch.SketchLines.AddByTwoPoints(oWp03(3), oWp03(4))
        Dim oLn18 As SketchLine
        oLn18 = oLsideCutSketch.SketchLines.AddByTwoPoints(oWp03(4), oWp03(5))
        Dim oLn19 As SketchLine
        oLn19 = oLsideCutSketch.SketchLines.AddByTwoPoints(oWp03(5), oWp03(6))
        Dim oLn20 As SketchLine
        oLn20 = oLsideCutSketch.SketchLines.AddByTwoPoints(oWp03(6), oWp03(7))
        Dim oLn21 As SketchLine
        oLn21 = oLsideCutSketch.SketchLines.AddByTwoPoints(oWp03(7), oWp03(8))
        Dim oLn22 As SketchLine
        oLn22 = oLsideCutSketch.SketchLines.AddByTwoPoints(oWp03(8), oWp03(1))
        ProgressBar2.Value = 30

        Dim oProfile3 As Profile
        oProfile3 = oLsideCutSketch.Profiles.AddForSolid
        Dim oExtrudeDef3 As ExtrudeDefinition
        oExtrudeDef3 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile3, PartFeatureOperationEnum.kCutOperation)
        Call oExtrudeDef3.SetDistanceExtent(6.4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
        Dim oExtrude3 As ExtrudeFeature
        oExtrude3 = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef3)
        'oExtrude3.Name = "Left Side Ladder Beam Cut"
        '--------------------------------------------------------------------------

        '----------------Sketch Feet----------------------

        Dim oFeetSketch As PlanarSketch
        oFeetSketch = oDef.Sketches.Add(oLadderMidPlane)

        Dim oWp01 As SketchPoints
        oWp01 = oFeetSketch.SketchPoints

        oInvApp.ScreenUpdating = True
        ProgressBar2.Value = 40
        oInvApp.ScreenUpdating = False

        '------------Sketch points Feet----------------------

        Call oWp01.Add(oTG.CreatePoint2d(22.5 + 0.6, 0), False)
        Call oWp01.Add(oTG.CreatePoint2d(27.5 + 0.6, 0), False)
        Call oWp01.Add(oTG.CreatePoint2d(27.5 + 0.6, 1), False)
        Call oWp01.Add(oTG.CreatePoint2d(23.5 + 0.6, 1), False)
        Call oWp01.Add(oTG.CreatePoint2d(23.5 + 0.6, 10), False)
        Call oWp01.Add(oTG.CreatePoint2d(22.5 + 0.6, 10), False)

        '-------------Draw lines----------------------------------------------------

        Dim oLn01 As SketchLine
        oLn01 = oFeetSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
        Dim oLn02 As SketchLine
        oLn02 = oFeetSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
        Dim oLn03 As SketchLine
        oLn03 = oFeetSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
        Dim oLn04 As SketchLine
        oLn04 = oFeetSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(5))
        Dim oLn05 As SketchLine
        oLn05 = oFeetSketch.SketchLines.AddByTwoPoints(oWp01(5), oWp01(6))
        Dim oLn06 As SketchLine
        oLn06 = oFeetSketch.SketchLines.AddByTwoPoints(oWp01(6), oWp01(1))

        '--------------Extrusion of the Feet--------------------------------------


        Dim oProfile As Profile
        oProfile = oFeetSketch.Profiles.AddForSolid
        Dim oExtrudeDef As ExtrudeDefinition
        'Dim kJoinOperation As Object = Nothing
        oExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kJoinOperation)
        Call oExtrudeDef.SetDistanceExtent(5, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
        Dim oExtrude As ExtrudeFeature
        oExtrude = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)

        'start---------Maak een AttributeSet "ConnectFace" ----------

        Dim entity As Face
        entity = oDef.SurfaceBodies.Item(1).Faces.Item(7)
        Dim attribSets As AttributeSets
        attribSets = entity.AttributeSets
        Dim attribSet As AttributeSet
        attribSet = attribSets.Add("ConnectFace")

        'end---------Maak een AttributeSet "ConnectFace" ----------



        'oExtrude.Name = "Left Feet"
        '--------------------------------------------------------------------------

        '----------------Sketch Rung----------------------

        Dim oMiddlePlane As WorkPlane
        oMiddlePlane = oDef.WorkPlanes.Item(1)

        Dim oRungSketch As PlanarSketch
        oRungSketch = oDef.Sketches.Add(oMiddlePlane)


        Dim oRungCircle As SketchCircle
        oRungCircle = oRungSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(-20, oRungheight / 10), 1)



        Dim oProfile4 As Profile
        oProfile4 = oRungSketch.Profiles.AddForSolid
        Dim oExtrudeDef4 As ExtrudeDefinition
        oExtrudeDef4 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile4, PartFeatureOperationEnum.kNewBodyOperation)
        Call oExtrudeDef4.SetDistanceExtent(45, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
        Dim oExtrude4 As ExtrudeFeature
        oExtrude4 = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef4)
        'oExtrude4.Name = "Rung"

        '-------------------------------------Pattern Rungs-------------------------------------------------
        Dim oRPatCol As ObjectCollection
        Dim CreateObjectCollectionR As ObjectCollection = Nothing

        oRPatCol = (oExtrude4.Definition.AffectedBodies)


        Dim oRPattern As RectangularPatternFeature
        Dim kDefault As PatternSpacingTypeEnum = Nothing
        oRPattern = oDef.Features.RectangularPatternFeatures.Add(oRPatCol, oDef.WorkPlanes.Item(3), True, oNumberOfRungs - 1, oRungheight / 10, kDefault)
        'oRPattern.Name = "Rungs"

        '-----------------------------------Wall Supports---------------------------------------------------

        Dim oSupportPlane As WorkPlane
        oSupportPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oLadderHeight / 10 - 15, False)
        'oSupportPlane.Name = "Wall Support Top Plane"

        Dim oSupportSketch As PlanarSketch
        oSupportSketch = oDef.Sketches.Add(oSupportPlane)

        Dim oWp04 As SketchPoints
        oWp04 = oSupportSketch.SketchPoints

        oInvApp.ScreenUpdating = True
        ProgressBar2.Value = 50
        oInvApp.ScreenUpdating = False

        '------------Sketch points Support----------------------

        Call oWp04.Add(oTG.CreatePoint2d(-22.5, 0), False)
        Call oWp04.Add(oTG.CreatePoint2d(-22.5, -16), False)
        Call oWp04.Add(oTG.CreatePoint2d(-23.5, -16), False)
        Call oWp04.Add(oTG.CreatePoint2d(-23.5, -1), False)
        Call oWp04.Add(oTG.CreatePoint2d(-30.5, -1), False)
        Call oWp04.Add(oTG.CreatePoint2d(-30.5, 0), False)

        '-------------Draw Support lines----------------------------------------------------

        Dim oLn23 As SketchLine
        oLn23 = oSupportSketch.SketchLines.AddByTwoPoints(oWp04(1), oWp04(2))
        Dim oLn24 As SketchLine
        oLn24 = oSupportSketch.SketchLines.AddByTwoPoints(oWp04(2), oWp04(3))
        Dim oLn25 As SketchLine
        oLn25 = oSupportSketch.SketchLines.AddByTwoPoints(oWp04(3), oWp04(4))
        Dim oLn26 As SketchLine
        oLn26 = oSupportSketch.SketchLines.AddByTwoPoints(oWp04(4), oWp04(5))
        Dim oLn27 As SketchLine
        oLn27 = oSupportSketch.SketchLines.AddByTwoPoints(oWp04(5), oWp04(6))
        Dim oLn28 As SketchLine
        oLn28 = oSupportSketch.SketchLines.AddByTwoPoints(oWp04(6), oWp04(1))

        Dim oProfile5 As Profile
        oProfile5 = oSupportSketch.Profiles.AddForSolid
        Dim oExtrudeDef5 As ExtrudeDefinition
        oExtrudeDef5 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile5, PartFeatureOperationEnum.kNewBodyOperation)
        Call oExtrudeDef5.SetDistanceExtent(8, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        Dim oExtrude5 As ExtrudeFeature
        oExtrude5 = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef5)
        'oExtrude5.Name = "Top Left Wall Support"

        '----------------------------------pattern Wall Supports--------------------------------------------

        Dim oSPatCol As ObjectCollection
        Dim CreateObjectCollectionS As ObjectCollection = Nothing
        Dim oNumberOfSupports As Integer
        oSPatCol = (oExtrude5.Definition.AffectedBodies)
        oNumberOfSupports = (oLadderHeight - 150 - 500) / 1500


        Dim oSPattern As RectangularPatternFeature
        kDefault = Nothing
        oSPattern = oDef.Features.RectangularPatternFeatures.Add(oSPatCol, oDef.WorkPlanes.Item(3), False, oNumberOfSupports, 150, kDefault)
        'oSPattern.Name = "Left Wall Supports"

        '----------------------------------mirror Ladder beam-----------------------------------------------

        Dim oLadCol As ObjectCollection
        Dim CreateObjectCollectionL As ObjectCollection = Nothing

        'oLadCol = (oExtrude.Definition.AffectedBodies)
        'oLadCol = (oExtrude2.Definition.AffectedBodies)
        oLadCol = (oExtrude3.Definition.AffectedBodies)
        'oLadCol = (oExtrude5.Definition.AffectedBodies)
        'oLadCol = (oSPattern.Definition.AffectedBodies)
        'oCol.Add(oDoc.ComponentDefinition.SurfaceBodies(1))

        Dim oPlane As WorkPlane
        oPlane = oStairsDoc.ComponentDefinition.WorkPlanes(1)
        Dim oMirror As MirrorFeature
        oMirror = oStairsDoc.ComponentDefinition.Features.MirrorFeatures.Add(oLadCol, oMiddlePlane, False, PatternComputeTypeEnum.kOptimizedCompute)
        'oMirror.Name = "Right Side"


        '----------------------------------mirror Ladder Support-----------------------------------------------

        Dim oSupCol As ObjectCollection
        Dim CreateObjectCollectionSup As ObjectCollection = Nothing

        'oLadCol = (oExtrude.Definition.AffectedBodies)
        'oLadCol = (oExtrude2.Definition.AffectedBodies)
        'oLadCol = (oExtrude3.Definition.AffectedBodies)
        oSupCol = (oExtrude5.Definition.AffectedBodies)
        'oSupCol = (oSPattern.Definition.AffectedBodies)
        'oCol.Add(oDoc.ComponentDefinition.SurfaceBodies(1))

        oPlane = oStairsDoc.ComponentDefinition.WorkPlanes(1)
        Dim oMirSup As MirrorFeature
        oMirSup = oStairsDoc.ComponentDefinition.Features.MirrorFeatures.Add(oSupCol, oMiddlePlane, False, PatternComputeTypeEnum.kOptimizedCompute)
        'oMirSup.Name = "Right Side Supports"


        If oCageState = "NoCage" Then
            GoTo EndNoCage
        End If

        '-------------------CAGE-------------------------------------------------------------------------------
        '-------------------CAGE-------------------------------------------------------------------------------
        '-------------------CAGE-------------------------------------------------------------------------------
        '-------------------CAGE-------------------------------------------------------------------------------


        '-----------------------------------Top Hoops---------------------------------------------------

        Dim oTHPlane As WorkPlane
        oTHPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oLadderHeight / 10 + 120, False)
        'oTHPlane.Name = "Top Hoop Plane"

        Dim oTHSketch As PlanarSketch
        oTHSketch = oDef.Sketches.Add(oTHPlane)

        Dim oWp05 As SketchPoints
        oWp05 = oTHSketch.SketchPoints

        oInvApp.ScreenUpdating = True
        ProgressBar2.Value = 60
        oInvApp.ScreenUpdating = False

        '------------Sketch points Top Hoops----------------------

        Call oWp05.Add(oTG.CreatePoint2d(-39, -24), False)
        Call oWp05.Add(oTG.CreatePoint2d(-39, -59), False)
        Call oWp05.Add(oTG.CreatePoint2d(0, -93), False)
        Call oWp05.Add(oTG.CreatePoint2d(39, -59), False)
        Call oWp05.Add(oTG.CreatePoint2d(39, -24), False)
        Call oWp05.Add(oTG.CreatePoint2d(40, -24), False)
        Call oWp05.Add(oTG.CreatePoint2d(40, -59), False)
        Call oWp05.Add(oTG.CreatePoint2d(0, -94), False)
        Call oWp05.Add(oTG.CreatePoint2d(-40, -59), False)
        Call oWp05.Add(oTG.CreatePoint2d(-40, -24), False)


        '-------------Draw Top Hoops lines----------------------------------------------------

        Dim oLn29 As SketchLine
        oLn29 = oTHSketch.SketchLines.AddByTwoPoints(oWp05(1), oWp05(2))
        Dim oLn30 As SketchLine
        oLn30 = oTHSketch.SketchLines.AddByTwoPoints(oWp05(4), oWp05(5))
        Dim oLn31 As SketchLine
        oLn31 = oTHSketch.SketchLines.AddByTwoPoints(oWp05(5), oWp05(6))
        Dim oLn32 As SketchLine
        oLn32 = oTHSketch.SketchLines.AddByTwoPoints(oWp05(6), oWp05(7))
        Dim oLn33 As SketchLine
        oLn33 = oTHSketch.SketchLines.AddByTwoPoints(oWp05(9), oWp05(10))
        Dim oLn34 As SketchLine
        oLn34 = oTHSketch.SketchLines.AddByTwoPoints(oWp05(10), oWp05(1))

        Dim oAr01 As SketchArc
        oAr01 = oTHSketch.SketchArcs.AddByThreePoints(oLn30.StartSketchPoint, (oTG.CreatePoint2d(0, -98)), oLn29.EndSketchPoint)
        Dim oAr02 As SketchArc
        oAr02 = oTHSketch.SketchArcs.AddByThreePoints(oLn32.EndSketchPoint, (oTG.CreatePoint2d(0, -99)), oLn33.StartSketchPoint)

        Dim oProfile6 As Profile
        oProfile6 = oTHSketch.Profiles.AddForSolid
        Dim oExtrudeDef6 As ExtrudeDefinition
        oExtrudeDef6 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile6, PartFeatureOperationEnum.kNewBodyOperation)
        Call oExtrudeDef6.SetDistanceExtent(5, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        Dim oExtrude6 As ExtrudeFeature
        oExtrude6 = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef6)
        ' oExtrude6.Name = "Top Hoop"

        '----------------------------------pattern Top Hoops--------------------------------------------

        Dim oTHCol As ObjectCollection
        Dim CreateObjectCollectionTH As ObjectCollection = Nothing
        oTHCol = (oExtrude6.Definition.AffectedBodies)
        Dim oTHPattern As RectangularPatternFeature
        kDefault = Nothing
        oTHPattern = oDef.Features.RectangularPatternFeatures.Add(oTHCol, oDef.WorkPlanes.Item(3), False, 2, 115, kDefault)
        'oTHPattern.Name = "Top Hoops"

        '-----------------------------------Lower Hoops---------------------------------------------------

        Dim oLHPlane As WorkPlane
        oLHPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), 215, False)
        ' oLHPlane.Name = "Lower Hoop Plane"

        Dim oLHSketch As PlanarSketch
        oLHSketch = oDef.Sketches.Add(oLHPlane)

        Dim oWp06 As SketchPoints
        oWp06 = oLHSketch.SketchPoints

        oInvApp.ScreenUpdating = True
        ProgressBar2.Value = 70
        oInvApp.ScreenUpdating = False

        '------------Sketch points Lower Hoops----------------------

        Call oWp06.Add(oTG.CreatePoint2d(-39 + 15, -24), False)
        Call oWp06.Add(oTG.CreatePoint2d(-39, -59), False)
        Call oWp06.Add(oTG.CreatePoint2d(0, -98), False)
        Call oWp06.Add(oTG.CreatePoint2d(39, -59), False)
        Call oWp06.Add(oTG.CreatePoint2d(39 - 15, -24), False)
        Call oWp06.Add(oTG.CreatePoint2d(40 - 15, -24), False)
        Call oWp06.Add(oTG.CreatePoint2d(40, -59), False)
        Call oWp06.Add(oTG.CreatePoint2d(0, -99), False)
        Call oWp06.Add(oTG.CreatePoint2d(-40, -59), False)
        Call oWp06.Add(oTG.CreatePoint2d(-40 + 15, -24), False)


        '-------------Draw Lower Hoops lines----------------------------------------------------

        Dim oLn35 As SketchLine
        oLn35 = oLHSketch.SketchLines.AddByTwoPoints(oWp06(1), oWp06(2))
        Dim oLn36 As SketchLine
        oLn36 = oLHSketch.SketchLines.AddByTwoPoints(oWp06(4), oWp06(5))
        Dim oLn37 As SketchLine
        oLn37 = oLHSketch.SketchLines.AddByTwoPoints(oWp06(5), oWp06(6))
        Dim oLn38 As SketchLine
        oLn38 = oLHSketch.SketchLines.AddByTwoPoints(oWp06(6), oWp06(7))
        Dim oLn39 As SketchLine
        oLn39 = oLHSketch.SketchLines.AddByTwoPoints(oWp06(9), oWp06(10))
        Dim oLn40 As SketchLine
        oLn40 = oLHSketch.SketchLines.AddByTwoPoints(oWp06(10), oWp06(1))

        Dim oAr03 As SketchArc
        oAr03 = oLHSketch.SketchArcs.AddByThreePoints(oLn36.StartSketchPoint, (oTG.CreatePoint2d(0, -98)), oLn35.EndSketchPoint)
        Dim oAr04 As SketchArc
        oAr04 = oLHSketch.SketchArcs.AddByThreePoints(oLn38.EndSketchPoint, (oTG.CreatePoint2d(0, -99)), oLn39.StartSketchPoint)

        Call oLHSketch.GeometricConstraints.AddGround(oLn40.EndSketchPoint)
        Call oLHSketch.GeometricConstraints.AddGround(oLn40.StartSketchPoint)
        Call oLHSketch.GeometricConstraints.AddGround(oLn37.EndSketchPoint)
        Call oLHSketch.GeometricConstraints.AddGround(oLn37.StartSketchPoint)
        Call oLHSketch.GeometricConstraints.AddGround(oAr03)
        Call oLHSketch.GeometricConstraints.AddGround(oAr04)
        Call oLHSketch.GeometricConstraints.AddTangent(oAr03, oLn36)
        Call oLHSketch.GeometricConstraints.AddTangent(oAr03, oLn35)
        Call oLHSketch.GeometricConstraints.AddTangent(oAr04, oLn38)
        Call oLHSketch.GeometricConstraints.AddTangent(oAr04, oLn39)

        Dim oProfile7 As Profile
        oProfile7 = oLHSketch.Profiles.AddForSolid
        Dim oExtrudeDef7 As ExtrudeDefinition
        oExtrudeDef7 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile7, PartFeatureOperationEnum.kNewBodyOperation)
        Call oExtrudeDef7.SetDistanceExtent(5, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
        Dim oExtrude7 As ExtrudeFeature
        oExtrude7 = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef7)
        'oExtrude7.Name = "Lower Hoop"





        '----------------------------------pattern Lower Hoops--------------------------------------------

        Dim oLHCol As ObjectCollection
        Dim CreateObjectCollectionLH As ObjectCollection = Nothing
        oLHCol = (oExtrude7.Definition.AffectedBodies)
        Dim oLHPattern As RectangularPatternFeature
        Dim oNumberLowerHoops As Integer
        Dim oDistLowerHoops As Double

        oNumberLowerHoops = (oLadderHeight - 2100) / 1000
        If oNumberLowerHoops = 0 Then
            oNumberLowerHoops = 1
        End If
        oDistLowerHoops = (oLadderHeight - 2100) / oNumberLowerHoops
        kDefault = Nothing
        oLHPattern = oDef.Features.RectangularPatternFeatures.Add(oLHCol, oDef.WorkPlanes.Item(3), True, oNumberLowerHoops, oDistLowerHoops / 10, kDefault)
        'oLHPattern.Name = "Lower Hoops"

        '----------------------------------Hoops Connection Lath--------------------------------------------

        Dim oCLSketch As PlanarSketch
        oCLSketch = oDef.Sketches.Add(oTHPlane)

        Dim oWp07 As SketchPoints
        oWp07 = oCLSketch.SketchPoints

        oInvApp.ScreenUpdating = True
        ProgressBar2.Value = 80
        oInvApp.ScreenUpdating = False

        '------------Sketch points Hoops Connection Lath----------------------

        Call oWp07.Add(oTG.CreatePoint2d(39.5, -56.5), False)
        Call oWp07.Add(oTG.CreatePoint2d(39.5, -61.5), False)
        Call oWp07.Add(oTG.CreatePoint2d(40.5, -61.5), False)
        Call oWp07.Add(oTG.CreatePoint2d(40.5, -56.5), False)

        Dim oLn41 As SketchLine
        oLn41 = oCLSketch.SketchLines.AddByTwoPoints(oWp07(1), oWp07(2))
        Dim oLn42 As SketchLine
        oLn42 = oCLSketch.SketchLines.AddByTwoPoints(oWp07(2), oWp07(3))
        Dim oLn43 As SketchLine
        oLn43 = oCLSketch.SketchLines.AddByTwoPoints(oWp07(3), oWp07(4))
        Dim oLn44 As SketchLine
        oLn44 = oCLSketch.SketchLines.AddByTwoPoints(oWp07(4), oWp07(1))

        Dim oProfile8 As Profile
        oProfile8 = oCLSketch.Profiles.AddForSolid
        Dim oExtrudeDef8 As ExtrudeDefinition
        oExtrudeDef8 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile8, PartFeatureOperationEnum.kNewBodyOperation)
        Call oExtrudeDef8.SetDistanceExtent((oLadderHeight + 1200 - 2150) / 10, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
        Dim oExtrude8 As ExtrudeFeature
        oExtrude8 = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef8)
        'oExtrude8.Name = "Hoop Connection Lath"

        '----------------------------------pattern Hoops Connection Lath--------------------------------------------

        Dim oHCLCol As ObjectCollection
        'Dim CreateObjectCollectionHCL As ObjectCollection = Nothing
        oHCLCol = (oExtrude8.Definition.AffectedBodies)
        Dim oHCLPattern As CircularPatternFeature
        kDefault = Nothing

        'Dim oAxSketch As PlanarSketch
        'oAxSketch = oDef.Sketches.Add(oMiddlePlane)

        'Dim oWp08 As SketchPoints
        'oWp08 = oAxSketch.SketchPoints

        'Call oWp08.Add(oTG.CreatePoint2d(-59, 0), False)
        'Call oWp08.Add(oTG.CreatePoint2d(-59, 10), False)

        'Dim oLn45 As SketchLine
        'oLn45 = oAxSketch.SketchLines.AddByTwoPoints(oWp08(1), oWp08(2))

        'Dim oAxPatHCL As WorkAxes
        'oAxPatHCL = oAxSketch.WorkAxes.AddByLine(oLn45, False)

        'create a workaxis passing between two Points
        Dim oPoint1 As Point
        Dim oPoint2 As Point

        oPoint1 = oTG.CreatePoint(0, -59, 0)
        oPoint2 = oTG.CreatePoint(0, -59, 10)

        Dim oUnitVec As UnitVector
        oUnitVec = oTG.CreateUnitVector((oPoint1.X - oPoint2.X), (oPoint1.Y - oPoint2.Y), (oPoint1.Z - oPoint2.Z))

        Dim oWorkAxis2 As WorkAxis
        oWorkAxis2 = oDef.WorkAxes.AddFixed(oPoint1, oUnitVec)

        oInvApp.ScreenUpdating = True
        ProgressBar2.Value = 90
        oInvApp.ScreenUpdating = False

        oHCLPattern = oDef.Features.CircularPatternFeatures.Add(oHCLCol, oWorkAxis2, True, 5, 180 / 180 * pi, True,)
        'oHCLPattern.Name = "Hoops Connection Laths"

        Dim RS As RenderStyle
        Try
            RS = oStairsDoc.RenderStyles.Item("LadderCageStyle")
        Catch ex As Exception
            RS = oStairsDoc.RenderStyles.Add("LadderCageStyle")
            RS.Reflectivity = 3
            RS.SetDiffuseColor(255, 255, 0)  'Yellow
        End Try
        'assign new render style to part feature
        'Dim oRailFeature As PartFeature = oExtrude
        Call oExtrude8.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
        Call oExtrude7.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
        Call oExtrude6.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
        Call oHCLPattern.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
        Call oTHPattern.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
        Call oLHPattern.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)


        oLadderMidPlane.Visible = False
        oSupportPlane.Visible = False
        oTHPlane.Visible = False
        oLHPlane.Visible = False
        oWorkAxis2.Visible = False

EndNoCage:
        oLadderMidPlane.Visible = False
        oSupportPlane.Visible = False


        'Hide()

        'Return view to Home view
        invApp.CommandManager.ControlDefinitions.Item("AppViewCubeHomeCmd").Execute()
        'zoom all
        invApp.ActiveView.GoHome()

        oInvApp.ScreenUpdating = True
        ProgressBar2.Value = 100
        oInvApp.ScreenUpdating = False

        oStairsDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits
        'Show()

        Try
            oStairsDoc.SaveAs(lblLadderPath.Text & oLadderFilename, False)
        Catch ex As Exception
            MsgBox("Could not save the part." & vbCrLf & "Try again or check your credentials", vbOKOnly + "4064", "Warning")
            Me.oInvApp.ScreenUpdating = True
            GoTo EndRoutine
        End Try



        oStairsDoc.Close()
        oInvApp.ScreenUpdating = True

        '--------------Place the part interactive--------------------------------------
PlacePart:
        ProgressBar2.Value = 0
        ' Get the active assembly.

        asmDoc = oInvApp.ActiveDocument

        ' If assembly empty place stairs grounded at origin.

        If asmDoc.ComponentDefinition.ImmediateReferencedDefinitions.Count < 1 Then
            Dim trans As Matrix
            trans = oInvApp.TransientGeometry.CreateMatrix
            stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblLadderPath.Text & oLadderFilename), trans)
            stairsOcc.Grounded = True
            GoTo EndRoutine
        End If

        'Else place stairs on a floor surface.

        Hide()
        oFloorFace = oInvApp.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, "Select a floor to place the stairs.")
        If oFloorFace Is Nothing Then
            GoTo EndRoutine
        End If

        stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblLadderPath.Text & oLadderFilename), oInvApp.TransientGeometry.CreateMatrix)
        oStairsDoc = stairsOcc.Definition.Document

        Try
            attrSets = oStairsDoc.AttributeManager.FindAttributeSets("ConnectFace")
            Dim oStairsCon As Face
            oStairsCon = attrSets.Item(1).Parent.Parent
            Dim oStairsFaceProxy As FaceProxy
            oStairsFaceProxy = Nothing
            Call stairsOcc.CreateGeometryProxy(oStairsCon, oStairsFaceProxy)
            Call asmDoc.ComponentDefinition.Constraints.AddMateConstraint(oFloorFace, oStairsFaceProxy, 0)
        Catch ex As Exception

        End Try

EndRoutine:
        Show()
    End Sub

    '-------------------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------


    '                  PLACING RAILING IN AN ASSEMBLY


    '-------------------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------


    Private Sub btnPlaceRailing_Click(sender As Object, e As EventArgs) Handles btnPlaceRailing.Click
        PlaceRailing()
    End Sub

    Private Sub PlaceRailing()

        Dim oDef As PartComponentDefinition
        Dim oTG As TransientGeometry
        Dim asmDoc As AssemblyDocument
        Dim oFloorFace As Face
        Dim oStairsDoc As PartDocument
        Dim attrSets As AttributeSetsEnumerator
        Dim stairsOcc As ComponentOccurrence
        Dim propSet1 As PropertySet
        Dim propSet2 As PropertySet
        Dim propSet3 As PropertySet

        'Check of Inventor nog steeds actief is
        Try
            Me.oInvApp = GetObject(, "Inventor.Application")
        Catch ex As Exception
            MsgBox("Inventor must be running!.", vbExclamation, "")
            btnPlaceStairsInAssem.Enabled = False
            btnPlaceLadder.Enabled = False
            btnPlaceRailing.Enabled = False
            btnPlacePlatform.Enabled = False
            btnPlaceCurvedStairs.Enabled = False
            Exit Sub
        End Try

        'Check of er een assembly open is
        If Me.oInvApp.ActiveDocumentType = Global.Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            GoTo BeginRoutine
        Else MsgBox("Open, Create or Activate an Assembly !", vbExclamation, "")
            GoTo EndRoutine
        End If

BeginRoutine:
        If Dir(lblRailPath.Text & oRailFilename) = "" Then
            GoTo StartPart
        Else
            GoTo PlacePart
        End If
StartPart:

        oInvApp = GetObject(, "Inventor.Application")

        ProgressBar3.Visible = True
        ProgressBar3.Value = 5
        oInvApp.ScreenUpdating = False




        Try
            oInvApp = GetObject(, "Inventor.Application")
        Catch ex As Exception
            MsgBox("Inventor must be running.")
            btnPlaceStairsInAssem.Enabled = False
            btnPlaceLadder.Enabled = False
            btnPlaceRailing.Enabled = False
            btnPlacePlatform.Enabled = False
            btnPlaceCurvedStairs.Enabled = False
            Exit Sub
        End Try



        oInvApp = GetObject(, "Inventor.Application")
        oStairsDoc = oInvApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject)
        oStairsDoc = oInvApp.ActiveDocument
        oDef = oInvApp.ActiveDocument.ComponentDefinition
        oTG = oInvApp.TransientGeometry
        pi = Math.Acos(-1) 'radians van 45° hoek





        '-----------------------------------------------------------------------------Setting Properties------------------------------------------------- 
        ' Get the design tracking property set.


        propSet1 = oStairsDoc.PropertySets.Item("Inventor Summary Information")
        propSet1.ItemByPropId(2).Value = "Steel Railing"
        propSet1.ItemByPropId(4).Value = "Marc Crauwels"
        propSet1.ItemByPropId(6).Value = "This Steel Railing is generated with the Pantarein Water Panta Stairs application. This software is part of the Pantarein Water Design Suite."

        propSet2 = oStairsDoc.PropertySets.Item("Inventor Document Summary Information")
        propSet2.ItemByPropId(2).Value = "Plant Design Mechanical"
        propSet2.ItemByPropId(15).Value = "Pantarein Water BVBA"

        propSet3 = oStairsDoc.PropertySets.Item("Design Tracking Properties")
        propSet3.ItemByPropId(41).Value = "Marc Crauwels"
        propSet3.ItemByPropId(29).Value = "Steel Railing"
        propSet3.ItemByPropId(23).Value = "http://www.pantareinwater.be/en"





        propSet1 = Nothing
        propSet2 = Nothing
        propSet3 = Nothing
        '-------------------------------------------------------------------------------------------------------------------------------------------------
        '----------Adjust Orientation XY is floor Z is up----------

        Dim v As Inventor.View
        v = oInvApp.ActiveView
        Dim c As Camera
        c = oInvApp.ActiveView.Camera
        Dim t2eDist As Double
        t2eDist = c.Target.DistanceTo(c.Eye)
        Dim t2e As Vector
        t2e = oTG.CreateVector(0, -t2eDist, 0)
        Dim newEye As Point
        newEye = c.Target.Copy
        Call newEye.TranslateBy(t2e)
        c.Eye = newEye
        c.UpVector = oTG.CreateUnitVector(0, 0, 1)
        c.ApplyWithoutTransition()
        v.SetCurrentAsFront()


        '---------------------1 SIDE ----------------------------------

        If oRailingState = "1Side" Then

            '----start----------------------Make the yellow color-------------------------------------------------------------------------------

            Dim RS As RenderStyle
            Try
                RS = oStairsDoc.RenderStyles.Item("RailingStyle")
            Catch ex As Exception
                RS = oStairsDoc.RenderStyles.Add("RailingStyle")
                RS.Reflectivity = 3
                RS.SetDiffuseColor(255, 255, 0)  'Yellow
            End Try

            '----end----------------------Make the yellow color-------------------------------------------------------------------------------

            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)
            'oGroundPlane.Name = "Bottom Of Railing"

            Dim oKickplateSketch As PlanarSketch
            oKickplateSketch = oDef.Sketches.Add(oGroundPlane)

            Dim oWp01 As SketchPoints
            oWp01 = oKickplateSketch.SketchPoints

            oInvApp.ScreenUpdating = True
            ProgressBar3.Value = 15
            oInvApp.ScreenUpdating = False

            '------------Sketch points Kickplate----------------------

            Call oWp01.Add(oTG.CreatePoint2d(-0.4, 0), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4, oRailLng1 / 10), False)
            Call oWp01.Add(oTG.CreatePoint2d(0.4, oRailLng1 / 10), False)
            Call oWp01.Add(oTG.CreatePoint2d(0.4, 0), False)

            '-------------Draw lines Kickplate----------------------------------------------------

            Dim oLn01 As SketchLine
            oLn01 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
            Dim oLn02 As SketchLine
            oLn02 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
            Dim oLn03 As SketchLine
            oLn03 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
            Dim oLn04 As SketchLine
            oLn04 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(1))

            ProgressBar3.Value = 50

            '--------------Extrusion of the Kickplate--------------------------------------


            Dim oKickplateProfile As Profile
            oKickplateProfile = oKickplateSketch.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oKickplateProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oExtrudeDef.SetDistanceExtent(15, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrude As ExtrudeFeature
            oExtrude = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)



            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(5)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------

            '--------------------------------------------------------------------------

            If oRailLng1 >= 300 And oRailLng1 < 600 Then
                GoTo noSpils
            End If



            Dim oRailSpilSketch1 As PlanarSketch
            oRailSpilSketch1 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle1 As SketchCircle
            oSpilCircle1 = oRailSpilSketch1.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 30), 3.37 / 2)

            oInvApp.ScreenUpdating = True
            ProgressBar3.Value = 50
            oInvApp.ScreenUpdating = False

            Dim oRailspilProfile1 As Profile
            oRailspilProfile1 = oRailSpilSketch1.Profiles.AddForSolid
            Dim oExtrDefSp01 As ExtrudeDefinition
            oExtrDefSp01 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile1, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp01.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp01 As ExtrudeFeature
            oExtrSp01 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp01)


            Try
                Call oExtrSp01.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try





            'oExtrSp01.Name = "Rail Spil 1"


            If oRailLng1 >= 600 And oRailLng1 < 900 Then
                GoTo oneSpilOnly
            End If


            '-------------------------------------Pattern Rail Spil 1-------------------------------------------------
            Dim oSpil1Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil1Col = (oExtrSp01.Definition.AffectedBodies)

            Dim oRPattern As RectangularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls01 As Integer
            Dim oDistSpls01 As Double

            oNbrSpls01 = Math.Ceiling((oRailLng1 - 600) / 900)

            If oRailLng1 <= 1600 Then
                oNbrSpls01 = 2
            End If

            oDistSpls01 = (oRailLng1 / 10 - 60) / (oNbrSpls01 - 1)


            oRPattern = oDef.Features.RectangularPatternFeatures.Add(oSpil1Col, oDef.WorkPlanes.Item(2), True, oNbrSpls01, oDistSpls01, kDefault)
            ' oRPattern.Name = "Rail Spil 1 Pattern"

            Try
                Call oRPattern.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try





noSpils:
oneSpilOnly:
            oInvApp.ScreenUpdating = True
            ProgressBar3.Value = 70
            oInvApp.ScreenUpdating = False


            'Add a new Sketch3D

            Dim oSketch3dRail01 As Sketch3D

            oSketch3dRail01 = oDef.Sketches3D.Add

            'Add two lines to the sketch

            Dim oLine1 As SketchLine3D
            Dim oLine2 As SketchLine3D
            Dim oLine3 As SketchLine3D
            Dim oLine4 As SketchLine3D



            oLine1 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oTG.CreatePoint(0, 2.1, 50), oTG.CreatePoint(0, oRailLng1 / 10 - 2.1, 50), True, 7.5)
            oLine2 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine1.EndPoint, oTG.CreatePoint(0, oRailLng1 / 10 - 2.1, 105), True, 7.5)
            oLine3 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine2.EndPoint, oTG.CreatePoint(0, 2.1, 105), True, 7.5)
            oLine4 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine3.EndPoint, oLine1.StartPoint, True, 7.5)

            Dim oRailingPath As Path
            oRailingPath = oDef.Features.CreatePath(oLine1)

            oGroundPlane = oDef.WorkPlanes.Item(3)
            Dim oRailDiaPlane As WorkPlane
            oRailDiaPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oGroundPlane, 75, False)

            Dim oRailingSketch2 As PlanarSketch
            oRailingSketch2 = oDef.Sketches.Add(oRailDiaPlane)
            Dim oRailDiaCircle As SketchCircle
            oRailDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 2.1), 2.1)
            Dim oRailIntDiaCircle As SketchCircle
            oRailIntDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 2.1), 1.8)
            Dim oRailProfile As Profile
            oRailProfile = oRailingSketch2.Profiles.AddForSolid



            Dim oSweep As SweepFeature
            oSweep = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile, oRailingPath, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)
            oGroundPlane.Visible = False
            oRailDiaPlane.Visible = False

            Try
                Call oExtrude.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

            Try
                Call oSweep.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

            oInvApp.ScreenUpdating = True
            ProgressBar3.Value = 100
            oInvApp.ScreenUpdating = False

        End If


        '--------------------------------------------------------------------------------------

        '-------------END  1 SIDE--------------------------------------------------------------

        '-------------START 2 SIDES------------------------------------------------------------

        '--------------------------------------------------------------------------------------




        If oRailingState = "2Sides" Then

            Dim RS As RenderStyle
            Try
                RS = oStairsDoc.RenderStyles.Item("RailingStyle")
            Catch ex As Exception
                RS = oStairsDoc.RenderStyles.Add("RailingStyle")
                RS.Reflectivity = 3
                RS.SetDiffuseColor(255, 255, 0)  'Yellow
            End Try

            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)
            ' oGroundPlane.Name = "Bottom Of Railing"

            Dim oKickplateSketch As PlanarSketch
            oKickplateSketch = oDef.Sketches.Add(oGroundPlane)

            Dim oWp01 As SketchPoints
            oWp01 = oKickplateSketch.SketchPoints


            ProgressBar3.Value = 40

            '------------Sketch points Kickplate----------------------

            Call oWp01.Add(oTG.CreatePoint2d(-0.4, 0), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4, oRailLng1 / 10), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + oRailLng2 / 10, oRailLng1 / 10), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + oRailLng2 / 10, oRailLng1 / 10 - 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(0.4, oRailLng1 / 10 - 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(0.4, 0), False)

            '-------------Draw lines Kickplate----------------------------------------------------

            Dim oLn01 As SketchLine
            oLn01 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
            Dim oLn02 As SketchLine
            oLn02 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
            Dim oLn03 As SketchLine
            oLn03 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
            Dim oLn04 As SketchLine
            oLn04 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(5))
            Dim oLn05 As SketchLine
            oLn05 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(5), oWp01(6))
            Dim oLn06 As SketchLine
            oLn06 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(6), oWp01(1))

            ProgressBar3.Value = 50

            '--------------Extrusion of the Kickplate--------------------------------------


            Dim oKickplateProfile As Profile
            oKickplateProfile = oKickplateSketch.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oKickplateProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oExtrudeDef.SetDistanceExtent(15, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrude As ExtrudeFeature
            oExtrude = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)


            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(7)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------



            ' oExtrude.Name = "Kickplate"
            '--------------------------------------------------------------------------

            If oRailLng1 >= 300 And oRailLng1 < 600 Then
                GoTo noSpils2
            End If

            '-------------------------------------Spils on side 1-----------------------------------

            Dim oRailSpilSketch1 As PlanarSketch
            oRailSpilSketch1 = oDef.Sketches.Add(oGroundPlane)
            Dim oSpilCircle1 As SketchCircle
            oSpilCircle1 = oRailSpilSketch1.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 30), 3.37 / 2)
            ProgressBar3.Value = 60
            Dim oRailspilProfile1 As Profile
            oRailspilProfile1 = oRailSpilSketch1.Profiles.AddForSolid
            Dim oExtrDefSp01 As ExtrudeDefinition
            oExtrDefSp01 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile1, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp01.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp01 As ExtrudeFeature
            oExtrSp01 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp01)
            'oExtrSp01.Name = "Rail Spil 1"
            If oRailLng1 >= 600 And oRailLng1 < 900 Then
                GoTo oneSpilOnly2
            End If

            '-------------------------------------Pattern Rail Spil 1-------------------------------------------------
            Dim oSpil1Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil1Col = (oExtrSp01.Definition.AffectedBodies)

            Dim oRPattern1 As RectangularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls01 As Integer
            Dim oDistSpls01 As Double

            oNbrSpls01 = Math.Ceiling((oRailLng1 - 600) / 900)
            If oRailLng1 <= 1600 Then
                oNbrSpls01 = 2
            End If
            oDistSpls01 = (oRailLng1 / 10 - 60) / (oNbrSpls01 - 1)
            oRPattern1 = oDef.Features.RectangularPatternFeatures.Add(oSpil1Col, oDef.WorkPlanes.Item(2), True, oNbrSpls01, oDistSpls01, kDefault)
            ' oRPattern1.Name = "Rail Spil 1 Pattern"
            Try
                Call oExtrSp01.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

            Try
                Call oRPattern1.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try




noSpils2:
oneSpilOnly2:

            If oRailLng2 >= 300 And oRailLng2 < 600 Then
                GoTo noSpils3
            End If
            '-------------------------------------Spils on side 2-----------------------------------

            Dim oRailSpilSketch2 As PlanarSketch
            oRailSpilSketch2 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle2 As SketchCircle
            oSpilCircle2 = oRailSpilSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(30, oRailLng1 / 10 - 0.4), 3.37 / 2)

            ProgressBar3.Value = 60

            Dim oRailspilProfile2 As Profile
            oRailspilProfile2 = oRailSpilSketch2.Profiles.AddForSolid
            Dim oExtrDefSp02 As ExtrudeDefinition
            oExtrDefSp02 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile2, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp02.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp02 As ExtrudeFeature
            oExtrSp02 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp02)
            'oExtrSp02.Name = "Rail Spil 2"


            If oRailLng2 >= 600 And oRailLng2 < 900 Then
                GoTo oneSpilOnly3
            End If

            '-------------------------------------Pattern Rail Spil 2-------------------------------------------------
            Dim oSpil2Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil2Col = (oExtrSp02.Definition.AffectedBodies)

            Try
                Call oExtrSp02.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try


            Dim oRPattern2 As RectangularPatternFeature
            'Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls02 As Integer
            Dim oDistSpls02 As Double

            oNbrSpls02 = Math.Ceiling((oRailLng2 - 600) / 900)
            If oRailLng2 <= 1600 Then
                oNbrSpls02 = 2
            End If
            oDistSpls02 = (oRailLng2 / 10 - 60) / (oNbrSpls02 - 1)
            oRPattern2 = oDef.Features.RectangularPatternFeatures.Add(oSpil2Col, oDef.WorkPlanes.Item(1), True, oNbrSpls02, oDistSpls02, kDefault)
            ' oRPattern2.Name = "Rail Spil 2 Pattern"

            Try
                Call oRPattern2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try



noSpils3:
oneSpilOnly3:

            ProgressBar3.Value = 70
            '---------------------------------------HANDGUIDE-------------------------------------------

            'Add a new Sketch3D

            Dim oSketch3dRail01 As Sketch3D

            oSketch3dRail01 = oDef.Sketches3D.Add

            'Add two lines to the sketch

            Dim oLine1 As SketchLine3D
            Dim oLine2 As SketchLine3D
            Dim oLine3 As SketchLine3D
            Dim oLine4 As SketchLine3D
            Dim oLine5 As SketchLine3D
            Dim oLine6 As SketchLine3D


            oLine1 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oTG.CreatePoint(0, 2.1, 50), oTG.CreatePoint(0, oRailLng1 / 10 - 0.4, 50), True, 7.5)
            oLine2 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine1.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 2.1, oRailLng1 / 10 - 0.4, 50), True, 7.5)
            oLine3 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine2.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 2.1, oRailLng1 / 10 - 0.4, 105), True, 7.5)
            oLine4 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine3.EndPoint, oTG.CreatePoint(0, oRailLng1 / 10 - 0.4, 105), True, 7.5)
            oLine5 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine4.EndPoint, oTG.CreatePoint(0, 2.1, 105), True, 7.5)
            oLine6 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine5.EndPoint, oLine1.StartPoint, True, 7.5)

            Dim oRailingPath As Path
            oRailingPath = oDef.Features.CreatePath(oLine1)

            oGroundPlane = oDef.WorkPlanes.Item(3)
            Dim oRailDiaPlane As WorkPlane
            oRailDiaPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oGroundPlane, 75, False)

            Dim oRailingSketch2 As PlanarSketch
            oRailingSketch2 = oDef.Sketches.Add(oRailDiaPlane)
            Dim oRailDiaCircle As SketchCircle
            oRailDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 2.1), 2.1)
            Dim oRailIntDiaCircle As SketchCircle
            oRailIntDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 2.1), 1.8)
            Dim oRailProfile As Profile
            oRailProfile = oRailingSketch2.Profiles.AddForSolid

            ProgressBar3.Value = 80

            Dim oSweep As SweepFeature
            oSweep = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile, oRailingPath, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)
            oGroundPlane.Visible = False
            oRailDiaPlane.Visible = False

            Try
                Call oExtrude.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

            Try
                Call oSweep.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

        End If
        ' 
        '--------------------------------------------------------------------------------------------------------
        '
        '---------------END-----------------2 SIDES--------------------------------------------------------------
        '                                                                                                       '
        '                                                                                                       '
        '                                                                                                       '
        '                                                                                                       '
        '                                                                                                       '
        '---------------START-----------------3 SIDES--------------------------------------------------------------
        '
        '
        '----------------------------------------------------------------------------------------------------------
        '

        If oRailingState = "3Sides" Then


            Dim RS As RenderStyle
            Try
                RS = oStairsDoc.RenderStyles.Item("RailingStyle")
            Catch ex As Exception
                RS = oStairsDoc.RenderStyles.Add("RailingStyle")
                RS.Reflectivity = 3
                RS.SetDiffuseColor(255, 255, 0)  'Yellow
            End Try



            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)
            ' oGroundPlane.Name = "Bottom Of Railing"

            Dim oKickplateSketch As PlanarSketch
            oKickplateSketch = oDef.Sketches.Add(oGroundPlane)

            Dim oWp01 As SketchPoints
            oWp01 = oKickplateSketch.SketchPoints


            ProgressBar3.Value = 40

            '------------Sketch points Kickplate----------------------

            Call oWp01.Add(oTG.CreatePoint2d(-0.4, 0), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4, oRailLng1 / 10), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + oRailLng2 / 10, oRailLng1 / 10), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + oRailLng2 / 10, (oRailLng1 / 10) - (oRailLng3 / 10)), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 - 0.8 + oRailLng2 / 10, (oRailLng1 / 10) - (oRailLng3 / 10)), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 - 0.8 + oRailLng2 / 10, oRailLng1 / 10 - 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(0.4, oRailLng1 / 10 - 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(0.4, 0), False)

            '-------------Draw lines Kickplate----------------------------------------------------

            Dim oLn01 As SketchLine
            oLn01 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
            Dim oLn02 As SketchLine
            oLn02 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
            Dim oLn03 As SketchLine
            oLn03 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
            Dim oLn04 As SketchLine
            oLn04 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(5))
            Dim oLn05 As SketchLine
            oLn05 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(5), oWp01(6))
            Dim oLn06 As SketchLine
            oLn06 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(6), oWp01(7))
            Dim oLn07 As SketchLine
            oLn07 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(7), oWp01(8))
            Dim oLn08 As SketchLine
            oLn08 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(8), oWp01(1))

            ProgressBar3.Value = 50

            '--------------Extrusion of the Kickplate--------------------------------------


            Dim oKickplateProfile As Profile
            oKickplateProfile = oKickplateSketch.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oKickplateProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oExtrudeDef.SetDistanceExtent(15, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrude As ExtrudeFeature
            oExtrude = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)

            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(9)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------


            'oExtrude.Name = "Kickplate"
            '--------------------------------------------------------------------------

            If oRailLng1 >= 300 And oRailLng1 < 600 Then
                GoTo noSpils4
            End If

            '-------------------------------------Spils on side 1-----------------------------------

            Dim oRailSpilSketch1 As PlanarSketch
            oRailSpilSketch1 = oDef.Sketches.Add(oGroundPlane)
            Dim oSpilCircle1 As SketchCircle
            oSpilCircle1 = oRailSpilSketch1.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 30), 3.37 / 2)
            ProgressBar3.Value = 60
            Dim oRailspilProfile1 As Profile
            oRailspilProfile1 = oRailSpilSketch1.Profiles.AddForSolid
            Dim oExtrDefSp01 As ExtrudeDefinition
            oExtrDefSp01 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile1, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp01.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp01 As ExtrudeFeature
            oExtrSp01 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp01)
            ' oExtrSp01.Name = "Rail Spil 1"
            If oRailLng1 >= 600 And oRailLng1 < 900 Then
                GoTo oneSpilOnly4
            End If

            '-------------------------------------Pattern Rail Spil 1-------------------------------------------------
            Dim oSpil1Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil1Col = (oExtrSp01.Definition.AffectedBodies)

            Dim oRPattern1 As RectangularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls01 As Integer
            Dim oDistSpls01 As Double

            oNbrSpls01 = Math.Ceiling((oRailLng1 - 600) / 900)
            If oRailLng1 <= 1600 Then
                oNbrSpls01 = 2
            End If
            oDistSpls01 = (oRailLng1 / 10 - 60) / (oNbrSpls01 - 1)
            oRPattern1 = oDef.Features.RectangularPatternFeatures.Add(oSpil1Col, oDef.WorkPlanes.Item(2), True, oNbrSpls01, oDistSpls01, kDefault)

            Try
                Call oExtrSp01.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

            Try
                Call oRPattern1.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

noSpils4:
oneSpilOnly4:


            If oRailLng2 >= 300 And oRailLng2 < 600 Then
                GoTo noSpils5
            End If
            '-------------------------------------Spils on side 2-----------------------------------

            Dim oRailSpilSketch2 As PlanarSketch
            oRailSpilSketch2 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle2 As SketchCircle
            oSpilCircle2 = oRailSpilSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(30, oRailLng1 / 10 - 0.4), 3.37 / 2)

            ProgressBar3.Value = 60

            Dim oRailspilProfile2 As Profile
            oRailspilProfile2 = oRailSpilSketch2.Profiles.AddForSolid
            Dim oExtrDefSp02 As ExtrudeDefinition
            oExtrDefSp02 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile2, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp02.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp02 As ExtrudeFeature
            oExtrSp02 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp02)
            'oExtrSp02.Name = "Rail Spil 2"


            If oRailLng2 >= 600 And oRailLng2 < 900 Then
                GoTo oneSpilOnly5
            End If

            '-------------------------------------Pattern Rail Spil 2-------------------------------------------------
            Dim oSpil2Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil2Col = (oExtrSp02.Definition.AffectedBodies)

            Dim oRPattern2 As RectangularPatternFeature
            'Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls02 As Integer
            Dim oDistSpls02 As Double

            oNbrSpls02 = Math.Ceiling((oRailLng2 - 600) / 900)
            If oRailLng2 <= 1600 Then
                oNbrSpls02 = 2
            End If
            oDistSpls02 = (oRailLng2 / 10 - 60) / (oNbrSpls02 - 1)
            oRPattern2 = oDef.Features.RectangularPatternFeatures.Add(oSpil2Col, oDef.WorkPlanes.Item(1), True, oNbrSpls02, oDistSpls02, kDefault)
            ' oRPattern2.Name = "Rail Spil 2 Pattern"

            Try
                Call oExtrSp02.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try


            Try
                Call oRPattern2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try



noSpils5:
oneSpilOnly5:




            If oRailLng3 >= 300 And oRailLng3 < 600 Then
                GoTo noSpils6
            End If
            '-------------------------------------Spils on side 3-----------------------------------

            Dim oRailSpilSketch3 As PlanarSketch
            oRailSpilSketch3 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle3 As SketchCircle
            oSpilCircle3 = oRailSpilSketch3.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oRailLng2 / 10 - 0.8, oRailLng1 / 10 - 0.4 - 30), 3.37 / 2)

            ProgressBar3.Value = 60

            Dim oRailspilProfile3 As Profile
            oRailspilProfile3 = oRailSpilSketch3.Profiles.AddForSolid
            Dim oExtrDefSp03 As ExtrudeDefinition
            oExtrDefSp03 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile3, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp03.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp03 As ExtrudeFeature
            oExtrSp03 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp03)
            ' oExtrSp03.Name = "Rail Spil 3"


            If oRailLng3 >= 600 And oRailLng3 < 900 Then
                GoTo oneSpilOnly6
            End If

            '-------------------------------------Pattern Rail Spil 3-------------------------------------------------
            Dim oSpil3Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil3Col = (oExtrSp03.Definition.AffectedBodies)

            Dim oRPattern3 As RectangularPatternFeature
            'Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls03 As Integer
            Dim oDistSpls03 As Double

            oNbrSpls03 = Math.Ceiling((oRailLng3 - 600) / 900)
            If oRailLng3 <= 1600 Then
                oNbrSpls03 = 2
            End If
            oDistSpls03 = (oRailLng3 / 10 - 60) / (oNbrSpls03 - 1)
            oRPattern3 = oDef.Features.RectangularPatternFeatures.Add(oSpil3Col, oDef.WorkPlanes.Item(2), False, oNbrSpls03, oDistSpls03, kDefault)
            ' oRPattern3.Name = "Rail Spil 3 Pattern"

            Try
                Call oExtrSp03.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

            Try
                Call oRPattern3.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try


noSpils6:
oneSpilOnly6:

            ProgressBar3.Value = 70
            '---------------------------------------HANDGUIDE-------------------------------------------

            'Add a new Sketch3D

            Dim oSketch3dRail01 As Sketch3D

            oSketch3dRail01 = oDef.Sketches3D.Add

            'Add two lines to the sketch

            Dim oLine1 As SketchLine3D
            Dim oLine2 As SketchLine3D
            Dim oLine3 As SketchLine3D
            Dim oLine4 As SketchLine3D
            Dim oLine5 As SketchLine3D
            Dim oLine6 As SketchLine3D
            Dim oLine7 As SketchLine3D
            Dim oLine8 As SketchLine3D

            oLine1 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oTG.CreatePoint(0, 2.1, 50), oTG.CreatePoint(0, oRailLng1 / 10 - 0.4, 50), True, 7.5)
            oLine2 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine1.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, oRailLng1 / 10 - 0.4, 50), True, 7.5)
            oLine3 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine2.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10), 50), True, 7.5)
            oLine4 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine3.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10), 105), True, 7.5)
            oLine5 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine4.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, oRailLng1 / 10 - 0.4, 105), True, 7.5)
            oLine6 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine5.EndPoint, oTG.CreatePoint(0, oRailLng1 / 10 - 0.4, 105), True, 7.5)
            oLine7 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine6.EndPoint, oTG.CreatePoint(0, 2.1, 105), True, 7.5)
            oLine8 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine7.EndPoint, oLine1.StartPoint, True, 7.5)

            Dim oRailingPath As Path
            oRailingPath = oDef.Features.CreatePath(oLine1)

            oGroundPlane = oDef.WorkPlanes.Item(3)
            Dim oRailDiaPlane As WorkPlane
            oRailDiaPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oGroundPlane, 75, False)

            Dim oRailingSketch2 As PlanarSketch
            oRailingSketch2 = oDef.Sketches.Add(oRailDiaPlane)
            Dim oRailDiaCircle As SketchCircle
            oRailDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 2.1), 2.1)
            Dim oRailIntDiaCircle As SketchCircle
            oRailIntDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 2.1), 1.8)
            Dim oRailProfile As Profile
            oRailProfile = oRailingSketch2.Profiles.AddForSolid

            ProgressBar3.Value = 80

            Dim oSweep As SweepFeature
            oSweep = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile, oRailingPath, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)
            ' oSweep.Name = "Railing"

            oGroundPlane.Visible = False
            oRailDiaPlane.Visible = False


            Try
                Call oExtrude.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

            Try
                Call oSweep.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

        End If




        '--------------------------------4 SIDES--------------------------------------------------------------








        If oRailingState = "4Sides" Then

            Dim RS As RenderStyle
            Try
                RS = oStairsDoc.RenderStyles.Item("RailingStyle")
            Catch ex As Exception
                RS = oStairsDoc.RenderStyles.Add("RailingStyle")
                RS.Reflectivity = 3
                RS.SetDiffuseColor(255, 255, 0)  'Yellow
            End Try


            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)
            ' oGroundPlane.Name = "Bottom Of Railing"

            Dim oKickplateSketch As PlanarSketch
            oKickplateSketch = oDef.Sketches.Add(oGroundPlane)

            Dim oWp01 As SketchPoints
            oWp01 = oKickplateSketch.SketchPoints


            ProgressBar3.Value = 40

            '------------Sketch points Kickplate----------------------

            Call oWp01.Add(oTG.CreatePoint2d(-0.4, 0), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4, oRailLng1 / 10), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + oRailLng2 / 10, oRailLng1 / 10), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + oRailLng2 / 10, (oRailLng1 / 10) - (oRailLng3 / 10)), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + (oRailLng2 / 10) - (oRailLng4 / 10), (oRailLng1 / 10) - (oRailLng3 / 10)), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + (oRailLng2 / 10) - (oRailLng4 / 10), (oRailLng1 / 10) - (oRailLng3 / 10) + 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + (oRailLng2 / 10) - 0.8, (oRailLng1 / 10) - (oRailLng3 / 10) + 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 - 0.8 + oRailLng2 / 10, oRailLng1 / 10 - 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(0.4, oRailLng1 / 10 - 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(0.4, 0), False)


            '-------------Draw lines Kickplate----------------------------------------------------

            Dim oLn01 As SketchLine
            oLn01 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
            Dim oLn02 As SketchLine
            oLn02 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
            Dim oLn03 As SketchLine
            oLn03 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
            Dim oLn04 As SketchLine
            oLn04 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(5))
            Dim oLn05 As SketchLine
            oLn05 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(5), oWp01(6))
            Dim oLn06 As SketchLine
            oLn06 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(6), oWp01(7))
            Dim oLn07 As SketchLine
            oLn07 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(7), oWp01(8))
            Dim oLn08 As SketchLine
            oLn08 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(8), oWp01(9))
            Dim oLn09 As SketchLine
            oLn09 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(9), oWp01(10))
            Dim oLn10 As SketchLine
            oLn10 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(10), oWp01(1))

            ProgressBar3.Value = 50

            '--------------Extrusion of the Kickplate--------------------------------------


            Dim oKickplateProfile As Profile
            oKickplateProfile = oKickplateSketch.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oKickplateProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oExtrudeDef.SetDistanceExtent(15, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrude As ExtrudeFeature
            oExtrude = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)

            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(11)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------


            ' oExtrude.Name = "Kickplate"
            '--------------------------------------------------------------------------

            If oRailLng1 >= 300 And oRailLng1 < 600 Then
                GoTo noSpils7
            End If

            '-------------------------------------Spils on side 1-----------------------------------

            Dim oRailSpilSketch1 As PlanarSketch
            oRailSpilSketch1 = oDef.Sketches.Add(oGroundPlane)
            Dim oSpilCircle1 As SketchCircle
            oSpilCircle1 = oRailSpilSketch1.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 30), 3.37 / 2)
            ProgressBar3.Value = 60
            Dim oRailspilProfile1 As Profile
            oRailspilProfile1 = oRailSpilSketch1.Profiles.AddForSolid
            Dim oExtrDefSp01 As ExtrudeDefinition
            oExtrDefSp01 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile1, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp01.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp01 As ExtrudeFeature
            oExtrSp01 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp01)
            '  oExtrSp01.Name = "Rail Spil 1"
            If oRailLng1 >= 600 And oRailLng1 < 900 Then
                GoTo oneSpilOnly7
            End If

            '-------------------------------------Pattern Rail Spil 1-------------------------------------------------
            Dim oSpil1Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil1Col = (oExtrSp01.Definition.AffectedBodies)

            Dim oRPattern1 As RectangularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls01 As Integer
            Dim oDistSpls01 As Double

            oNbrSpls01 = Math.Ceiling((oRailLng1 - 600) / 900)
            If oRailLng1 <= 1600 Then
                oNbrSpls01 = 2
            End If
            oDistSpls01 = (oRailLng1 / 10 - 60) / (oNbrSpls01 - 1)
            oRPattern1 = oDef.Features.RectangularPatternFeatures.Add(oSpil1Col, oDef.WorkPlanes.Item(2), True, oNbrSpls01, oDistSpls01, kDefault)
            ' oRPattern1.Name = "Rail Spil 1 Pattern"
            Try
                Call oRPattern1.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

            Try
                Call oExtrSp01.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

noSpils7:
oneSpilOnly7:


            If oRailLng2 >= 300 And oRailLng2 < 600 Then
                GoTo noSpils8
            End If
            '-------------------------------------Spils on side 2-----------------------------------

            Dim oRailSpilSketch2 As PlanarSketch
            oRailSpilSketch2 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle2 As SketchCircle
            oSpilCircle2 = oRailSpilSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(30, oRailLng1 / 10 - 0.4), 3.37 / 2)

            ProgressBar3.Value = 60

            Dim oRailspilProfile2 As Profile
            oRailspilProfile2 = oRailSpilSketch2.Profiles.AddForSolid
            Dim oExtrDefSp02 As ExtrudeDefinition
            oExtrDefSp02 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile2, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp02.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp02 As ExtrudeFeature
            oExtrSp02 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp02)
            '  oExtrSp02.Name = "Rail Spil 2"


            If oRailLng2 >= 600 And oRailLng2 < 900 Then
                GoTo oneSpilOnly8
            End If

            '-------------------------------------Pattern Rail Spil 2-------------------------------------------------
            Dim oSpil2Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil2Col = (oExtrSp02.Definition.AffectedBodies)

            Dim oRPattern2 As RectangularPatternFeature
            'Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls02 As Integer
            Dim oDistSpls02 As Double

            oNbrSpls02 = Math.Ceiling((oRailLng2 - 600) / 900)
            If oRailLng2 <= 1600 Then
                oNbrSpls02 = 2
            End If
            oDistSpls02 = (oRailLng2 / 10 - 60) / (oNbrSpls02 - 1)
            oRPattern2 = oDef.Features.RectangularPatternFeatures.Add(oSpil2Col, oDef.WorkPlanes.Item(1), True, oNbrSpls02, oDistSpls02, kDefault)
            '  oRPattern2.Name = "Rail Spil 2 Pattern"
            Try
                Call oRPattern2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try
            Try
                Call oExtrSp02.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

noSpils8:
oneSpilOnly8:




            If oRailLng3 >= 300 And oRailLng3 < 600 Then
                GoTo noSpils9
            End If
            '-------------------------------------Spils on side 3-----------------------------------

            Dim oRailSpilSketch3 As PlanarSketch
            oRailSpilSketch3 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle3 As SketchCircle
            oSpilCircle3 = oRailSpilSketch3.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oRailLng2 / 10 - 0.8, oRailLng1 / 10 - 0.4 - 30), 3.37 / 2)

            ProgressBar3.Value = 60

            Dim oRailspilProfile3 As Profile
            oRailspilProfile3 = oRailSpilSketch3.Profiles.AddForSolid
            Dim oExtrDefSp03 As ExtrudeDefinition
            oExtrDefSp03 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile3, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp03.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp03 As ExtrudeFeature
            oExtrSp03 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp03)
            ' oExtrSp03.Name = "Rail Spil 3"


            If oRailLng3 >= 600 And oRailLng3 < 900 Then
                GoTo oneSpilOnly9
            End If

            '-------------------------------------Pattern Rail Spil 3-------------------------------------------------
            Dim oSpil3Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil3Col = (oExtrSp03.Definition.AffectedBodies)

            Dim oRPattern3 As RectangularPatternFeature
            'Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls03 As Integer
            Dim oDistSpls03 As Double

            oNbrSpls03 = Math.Ceiling((oRailLng3 - 600) / 900)
            If oRailLng3 <= 1600 Then
                oNbrSpls03 = 2
            End If
            oDistSpls03 = (oRailLng3 / 10 - 60) / (oNbrSpls03 - 1)
            oRPattern3 = oDef.Features.RectangularPatternFeatures.Add(oSpil3Col, oDef.WorkPlanes.Item(2), False, oNbrSpls03, oDistSpls03, kDefault)
            ' oRPattern3.Name = "Rail Spil 3 Pattern"
            Try
                Call oRPattern3.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try
            Try
                Call oExtrSp03.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

noSpils9:
oneSpilOnly9:


            '-------------------------------------Spils on side 4-----------------------------------

            If oRailLng4 >= 300 And oRailLng4 < 600 Then
                GoTo noSpils10
            End If

            Dim oRailSpilSketch4 As PlanarSketch
            oRailSpilSketch4 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle4 As SketchCircle
            oSpilCircle4 = oRailSpilSketch4.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oRailLng2 / 10 - 30, (oRailLng1 / 10) - 0.4 - (oRailLng3 / 10)), 3.37 / 2)

            ProgressBar3.Value = 60

            Dim oRailspilProfile4 As Profile
            oRailspilProfile4 = oRailSpilSketch4.Profiles.AddForSolid
            Dim oExtrDefSp04 As ExtrudeDefinition
            oExtrDefSp04 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile4, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp04.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp04 As ExtrudeFeature
            oExtrSp04 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp04)
            '  oExtrSp04.Name = "Rail Spil 4"


            If oRailLng4 >= 600 And oRailLng4 < 900 Then
                GoTo oneSpilOnly10
            End If

            '-------------------------------------Pattern Rail Spil 3-------------------------------------------------
            Dim oSpil4Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil4Col = (oExtrSp04.Definition.AffectedBodies)

            Dim oRPattern4 As RectangularPatternFeature
            'Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls04 As Integer
            Dim oDistSpls04 As Double

            oNbrSpls04 = Math.Ceiling((oRailLng4 - 600) / 900)
            If oRailLng4 <= 1600 Then
                oNbrSpls04 = 2
            End If
            oDistSpls04 = (oRailLng4 / 10 - 60) / (oNbrSpls04 - 1)
            oRPattern4 = oDef.Features.RectangularPatternFeatures.Add(oSpil4Col, oDef.WorkPlanes.Item(1), False, oNbrSpls04, oDistSpls04, kDefault)
            '  oRPattern4.Name = "Rail Spil 4 Pattern"
            Try
                Call oRPattern4.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try
            Try
                Call oExtrSp04.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try



noSpils10:
oneSpilOnly10:



            ProgressBar3.Value = 70
            '---------------------------------------HANDGUIDE-------------------------------------------

            'Add a new Sketch3D

            Dim oSketch3dRail01 As Sketch3D

            oSketch3dRail01 = oDef.Sketches3D.Add

            'Add two lines to the sketch

            Dim oLine1 As SketchLine3D
            Dim oLine2 As SketchLine3D
            Dim oLine3 As SketchLine3D
            Dim oLine4 As SketchLine3D
            Dim oLine5 As SketchLine3D
            Dim oLine6 As SketchLine3D
            Dim oLine7 As SketchLine3D
            Dim oLine8 As SketchLine3D
            Dim oLine9 As SketchLine3D
            Dim oLine10 As SketchLine3D

            oLine1 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oTG.CreatePoint(0, 2.1, 50), oTG.CreatePoint(0, oRailLng1 / 10 - 0.4, 50), True, 7.5)
            oLine2 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine1.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, oRailLng1 / 10 - 0.4, 50), True, 7.5)
            oLine3 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine2.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10) + 0.8, 50), True, 7.5)
            oLine4 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine3.EndPoint, oTG.CreatePoint((oRailLng2 / 10) - (oRailLng4 / 10) - 0.8, (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10) + 0.8, 50), True, 7.5)
            oLine5 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine4.EndPoint, oTG.CreatePoint((oRailLng2 / 10) - (oRailLng4 / 10) - 0.8, (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10) + 0.8, 105), True, 7.5)
            oLine6 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine5.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10) + 0.8, 105), True, 7.5)
            oLine7 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine6.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, oRailLng1 / 10 - 0.4, 105), True, 7.5)
            oLine8 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine7.EndPoint, oTG.CreatePoint(0, oRailLng1 / 10 - 0.4, 105), True, 7.5)
            oLine9 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine8.EndPoint, oTG.CreatePoint(0, 2.1, 105), True, 7.5)
            oLine10 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine9.EndPoint, oLine1.StartPoint, True, 7.5)

            Dim oRailingPath As Path
            oRailingPath = oDef.Features.CreatePath(oLine1)

            oGroundPlane = oDef.WorkPlanes.Item(3)
            Dim oRailDiaPlane As WorkPlane
            oRailDiaPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oGroundPlane, 75, False)

            Dim oRailingSketch2 As PlanarSketch
            oRailingSketch2 = oDef.Sketches.Add(oRailDiaPlane)
            Dim oRailDiaCircle As SketchCircle
            oRailDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 2.1), 2.1)
            Dim oRailIntDiaCircle As SketchCircle
            oRailIntDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 2.1), 1.8)
            Dim oRailProfile As Profile
            oRailProfile = oRailingSketch2.Profiles.AddForSolid

            ProgressBar3.Value = 80

            Dim oSweep As SweepFeature
            oSweep = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile, oRailingPath, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)
            'oSweep.Name = "Railing"

            oGroundPlane.Visible = False
            oRailDiaPlane.Visible = False


            'assign new render style to part feature
            'Dim oRailFeature As PartFeature = oExtrude
            Try
                Call oExtrude.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try


            Try
                Call oSweep.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

        End If

        '--------------------------------5 SIDES--------------------------------------------------------------








        If oRailingState = "5Sides" Then


            Dim RS As RenderStyle
            Try
                RS = oStairsDoc.RenderStyles.Item("RailingStyle")
            Catch ex As Exception
                RS = oStairsDoc.RenderStyles.Add("RailingStyle")
                RS.Reflectivity = 3
                RS.SetDiffuseColor(255, 255, 0)  'Yellow
            End Try



            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)
            '   oGroundPlane.Name = "Bottom Of Railing"

            Dim oKickplateSketch As PlanarSketch
            oKickplateSketch = oDef.Sketches.Add(oGroundPlane)

            Dim oWp01 As SketchPoints
            oWp01 = oKickplateSketch.SketchPoints


            ProgressBar3.Value = 40

            '------------Sketch points Kickplate----------------------

            Call oWp01.Add(oTG.CreatePoint2d(-0.4, 0), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4, oRailLng1 / 10), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + oRailLng2 / 10, oRailLng1 / 10), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + oRailLng2 / 10, (oRailLng1 / 10) - (oRailLng3 / 10)), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + (oRailLng2 / 10) - (oRailLng4 / 10), (oRailLng1 / 10) - (oRailLng3 / 10)), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + (oRailLng2 / 10) - (oRailLng4 / 10), (oRailLng1 / 10) + (oRailLng5 / 10) - (oRailLng3 / 10)), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + 0.8 + (oRailLng2 / 10) - (oRailLng4 / 10), (oRailLng1 / 10) + (oRailLng5 / 10) - (oRailLng3 / 10)), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + 0.8 + (oRailLng2 / 10) - (oRailLng4 / 10), (oRailLng1 / 10) - (oRailLng3 / 10) + 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 + (oRailLng2 / 10) - 0.8, (oRailLng1 / 10) - (oRailLng3 / 10) + 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(-0.4 - 0.8 + oRailLng2 / 10, oRailLng1 / 10 - 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(0.4, oRailLng1 / 10 - 0.8), False)
            Call oWp01.Add(oTG.CreatePoint2d(0.4, 0), False)


            '-------------Draw lines Kickplate----------------------------------------------------

            Dim oLn01 As SketchLine
            oLn01 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
            Dim oLn02 As SketchLine
            oLn02 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
            Dim oLn03 As SketchLine
            oLn03 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
            Dim oLn04 As SketchLine
            oLn04 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(5))
            Dim oLn05 As SketchLine
            oLn05 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(5), oWp01(6))
            Dim oLn06 As SketchLine
            oLn06 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(6), oWp01(7))
            Dim oLn07 As SketchLine
            oLn07 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(7), oWp01(8))
            Dim oLn08 As SketchLine
            oLn08 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(8), oWp01(9))
            Dim oLn09 As SketchLine
            oLn09 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(9), oWp01(10))
            Dim oLn10 As SketchLine
            oLn10 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(10), oWp01(11))
            Dim oLn11 As SketchLine
            oLn11 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(11), oWp01(12))
            Dim oLn12 As SketchLine
            oLn12 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(12), oWp01(1))

            ProgressBar3.Value = 50

            '--------------Extrusion of the Kickplate--------------------------------------


            Dim oKickplateProfile As Profile
            oKickplateProfile = oKickplateSketch.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oKickplateProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oExtrudeDef.SetDistanceExtent(15, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrude As ExtrudeFeature
            oExtrude = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)


            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(13)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------


            ' oExtrude.Name = "Kickplate"
            '--------------------------------------------------------------------------

            If oRailLng1 >= 300 And oRailLng1 < 600 Then
                GoTo noSpils12
            End If

            '-------------------------------------Spils on side 1-----------------------------------

            Dim oRailSpilSketch1 As PlanarSketch
            oRailSpilSketch1 = oDef.Sketches.Add(oGroundPlane)
            Dim oSpilCircle1 As SketchCircle
            oSpilCircle1 = oRailSpilSketch1.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 30), 3.37 / 2)
            ProgressBar3.Value = 60
            Dim oRailspilProfile1 As Profile
            oRailspilProfile1 = oRailSpilSketch1.Profiles.AddForSolid
            Dim oExtrDefSp01 As ExtrudeDefinition
            oExtrDefSp01 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile1, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp01.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp01 As ExtrudeFeature
            oExtrSp01 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp01)
            '  oExtrSp01.Name = "Rail Spil 1"
            If oRailLng1 >= 600 And oRailLng1 < 900 Then
                GoTo oneSpilOnly12
            End If

            '-------------------------------------Pattern Rail Spil 1-------------------------------------------------
            Dim oSpil1Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil1Col = (oExtrSp01.Definition.AffectedBodies)

            Dim oRPattern1 As RectangularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls01 As Integer
            Dim oDistSpls01 As Double

            oNbrSpls01 = Math.Ceiling((oRailLng1 - 600) / 900)
            If oRailLng1 <= 1600 Then
                oNbrSpls01 = 2
            End If
            oDistSpls01 = (oRailLng1 / 10 - 60) / (oNbrSpls01 - 1)
            oRPattern1 = oDef.Features.RectangularPatternFeatures.Add(oSpil1Col, oDef.WorkPlanes.Item(2), True, oNbrSpls01, oDistSpls01, kDefault)
            '    oRPattern1.Name = "Rail Spil 1 Pattern"
            Try
                Call oExtrSp01.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try
            Try
                Call oRPattern1.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try
noSpils12:
oneSpilOnly12:


            If oRailLng2 >= 300 And oRailLng2 < 600 Then
                GoTo noSpils13
            End If
            '-------------------------------------Spils on side 2-----------------------------------

            Dim oRailSpilSketch2 As PlanarSketch
            oRailSpilSketch2 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle2 As SketchCircle
            oSpilCircle2 = oRailSpilSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(30, oRailLng1 / 10 - 0.4), 3.37 / 2)

            ProgressBar3.Value = 60

            Dim oRailspilProfile2 As Profile
            oRailspilProfile2 = oRailSpilSketch2.Profiles.AddForSolid
            Dim oExtrDefSp02 As ExtrudeDefinition
            oExtrDefSp02 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile2, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp02.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp02 As ExtrudeFeature
            oExtrSp02 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp02)
            '   oExtrSp02.Name = "Rail Spil 2"


            If oRailLng2 >= 600 And oRailLng2 < 900 Then
                GoTo oneSpilOnly13
            End If

            '-------------------------------------Pattern Rail Spil 2-------------------------------------------------
            Dim oSpil2Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil2Col = (oExtrSp02.Definition.AffectedBodies)

            Dim oRPattern2 As RectangularPatternFeature
            'Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls02 As Integer
            Dim oDistSpls02 As Double

            oNbrSpls02 = Math.Ceiling((oRailLng2 - 600) / 900)
            If oRailLng2 <= 1600 Then
                oNbrSpls02 = 2
            End If
            oDistSpls02 = (oRailLng2 / 10 - 60) / (oNbrSpls02 - 1)
            oRPattern2 = oDef.Features.RectangularPatternFeatures.Add(oSpil2Col, oDef.WorkPlanes.Item(1), True, oNbrSpls02, oDistSpls02, kDefault)
            '  oRPattern2.Name = "Rail Spil 2 Pattern"
            Try
                Call oRPattern2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try
            Try
                Call oExtrSp02.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

noSpils13:
oneSpilOnly13:




            If oRailLng3 >= 300 And oRailLng3 < 600 Then
                GoTo noSpils14
            End If
            '-------------------------------------Spils on side 3-----------------------------------

            Dim oRailSpilSketch3 As PlanarSketch
            oRailSpilSketch3 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle3 As SketchCircle
            oSpilCircle3 = oRailSpilSketch3.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oRailLng2 / 10 - 0.8, oRailLng1 / 10 - 0.4 - 30), 3.37 / 2)

            ProgressBar3.Value = 60

            Dim oRailspilProfile3 As Profile
            oRailspilProfile3 = oRailSpilSketch3.Profiles.AddForSolid
            Dim oExtrDefSp03 As ExtrudeDefinition
            oExtrDefSp03 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile3, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp03.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp03 As ExtrudeFeature
            oExtrSp03 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp03)
            '   oExtrSp03.Name = "Rail Spil 3"


            If oRailLng3 >= 600 And oRailLng3 < 900 Then
                GoTo oneSpilOnly14
            End If

            '-------------------------------------Pattern Rail Spil 3-------------------------------------------------
            Dim oSpil3Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil3Col = (oExtrSp03.Definition.AffectedBodies)

            Dim oRPattern3 As RectangularPatternFeature
            'Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls03 As Integer
            Dim oDistSpls03 As Double

            oNbrSpls03 = Math.Ceiling((oRailLng3 - 600) / 900)
            If oRailLng3 <= 1600 Then
                oNbrSpls03 = 2
            End If
            oDistSpls03 = (oRailLng3 / 10 - 60) / (oNbrSpls03 - 1)
            oRPattern3 = oDef.Features.RectangularPatternFeatures.Add(oSpil3Col, oDef.WorkPlanes.Item(2), False, oNbrSpls03, oDistSpls03, kDefault)
            '   oRPattern3.Name = "Rail Spil 3 Pattern"
            Try
                Call oRPattern3.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try
            Try
                Call oExtrSp03.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

noSpils14:
oneSpilOnly14:


            '-------------------------------------Spils on side 4-----------------------------------

            If oRailLng4 >= 300 And oRailLng4 < 600 Then
                GoTo noSpils15
            End If

            Dim oRailSpilSketch4 As PlanarSketch
            oRailSpilSketch4 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle4 As SketchCircle
            oSpilCircle4 = oRailSpilSketch4.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oRailLng2 / 10 - 30, (oRailLng1 / 10) + 0.4 - (oRailLng3 / 10)), 3.37 / 2)

            ProgressBar3.Value = 60

            Dim oRailspilProfile4 As Profile
            oRailspilProfile4 = oRailSpilSketch4.Profiles.AddForSolid
            Dim oExtrDefSp04 As ExtrudeDefinition
            oExtrDefSp04 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile4, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp04.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp04 As ExtrudeFeature
            oExtrSp04 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp04)
            '   oExtrSp04.Name = "Rail Spil 4"


            If oRailLng4 >= 600 And oRailLng4 < 900 Then
                GoTo oneSpilOnly15
            End If

            '-------------------------------------Pattern Rail Spil 4-------------------------------------------------
            Dim oSpil4Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil4Col = (oExtrSp04.Definition.AffectedBodies)

            Dim oRPattern4 As RectangularPatternFeature
            'Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls04 As Integer
            Dim oDistSpls04 As Double

            oNbrSpls04 = Math.Ceiling((oRailLng4 - 600) / 900)
            If oRailLng4 <= 1600 Then
                oNbrSpls04 = 2
            End If
            oDistSpls04 = (oRailLng4 / 10 - 60) / (oNbrSpls04 - 1)
            oRPattern4 = oDef.Features.RectangularPatternFeatures.Add(oSpil4Col, oDef.WorkPlanes.Item(1), False, oNbrSpls04, oDistSpls04, kDefault)
            '   oRPattern4.Name = "Rail Spil 4 Pattern"
            Try
                Call oRPattern4.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try
            Try
                Call oExtrSp04.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

noSpils15:
oneSpilOnly15:

            '-------------------------------------Spils on side 5-----------------------------------

            If oRailLng5 >= 300 And oRailLng5 < 600 Then
                GoTo noSpils16
            End If

            Dim oRailSpilSketch5 As PlanarSketch
            oRailSpilSketch5 = oDef.Sketches.Add(oGroundPlane)

            Dim oSpilCircle5 As SketchCircle
            oSpilCircle5 = oRailSpilSketch5.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(oRailLng2 / 10 - (oRailLng4 / 10), (oRailLng1 / 10) + 0.4 - (oRailLng3 / 10) + 30), 3.37 / 2)

            ProgressBar3.Value = 60

            Dim oRailspilProfile5 As Profile
            oRailspilProfile5 = oRailSpilSketch5.Profiles.AddForSolid
            Dim oExtrDefSp05 As ExtrudeDefinition
            oExtrDefSp05 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile5, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp05.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp05 As ExtrudeFeature
            oExtrSp05 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp05)
            ' oExtrSp05.Name = "Rail Spil 5"


            If oRailLng5 >= 600 And oRailLng5 < 900 Then
                GoTo oneSpilOnly16
            End If

            '-------------------------------------Pattern Rail Spil 5-------------------------------------------------
            Dim oSpil5Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil5Col = (oExtrSp05.Definition.AffectedBodies)

            Dim oRPattern5 As RectangularPatternFeature
            'Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls05 As Integer
            Dim oDistSpls05 As Double

            oNbrSpls05 = Math.Ceiling((oRailLng5 - 600) / 900)
            If oRailLng5 <= 1600 Then
                oNbrSpls05 = 2
            End If
            oDistSpls05 = (oRailLng5 / 10 - 60) / (oNbrSpls05 - 1)
            oRPattern5 = oDef.Features.RectangularPatternFeatures.Add(oSpil5Col, oDef.WorkPlanes.Item(2), True, oNbrSpls05, oDistSpls05, kDefault)
            ' oRPattern5.Name = "Rail Spil 5 Pattern"
            Try
                Call oExtrSp05.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

            Try
                Call oRPattern5.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Catch ex As Exception
            End Try

noSpils16:
oneSpilOnly16:

            ProgressBar3.Value = 70
            '---------------------------------------HANDGUIDE-------------------------------------------

            'Add a new Sketch3D

            Dim oSketch3dRail01 As Sketch3D

            oSketch3dRail01 = oDef.Sketches3D.Add

            'Add two lines to the sketch

            Dim oLine1 As SketchLine3D
            Dim oLine2 As SketchLine3D
            Dim oLine3 As SketchLine3D
            Dim oLine4 As SketchLine3D
            Dim oLine5 As SketchLine3D
            Dim oLine6 As SketchLine3D
            Dim oLine7 As SketchLine3D
            Dim oLine8 As SketchLine3D
            Dim oLine9 As SketchLine3D
            Dim oLine10 As SketchLine3D
            Dim oLine11 As SketchLine3D
            Dim oLine12 As SketchLine3D

            oLine1 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oTG.CreatePoint(0, 2.1, 50), oTG.CreatePoint(0, oRailLng1 / 10 - 0.4, 50), True, 7.5)
            oLine2 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine1.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, oRailLng1 / 10 - 0.4, 50), True, 7.5)
            oLine3 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine2.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10) + 0.8, 50), True, 7.5)
            oLine4 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine3.EndPoint, oTG.CreatePoint((oRailLng2 / 10) - (oRailLng4 / 10), (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10) + 0.8, 50), True, 7.5)
            oLine5 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine4.EndPoint, oTG.CreatePoint((oRailLng2 / 10) - (oRailLng4 / 10), (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10) + 0.8 + (oRailLng5 / 10), 50), True, 7.5)
            oLine6 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine5.EndPoint, oTG.CreatePoint((oRailLng2 / 10) - (oRailLng4 / 10), (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10) + 0.8 + (oRailLng5 / 10), 105), True, 7.5)
            oLine7 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine6.EndPoint, oTG.CreatePoint((oRailLng2 / 10) - (oRailLng4 / 10), (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10) + 0.8, 105), True, 7.5)
            oLine8 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine7.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, (oRailLng1 / 10 - 0.4) - (oRailLng3 / 10) + 0.8, 105), True, 7.5)
            oLine9 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine8.EndPoint, oTG.CreatePoint(oRailLng2 / 10 - 0.8, oRailLng1 / 10 - 0.4, 105), True, 7.5)
            oLine10 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine9.EndPoint, oTG.CreatePoint(0, oRailLng1 / 10 - 0.4, 105), True, 7.5)
            oLine11 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine10.EndPoint, oTG.CreatePoint(0, 2.1, 105), True, 7.5)
            oLine12 = oSketch3dRail01.SketchLines3D.AddByTwoPoints(oLine11.EndPoint, oLine1.StartPoint, True, 7.5)

            Dim oRailingPath As Path
            oRailingPath = oDef.Features.CreatePath(oLine1)

            oGroundPlane = oDef.WorkPlanes.Item(3)
            Dim oRailDiaPlane As WorkPlane
            oRailDiaPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oGroundPlane, 75, False)

            Dim oRailingSketch2 As PlanarSketch
            oRailingSketch2 = oDef.Sketches.Add(oRailDiaPlane)
            Dim oRailDiaCircle As SketchCircle
            oRailDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 2.1), 2.1)
            Dim oRailIntDiaCircle As SketchCircle
            oRailIntDiaCircle = oRailingSketch2.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 2.1), 1.8)
            Dim oRailProfile As Profile
            oRailProfile = oRailingSketch2.Profiles.AddForSolid

            ProgressBar3.Value = 80

            Dim oSweep As SweepFeature
            oSweep = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile, oRailingPath, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)
            '   oSweep.Name = "Railing"

            oGroundPlane.Visible = False
            oRailDiaPlane.Visible = False



            'assign new render style to part feature
            'Dim oRailFeature As PartFeature = oExtrude

            Call oExtrude.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)








            Call oSweep.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)

        End If

        '---------------------------------------ROUND-------------------------------

        If oRailingState = "Round" Then

            ProgressBar3.Visible = True
            ProgressBar3.Value = 10

            Dim oMiddlePlane As WorkPlane
            oMiddlePlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(2), False)
            '  oMiddlePlane.Name = "Middle Of Railing"

            Dim oKickplateSketch As PlanarSketch
            oKickplateSketch = oDef.Sketches.Add(oMiddlePlane)

            Dim oWp01 As SketchPoints
            oWp01 = oKickplateSketch.SketchPoints
            'Sketchpoints
            Call oWp01.Add(oTG.CreatePoint2d(oRailDia / 20 - 0.4, 0), False)
            Call oWp01.Add(oTG.CreatePoint2d(oRailDia / 20 - 0.4, 15), False)
            Call oWp01.Add(oTG.CreatePoint2d(oRailDia / 20 + 0.4, 15), False)
            Call oWp01.Add(oTG.CreatePoint2d(oRailDia / 20 + 0.4, 0), False)
            Dim oLn01 As SketchLine
            oLn01 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
            Dim oLn02 As SketchLine
            oLn02 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
            Dim oLn03 As SketchLine
            oLn03 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
            Dim oLn04 As SketchLine
            oLn04 = oKickplateSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(1))
            Dim oKickplateProfile As Profile
            oKickplateProfile = oKickplateSketch.Profiles.AddForSolid
            Dim oRevDef As RevolveFeature
            oRevDef = oDef.Features.RevolveFeatures.AddByAngle(oKickplateProfile, oDef.WorkAxes.Item(3), oRailAngle * pi / 180, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection, PartFeatureOperationEnum.kJoinOperation)

            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(1)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------


            '   oRevDef.Name = "Kickplate"

            Dim o961Angl As Double
            o961Angl = (Math.Atan(9.61 / (oRailDia / 20)))



            Dim oRailingSketch As PlanarSketch
            oRailingSketch = oDef.Sketches.Add(oMiddlePlane)

            Dim oWp02 As SketchPoints
            oWp02 = oRailingSketch.SketchPoints

            Call oWp02.Add(oTG.CreatePoint2d(oRailDia / 20, 50), False)
            Call oWp02.Add(oTG.CreatePoint2d(oRailDia / 20, 105), False)

            Dim oCir01 As SketchCircle
            oCir01 = oRailingSketch.SketchCircles.AddByCenterRadius(oWp02(1), 2.1)
            Dim oCir02 As SketchCircle
            oCir02 = oRailingSketch.SketchCircles.AddByCenterRadius(oWp02(2), 2.1)
            Dim oPipeProfile As Profile
            oPipeProfile = oRailingSketch.Profiles.AddForSolid
            Dim oRevDef2 As RevolveFeature
            oRevDef2 = oDef.Features.RevolveFeatures.AddByAngle(oPipeProfile, oDef.WorkAxes.Item(3), (oRailAngle * pi / 180) - o961Angl - o961Angl, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection, PartFeatureOperationEnum.kJoinOperation)
            '    oRevDef2.Name = "HandGuidPipe"

            Dim oStartPlane As WorkPlane
            oStartPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oDef.WorkAxes(3), oDef.WorkPlanes(2), (oRailAngle * pi / 180 / 2) - o961Angl, False)
            '   oMiddlePlane.Name = "Middle Of Railing"
            ProgressBar3.Value = 100
            Dim oEndRailSketch As PlanarSketch
            oEndRailSketch = oDef.Sketches.Add(oStartPlane)
            Dim oWp03 As SketchPoints
            oWp03 = oEndRailSketch.SketchPoints
            Call oWp03.Add(oTG.CreatePoint2d(105, -oRailDia / 20), False)
            Call oWp03.Add(oTG.CreatePoint2d(0, -oRailDia / 20), False)

            Dim oCir03 As SketchCircle
            oCir03 = oEndRailSketch.SketchCircles.AddByCenterRadius(oWp03(1), 2.1)
            Dim oLn05 As SketchLine
            oLn05 = oEndRailSketch.SketchLines.AddByTwoPoints(oWp03(1), oWp03(2))
            Dim oEndMidPlane As WorkPlane
            oEndMidPlane = oDef.WorkPlanes.AddByLinePlaneAndAngle(oLn05, oStartPlane, 90 * pi / 180, False)

            Dim oEndMidSketch As PlanarSketch
            oEndMidSketch = oDef.Sketches.Add(oEndMidPlane)

            Dim oWp04 As SketchPoints
            oWp04 = oEndMidSketch.SketchPoints

            Call oWp04.Add(oTG.CreatePoint2d(0, 0), False)
            Call oWp04.Add(oTG.CreatePoint2d(0, -7.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(55, -7.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(55, 0), False)

            Dim oLn07 As SketchLine
            oLn07 = oEndMidSketch.SketchLines.AddByTwoPoints(oWp04(1), oWp04(2))
            Dim oLn08 As SketchLine
            oLn08 = oEndMidSketch.SketchLines.AddByTwoPoints(oWp04(2), oWp04(3))
            Dim oLn09 As SketchLine
            oLn09 = oEndMidSketch.SketchLines.AddByTwoPoints(oWp04(3), oWp04(4))
            Dim oArc1 As SketchArc
            oArc1 = oEndMidSketch.SketchArcs.AddByFillet(oLn07, oLn08, 7.5, oLn07.Geometry.MidPoint, oLn08.Geometry.MidPoint)
            Dim oArc2 As SketchArc
            oArc2 = oEndMidSketch.SketchArcs.AddByFillet(oLn08, oLn09, 7.5, oLn08.Geometry.MidPoint, oLn09.Geometry.MidPoint)
            Dim oRailProfile As Profile
            oRailProfile = oEndRailSketch.Profiles.AddForSolid
            Dim oRailingPath As Path
            oRailingPath = oDef.Features.CreatePath(oLn08)

            Dim oSweep As SweepFeature
            oSweep = oDef.Features.SweepFeatures.AddUsingPath(oRailProfile, oRailingPath, PartFeatureOperationEnum.kNewBodyOperation, SweepProfileOrientationEnum.kNormalToPath, 0)
            ' oSweep.Name = "Railing"

            ' Dim oMir As MirrorFeature
            ' oMir = oDef.Features.MirrorFeatures.CreateDefinition(oSweep, oDef.WorkPlanes.Item(2), PatternComputeTypeEnum.kOptimizedCompute)

            Dim oCol As ObjectCollection
            Dim CreateObjectCollection3 As ObjectCollection = Nothing

            oCol = (oSweep.Definition.AffectedBodies)
            Dim oPlane As WorkPlane
            oPlane = oStairsDoc.ComponentDefinition.WorkPlanes(2)
            Dim oMirror As MirrorFeature
            oMirror = oStairsDoc.ComponentDefinition.Features.MirrorFeatures.Add(oCol, oPlane, False, PatternComputeTypeEnum.kOptimizedCompute)


            '------------------------------------Spils-------------------------------
            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)
            '   oGroundPlane.Name = "Bottom of Railing"
            Dim oSplSketch As PlanarSketch
            oSplSketch = oDef.Sketches.Add(oGroundPlane)
            Dim oWp05 As SketchPoints
            oWp05 = oSplSketch.SketchPoints
            Call oWp05.Add(oTG.CreatePoint2d(0, 0), False)
            Dim oArc03 As SketchArc
            oArc03 = oSplSketch.SketchArcs.AddByCenterStartSweepAngle(oWp05(1), oRailDia / 20, 180 * pi / 180, (oRailAngle * pi / 180 / 2) - o961Angl * 3)
            Dim oSpilCircle1 As SketchCircle
            oSpilCircle1 = oSplSketch.SketchCircles.AddByCenterRadius((oArc03.EndSketchPoint), 3.37 / 2)
            ProgressBar3.Value = 60
            Dim oRailspilProfile1 As Profile
            oRailspilProfile1 = oSplSketch.Profiles.AddForSolid
            Dim oExtrDefSp01 As ExtrudeDefinition
            oExtrDefSp01 = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oRailspilProfile1, PartFeatureOperationEnum.kNewBodyOperation)
            Call oExtrDefSp01.SetDistanceExtent(105, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrSp01 As ExtrudeFeature
            oExtrSp01 = oDef.Features.ExtrudeFeatures.Add(oExtrDefSp01)
            '  oExtrSp01.Name = "Rail Spil 1"

            Dim oSpil1Col As ObjectCollection
            'Dim CreateObjColA As ObjectCollection = Nothing

            oSpil1Col = (oExtrSp01.Definition.AffectedBodies)

            Dim oRPattern1 As CircularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing
            Dim oNbrSpls01 As Integer
            Dim oAnglSpils As Double
            Dim oAxis As WorkAxis

            oAxis = oDef.WorkAxes.Item(3)

            oNbrSpls01 = Math.Ceiling((oRailDia * pi / 1200) / (360 / oRailAngle))

            oAnglSpils = (oRailAngle * pi / 180) - o961Angl * 6


            oRPattern1 = oDef.Features.CircularPatternFeatures.Add(oSpil1Col, oAxis, False, oNbrSpls01, oAnglSpils, True,)
            '  oRPattern1.Name = "Rail Spil 1 Pattern"

            oStartPlane.Visible = False
            oGroundPlane.Visible = False
            oPlane.Visible = False
            oEndMidPlane.Visible = False
            oMiddlePlane.Visible = False

            Dim RS As RenderStyle
            Try
                RS = oStairsDoc.RenderStyles.Item("RailingStyle")
            Catch ex As Exception
                RS = oStairsDoc.RenderStyles.Add("RailingStyle")
                RS.Reflectivity = 3
                RS.SetDiffuseColor(255, 255, 0)  'Yellow
            End Try
            'assign new render style to part feature
            'Dim oRailFeature As PartFeature = oExtrude

            Call oExtrSp01.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oRPattern1.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oSweep.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oMirror.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oRevDef2.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
            Call oRevDef.SetRenderStyle(StyleSourceTypeEnum.kOverrideRenderStyle, RS)
        End If


        'Return view to Home view
        oInvApp.CommandManager.ControlDefinitions.Item("AppViewCubeHomeCmd").Execute()
        'zoom all
        oInvApp.ActiveView.GoHome()
        ProgressBar3.Value = 100
        oStairsDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits



        Try
            oStairsDoc.SaveAs(lblRailPath.Text & oRailFilename, False)
        Catch ex As Exception
            MsgBox("Could not save the part." & vbCrLf & "Try again or check your credentials", vbOKOnly + "4064", "Warning")
            oInvApp.ScreenUpdating = True
            GoTo EndRoutine
        End Try



        oStairsDoc.Close()
        oInvApp.ScreenUpdating = True

        '--------------Place the part interactive--------------------------------------
PlacePart:
        ProgressBar3.Value = 0
        ' Get the active assembly.

        asmDoc = oInvApp.ActiveDocument

        ' If assembly empty place stairs grounded at origin.

        If asmDoc.ComponentDefinition.ImmediateReferencedDefinitions.Count < 1 Then
            Dim trans As Matrix
            trans = oInvApp.TransientGeometry.CreateMatrix
            stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblRailPath.Text & oRailFilename), trans)
            stairsOcc.Grounded = True
            GoTo EndRoutine
        End If

        'Else place stairs on a floor surface.

        Hide()
        oFloorFace = oInvApp.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, "Select a floor to place the stairs.")
        If oFloorFace Is Nothing Then
            GoTo EndRoutine
        End If

        stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblRailPath.Text & oRailFilename), oInvApp.TransientGeometry.CreateMatrix)
        oStairsDoc = stairsOcc.Definition.Document

        Try
            attrSets = oStairsDoc.AttributeManager.FindAttributeSets("ConnectFace")
            Dim oStairsCon As Face
            oStairsCon = attrSets.Item(1).Parent.Parent
            Dim oStairsFaceProxy As FaceProxy
            oStairsFaceProxy = Nothing
            Call stairsOcc.CreateGeometryProxy(oStairsCon, oStairsFaceProxy)
            Call asmDoc.ComponentDefinition.Constraints.AddMateConstraint(oFloorFace, oStairsFaceProxy, 0)
        Catch ex As Exception

        End Try

EndRoutine:
        Show()
    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------------------


    '                         PLACING PLATFORMS IN AN ASSEMBLY



    '------------------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------------------




    Private Sub btnPlacePlatform_Click(sender As Object, e As EventArgs) Handles btnPlacePlatform.Click
        PlacePlatform()
    End Sub

    Private Sub PlacePlatform()
        Dim oDef As PartComponentDefinition
        Dim oTG As TransientGeometry
        Dim asmDoc As AssemblyDocument
        Dim oFloorFace As Face
        Dim oStairsDoc As PartDocument
        Dim attrSets As AttributeSetsEnumerator
        Dim stairsOcc As ComponentOccurrence
        Dim propSet1 As PropertySet
        Dim propSet2 As PropertySet
        Dim propSet3 As PropertySet

        'Check of Inventor nog steeds actief is
        Try
            Me.oInvApp = GetObject(, "Inventor.Application")
        Catch ex As Exception
            MsgBox("Inventor must be running!.", vbExclamation, "")
            btnPlaceStairsInAssem.Enabled = False
            btnPlaceLadder.Enabled = False
            btnPlaceRailing.Enabled = False
            btnPlacePlatform.Enabled = False
            btnPlaceCurvedStairs.Enabled = False
            Exit Sub
        End Try

        'Check of er een assembly open is
        If Me.oInvApp.ActiveDocumentType = Global.Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            GoTo BeginRoutine
        Else MsgBox("Open, Create or Activate an Assembly !", vbExclamation, "")
            GoTo EndRoutine
        End If

BeginRoutine:
        If Dir(lblPlatformPath.Text & oPlatformFilename) = "" Then
            GoTo StartPart
        Else
            GoTo PlacePart
        End If
StartPart:

        Try
            Me.oInvApp = GetObject(, "Inventor.Application")
        Catch ex As Exception
            MsgBox("Inventor must be running.")
            btnPlaceStairsInAssem.Enabled = False
            btnPlaceLadder.Enabled = False
            btnPlaceRailing.Enabled = False
            btnPlacePlatform.Enabled = False
            btnPlaceCurvedStairs.Enabled = False
            Exit Sub
        End Try

        ProgressBar4.Visible = True
        ProgressBar4.Value = 10




        Me.oInvApp.ScreenUpdating = False

        oInvApp = GetObject(, "Inventor.Application")
        oStairsDoc = oInvApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject)
        oStairsDoc = oInvApp.ActiveDocument
        oDef = oInvApp.ActiveDocument.ComponentDefinition
        oTG = oInvApp.TransientGeometry
        pi = Math.Acos(-1) 'radians van 45° hoek


        '-----------------------------------------------------------------------------Setting Properties------------------------------------------------- 
        ' Get the design tracking property set.

        propSet1 = oStairsDoc.PropertySets.Item("Inventor Summary Information")
        propSet1.ItemByPropId(2).Value = "Steel Platform"
        propSet1.ItemByPropId(4).Value = "Marc Crauwels"
        propSet1.ItemByPropId(6).Value = "This Steel Platform is generated with the Pantarein Water Panta Stairs application. This software is part of the Pantarein Water Design Suite."

        propSet2 = oStairsDoc.PropertySets.Item("Inventor Document Summary Information")
        propSet2.ItemByPropId(2).Value = "Plant Design Mechanical"
        propSet2.ItemByPropId(15).Value = "Pantarein Water BVBA"

        propSet3 = oStairsDoc.PropertySets.Item("Design Tracking Properties")
        propSet3.ItemByPropId(41).Value = "Marc Crauwels"
        propSet3.ItemByPropId(29).Value = "Steel Platform"
        propSet3.ItemByPropId(23).Value = "http://www.pantareinwater.be/en"




        Me.oInvApp.ScreenUpdating = True
        ProgressBar4.Value = 20
        Me.oInvApp.ScreenUpdating = False

        '-------------------------------------------------------------------------------------------------------------------------------------------------
        '----------Adjust Orientation XY is floor Z is up----------

        Dim v As Inventor.View
        v = oInvApp.ActiveView
        Dim c As Camera
        c = oInvApp.ActiveView.Camera
        Dim t2eDist As Double
        t2eDist = c.Target.DistanceTo(c.Eye)
        Dim t2e As Vector
        t2e = oTG.CreateVector(0, -t2eDist, 0)
        Dim newEye As Point
        newEye = c.Target.Copy
        Call newEye.TranslateBy(t2e)
        c.Eye = newEye
        c.UpVector = oTG.CreateUnitVector(0, 0, 1)
        c.ApplyWithoutTransition()
        v.SetCurrentAsFront()


        Me.oInvApp.ScreenUpdating = True
        ProgressBar4.Value = 30
        Me.oInvApp.ScreenUpdating = False


        '---------------------Rectangular Platform ----------------------------------

        If oPlatformState = "RECT_LEGS_FOOT" Then
            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)
            ' oGroundPlane.Name = "Bottom Of FootPlates"

            Dim oFootPlateSketch As PlanarSketch
            oFootPlateSketch = oDef.Sketches.Add(oGroundPlane)

            Dim oWp01 As SketchPoints
            oWp01 = oFootPlateSketch.SketchPoints





            '------------Sketch points Footplate----------------------

            Call oWp01.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) - 5, (-oPlatformWidth / 20) - 5), False)
            Call oWp01.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) - 5, (-oPlatformWidth / 20) + 19), False)
            Call oWp01.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 19, (-oPlatformWidth / 20) + 19), False)
            Call oWp01.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 19, (-oPlatformWidth / 20) - 5), False)

            '-------------Draw lines Footplate----------------------------------------------------

            Dim oLn01 As SketchLine
            oLn01 = oFootPlateSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
            Dim oLn02 As SketchLine
            oLn02 = oFootPlateSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
            Dim oLn03 As SketchLine
            oLn03 = oFootPlateSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
            Dim oLn04 As SketchLine
            oLn04 = oFootPlateSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(1))

            Me.oInvApp.ScreenUpdating = True
            ProgressBar4.Value = 40
            Me.oInvApp.ScreenUpdating = False

            '--------------Extrusion of the Footplate--------------------------------------


            Dim oFootPlateProfile As Profile
            oFootPlateProfile = oFootPlateSketch.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oFootPlateProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oExtrudeDef.SetDistanceExtent(2, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrude As ExtrudeFeature
            oExtrude = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)

            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(5)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------




            ' oExtrude.Name = "FootPlate"
            '--------------------------------------------------------------------------

            '------------Sketch points Column----------------------

            '--------------------------------------------------------------------------



            Dim oColumnSketch As PlanarSketch
            oColumnSketch = oDef.Sketches.Add(oGroundPlane)
            oGroundPlane.Visible = False
            Dim oWp02 As SketchPoints
            oWp02 = oColumnSketch.SketchPoints

            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20), (-oPlatformWidth / 20)), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20), (-oPlatformWidth / 20) + 14), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 1, (-oPlatformWidth / 20) + 14), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 1, (-oPlatformWidth / 20) + 7.5), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 13, (-oPlatformWidth / 20) + 7.5), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 13, (-oPlatformWidth / 20) + 14), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 14, (-oPlatformWidth / 20) + 14), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 14, (-oPlatformWidth / 20)), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 13, (-oPlatformWidth / 20)), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 13, (-oPlatformWidth / 20) + 6.5), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 1, (-oPlatformWidth / 20) + 6.5), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 1, (-oPlatformWidth / 20)), False)

            Me.oInvApp.ScreenUpdating = True
            ProgressBar4.Value = 50
            Me.oInvApp.ScreenUpdating = False

            '-------------Draw lines Footplate----------------------------------------------------

            Dim oLn05 As SketchLine
            oLn05 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(1), oWp02(2))
            Dim oLn06 As SketchLine
            oLn06 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(2), oWp02(3))
            Dim oLn07 As SketchLine
            oLn07 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(3), oWp02(4))
            Dim oLn08 As SketchLine
            oLn08 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(4), oWp02(5))
            Dim oLn09 As SketchLine
            oLn09 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(5), oWp02(6))
            Dim oLn10 As SketchLine
            oLn10 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(6), oWp02(7))
            Dim oLn11 As SketchLine
            oLn11 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(7), oWp02(8))
            Dim oLn12 As SketchLine
            oLn12 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(8), oWp02(9))
            Dim oLn13 As SketchLine
            oLn13 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(9), oWp02(10))
            Dim oLn14 As SketchLine
            oLn14 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(10), oWp02(11))
            Dim oLn15 As SketchLine
            oLn15 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(11), oWp02(12))
            Dim oLn16 As SketchLine
            oLn16 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(12), oWp02(1))

            Me.oInvApp.ScreenUpdating = True
            ProgressBar4.Value = 60
            Me.oInvApp.ScreenUpdating = False

            '--------------Extrusion of the Column--------------------------------------

            Dim oColumnProfile As Profile
            oColumnProfile = oColumnSketch.Profiles.AddForSolid
            Dim oColumnExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oColumnExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oColumnProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oColumnExtrudeDef.SetDistanceExtent(oPlatformHeight / 10 - 24, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oColumnExtrude As ExtrudeFeature
            oColumnExtrude = oDef.Features.ExtrudeFeatures.Add(oColumnExtrudeDef)
            '  oColumnExtrude.Name = "Column"
            '--------------------------------------------------------------------------


            '-------------------------------------Pattern the Column-------------------------------------------------
            Dim oCLMNCol As ObjectCollection

            ' Adding the extrusions, revolves and ... to the object Collection.

            oCLMNCol = (oExtrude.Definition.AffectedBodies)
            oCLMNCol = (oColumnExtrude.Definition.AffectedBodies)

            Dim oRPattern As RectangularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing

            ' A Rectangle Pattern in two directions
            Dim oLenglgsNr As Integer
            Dim oPatternDistL As Double
            If oPlatformLenght <= 4000 Then
                oLenglgsNr = 2
                oPatternDistL = oPlatformLenght / 10 - 14
            End If
            If oPlatformLenght >= 4000 Then
                oLenglgsNr = 3
                oPatternDistL = (oPlatformLenght / 10 - 14) / 2
            End If
            Dim oWidlgsNr As Integer
            Dim oPatternDistW As Double
            If oPlatformWidth <= 4000 Then
                oWidlgsNr = 2
                oPatternDistW = oPlatformWidth / 10 - 14
            End If
            If oPlatformWidth >= 4000 Then
                oPatternDistW = (oPlatformWidth / 10 - 14) / 2
                oWidlgsNr = 3
            End If
            oRPattern = oDef.Features.RectangularPatternFeatures.Add(oCLMNCol, oDef.WorkPlanes.Item(1), True, oLenglgsNr, oPatternDistL, kDefault, oDef.WorkPoints.Item(1), oDef.WorkPlanes.Item(2), True, oWidlgsNr, oPatternDistW, kDefault)
            ' oRPattern.Name = "Column Pattern"

            ProgressBar4.Value = 70

            Dim oTOSPlane As WorkPlane

            oTOSPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10 - 4, False)
            '  oTOSPlane.Name = "TOS Platform"

            Dim oTOSSketch As PlanarSketch
            oTOSSketch = oDef.Sketches.Add(oTOSPlane)

            Dim oWp03 As SketchPoints
            oWp03 = oTOSSketch.SketchPoints


            '------------Sketch points Platform----------------------

            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20, -oPlatformWidth / 20), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20, -oPlatformWidth / 20), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20, oPlatformWidth / 20), False)
            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20, oPlatformWidth / 20), False)

            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, oPlatformWidth / 20 - 6.5), False)
            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, oPlatformWidth / 20 - 6.5), False)

            '-------------Draw lines Platform----------------------------------------------------


            ProgressBar4.Value = 80

            Dim oLn17 As SketchLine
            oLn17 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(1), oWp03(2))
            Dim oLn18 As SketchLine
            oLn18 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(2), oWp03(3))
            Dim oLn19 As SketchLine
            oLn19 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(3), oWp03(4))
            Dim oLn20 As SketchLine
            oLn20 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(4), oWp03(1))
            Dim oLn21 As SketchLine
            oLn21 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(5), oWp03(6))
            Dim oLn22 As SketchLine
            oLn22 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(6), oWp03(7))
            Dim oLn23 As SketchLine
            oLn23 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(7), oWp03(8))
            Dim oLn24 As SketchLine
            oLn24 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(8), oWp03(5))

            '--------------Extrusion of the Platform steel--------------------------------------


            Dim oSteelProfile As Profile
            oSteelProfile = oTOSSketch.Profiles.AddForSolid
            Dim oSteelExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oSteelExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSteelProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSteelExtrudeDef.SetDistanceExtent(20, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oSteelExtrude As ExtrudeFeature
            oSteelExtrude = oDef.Features.ExtrudeFeatures.Add(oSteelExtrudeDef)
            ' oSteelExtrude.Name = "Steel Frame"
            '--------------------------------------------------------------------------

            Me.oInvApp.ScreenUpdating = True
            ProgressBar4.Value = 90
            Me.oInvApp.ScreenUpdating = False

            Dim oTOGPlane As WorkPlane

            oTOGPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10, False)
            '  oTOGPlane.Name = "TOG"

            Dim oTOGSketch As PlanarSketch
            oTOGSketch = oDef.Sketches.Add(oTOGPlane)

            Dim oWp04 As SketchPoints
            oWp04 = oTOGSketch.SketchPoints


            '------------Sketch points Platform----------------------



            Call oWp04.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, oPlatformWidth / 20 - 6.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, oPlatformWidth / 20 - 6.5), False)

            '-------------Draw lines Platform----------------------------------------------------




            Dim oLn25 As SketchLine
            oLn25 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(1), oWp04(2))
            Dim oLn26 As SketchLine
            oLn26 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(2), oWp04(3))
            Dim oLn27 As SketchLine
            oLn27 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(3), oWp04(4))
            Dim oLn28 As SketchLine
            oLn28 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(4), oWp04(1))


            ProgressBar4.Value = 100

            '--------------Extrusion of the Grating--------------------------------------


            Dim oGrateProfile As Profile
            oGrateProfile = oTOGSketch.Profiles.AddForSolid
            Dim oGrateExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oGrateExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oGrateProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oGrateExtrudeDef.SetDistanceExtent(4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oGrateExtrude As ExtrudeFeature
            oGrateExtrude = oDef.Features.ExtrudeFeatures.Add(oGrateExtrudeDef)
            ' oGrateExtrude.Name = "Grating"
            '--------------------------------------------------------------------------
            oInvApp.ScreenUpdating = True
            ProgressBar4.Value = 100
            oInvApp.ScreenUpdating = False

            oGroundPlane.Visible = False
            oTOSPlane.Visible = False
            oTOGPlane.Visible = False

        End If

        If oPlatformState = "RECT_LEGS" Then
            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)


            oInvApp.ScreenUpdating = True
            ProgressBar4.Value = 30
            oInvApp.ScreenUpdating = False



            '--------------------------------------------------------------------------

            '------------Sketch points Column----------------------

            '--------------------------------------------------------------------------

            Dim oColumnSketch As PlanarSketch
            oColumnSketch = oDef.Sketches.Add(oGroundPlane)
            oGroundPlane.Visible = False
            Dim oWp02 As SketchPoints
            oWp02 = oColumnSketch.SketchPoints

            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20), (-oPlatformWidth / 20)), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20), (-oPlatformWidth / 20) + 14), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 1, (-oPlatformWidth / 20) + 14), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 1, (-oPlatformWidth / 20) + 7.5), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 13, (-oPlatformWidth / 20) + 7.5), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 13, (-oPlatformWidth / 20) + 14), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 14, (-oPlatformWidth / 20) + 14), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 14, (-oPlatformWidth / 20)), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 13, (-oPlatformWidth / 20)), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 13, (-oPlatformWidth / 20) + 6.5), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 1, (-oPlatformWidth / 20) + 6.5), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformLenght / 20) + 1, (-oPlatformWidth / 20)), False)

            ProgressBar4.Value = 50

            '-------------Draw lines Footplate----------------------------------------------------

            Dim oLn05 As SketchLine
            oLn05 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(1), oWp02(2))
            Dim oLn06 As SketchLine
            oLn06 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(2), oWp02(3))
            Dim oLn07 As SketchLine
            oLn07 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(3), oWp02(4))
            Dim oLn08 As SketchLine
            oLn08 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(4), oWp02(5))
            Dim oLn09 As SketchLine
            oLn09 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(5), oWp02(6))
            Dim oLn10 As SketchLine
            oLn10 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(6), oWp02(7))
            Dim oLn11 As SketchLine
            oLn11 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(7), oWp02(8))
            Dim oLn12 As SketchLine
            oLn12 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(8), oWp02(9))
            Dim oLn13 As SketchLine
            oLn13 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(9), oWp02(10))
            Dim oLn14 As SketchLine
            oLn14 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(10), oWp02(11))
            Dim oLn15 As SketchLine
            oLn15 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(11), oWp02(12))
            Dim oLn16 As SketchLine
            oLn16 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(12), oWp02(1))

            ProgressBar4.Value = 60

            '--------------Extrusion of the Column--------------------------------------

            Dim oColumnProfile As Profile
            oColumnProfile = oColumnSketch.Profiles.AddForSolid
            Dim oColumnExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oColumnExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oColumnProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oColumnExtrudeDef.SetDistanceExtent(oPlatformHeight / 10 - 24, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oColumnExtrude As ExtrudeFeature
            oColumnExtrude = oDef.Features.ExtrudeFeatures.Add(oColumnExtrudeDef)


            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(13)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------




            ' oColumnExtrude.Name = "Column"
            '--------------------------------------------------------------------------


            '-------------------------------------Pattern the Column-------------------------------------------------
            Dim oCLMNCol As ObjectCollection

            ' Adding the extrusions, revolves and ... to the object Collection.


            oCLMNCol = (oColumnExtrude.Definition.AffectedBodies)

            Dim oRPattern As RectangularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing

            ' A Rectangle Pattern in two directions
            Dim oLenglgsNr As Integer
            Dim oPatternDistL As Double
            If oPlatformLenght <= 4000 Then
                oLenglgsNr = 2
                oPatternDistL = oPlatformLenght / 10 - 14
            End If
            If oPlatformLenght >= 4000 Then
                oLenglgsNr = 3
                oPatternDistL = (oPlatformLenght / 10 - 14) / 2
            End If
            Dim oWidlgsNr As Integer
            Dim oPatternDistW As Double
            If oPlatformWidth <= 4000 Then
                oWidlgsNr = 2
                oPatternDistW = oPlatformWidth / 10 - 14
            End If
            If oPlatformWidth >= 4000 Then
                oPatternDistW = (oPlatformWidth / 10 - 14) / 2
                oWidlgsNr = 3
            End If
            oRPattern = oDef.Features.RectangularPatternFeatures.Add(oCLMNCol, oDef.WorkPlanes.Item(1), True, oLenglgsNr, oPatternDistL, kDefault, oDef.WorkPoints.Item(1), oDef.WorkPlanes.Item(2), True, oWidlgsNr, oPatternDistW, kDefault)
            ' oRPattern.Name = "Column Pattern"

            ProgressBar4.Value = 70

            Dim oTOSPlane As WorkPlane

            oTOSPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10 - 4, False)
            ' oTOSPlane.Name = "TOS Platform"

            Dim oTOSSketch As PlanarSketch
            oTOSSketch = oDef.Sketches.Add(oTOSPlane)

            Dim oWp03 As SketchPoints
            oWp03 = oTOSSketch.SketchPoints


            '------------Sketch points Platform----------------------

            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20, -oPlatformWidth / 20), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20, -oPlatformWidth / 20), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20, oPlatformWidth / 20), False)
            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20, oPlatformWidth / 20), False)

            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, oPlatformWidth / 20 - 6.5), False)
            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, oPlatformWidth / 20 - 6.5), False)

            '-------------Draw lines Platform----------------------------------------------------




            Dim oLn17 As SketchLine
            oLn17 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(1), oWp03(2))
            Dim oLn18 As SketchLine
            oLn18 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(2), oWp03(3))
            Dim oLn19 As SketchLine
            oLn19 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(3), oWp03(4))
            Dim oLn20 As SketchLine
            oLn20 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(4), oWp03(1))
            Dim oLn21 As SketchLine
            oLn21 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(5), oWp03(6))
            Dim oLn22 As SketchLine
            oLn22 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(6), oWp03(7))
            Dim oLn23 As SketchLine
            oLn23 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(7), oWp03(8))
            Dim oLn24 As SketchLine
            oLn24 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(8), oWp03(5))

            '--------------Extrusion of the Platform steel--------------------------------------


            Dim oSteelProfile As Profile
            oSteelProfile = oTOSSketch.Profiles.AddForSolid
            Dim oSteelExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oSteelExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSteelProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSteelExtrudeDef.SetDistanceExtent(20, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oSteelExtrude As ExtrudeFeature
            oSteelExtrude = oDef.Features.ExtrudeFeatures.Add(oSteelExtrudeDef)
            ' oSteelExtrude.Name = "Steel Frame"
            '--------------------------------------------------------------------------

            ProgressBar4.Value = 80

            Dim oTOGPlane As WorkPlane

            oTOGPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10, False)
            '  oTOGPlane.Name = "TOG"

            Dim oTOGSketch As PlanarSketch
            oTOGSketch = oDef.Sketches.Add(oTOGPlane)

            Dim oWp04 As SketchPoints
            oWp04 = oTOGSketch.SketchPoints


            '------------Sketch points Platform----------------------



            Call oWp04.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, oPlatformWidth / 20 - 6.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, oPlatformWidth / 20 - 6.5), False)

            '-------------Draw lines Platform----------------------------------------------------




            Dim oLn25 As SketchLine
            oLn25 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(1), oWp04(2))
            Dim oLn26 As SketchLine
            oLn26 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(2), oWp04(3))
            Dim oLn27 As SketchLine
            oLn27 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(3), oWp04(4))
            Dim oLn28 As SketchLine
            oLn28 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(4), oWp04(1))


            ProgressBar4.Value = 90

            '--------------Extrusion of the Grating--------------------------------------


            Dim oGrateProfile As Profile
            oGrateProfile = oTOGSketch.Profiles.AddForSolid
            Dim oGrateExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oGrateExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oGrateProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oGrateExtrudeDef.SetDistanceExtent(4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oGrateExtrude As ExtrudeFeature
            oGrateExtrude = oDef.Features.ExtrudeFeatures.Add(oGrateExtrudeDef)
            ' oGrateExtrude.Name = "Grating"
            '--------------------------------------------------------------------------
            ProgressBar4.Value = 100

            oGroundPlane.Visible = False
            oTOSPlane.Visible = False
            oTOGPlane.Visible = False

        End If






        If oPlatformState = "RECT" Then
            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)








            ProgressBar4.Value = 70

            Dim oTOSPlane As WorkPlane

            oTOSPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10 - 4, False)
            oTOSPlane.Name = "TOS Platform"

            Dim oTOSSketch As PlanarSketch
            oTOSSketch = oDef.Sketches.Add(oTOSPlane)

            Dim oWp03 As SketchPoints
            oWp03 = oTOSSketch.SketchPoints


            '------------Sketch points Platform----------------------

            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20, -oPlatformWidth / 20), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20, -oPlatformWidth / 20), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20, oPlatformWidth / 20), False)
            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20, oPlatformWidth / 20), False)

            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp03.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, oPlatformWidth / 20 - 6.5), False)
            Call oWp03.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, oPlatformWidth / 20 - 6.5), False)

            '-------------Draw lines Platform----------------------------------------------------




            Dim oLn17 As SketchLine
            oLn17 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(1), oWp03(2))
            Dim oLn18 As SketchLine
            oLn18 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(2), oWp03(3))
            Dim oLn19 As SketchLine
            oLn19 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(3), oWp03(4))
            Dim oLn20 As SketchLine
            oLn20 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(4), oWp03(1))
            Dim oLn21 As SketchLine
            oLn21 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(5), oWp03(6))
            Dim oLn22 As SketchLine
            oLn22 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(6), oWp03(7))
            Dim oLn23 As SketchLine
            oLn23 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(7), oWp03(8))
            Dim oLn24 As SketchLine
            oLn24 = oTOSSketch.SketchLines.AddByTwoPoints(oWp03(8), oWp03(5))

            '--------------Extrusion of the Platform steel--------------------------------------


            Dim oSteelProfile As Profile
            oSteelProfile = oTOSSketch.Profiles.AddForSolid
            Dim oSteelExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oSteelExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSteelProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSteelExtrudeDef.SetDistanceExtent(20, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oSteelExtrude As ExtrudeFeature
            oSteelExtrude = oDef.Features.ExtrudeFeatures.Add(oSteelExtrudeDef)

            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(10)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------




            '  oSteelExtrude.Name = "Steel Frame"
            '--------------------------------------------------------------------------

            ProgressBar4.Value = 80

            Dim oTOGPlane As WorkPlane

            oTOGPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10, False)
            oTOGPlane.Name = "TOG"

            Dim oTOGSketch As PlanarSketch
            oTOGSketch = oDef.Sketches.Add(oTOGPlane)

            Dim oWp04 As SketchPoints
            oWp04 = oTOGSketch.SketchPoints


            '------------Sketch points Platform----------------------



            Call oWp04.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, -oPlatformWidth / 20 + 6.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(oPlatformLenght / 20 - 6.5, oPlatformWidth / 20 - 6.5), False)
            Call oWp04.Add(oTG.CreatePoint2d(-oPlatformLenght / 20 + 6.5, oPlatformWidth / 20 - 6.5), False)

            '-------------Draw lines Platform----------------------------------------------------




            Dim oLn25 As SketchLine
            oLn25 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(1), oWp04(2))
            Dim oLn26 As SketchLine
            oLn26 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(2), oWp04(3))
            Dim oLn27 As SketchLine
            oLn27 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(3), oWp04(4))
            Dim oLn28 As SketchLine
            oLn28 = oTOGSketch.SketchLines.AddByTwoPoints(oWp04(4), oWp04(1))


            ProgressBar4.Value = 90

            '--------------Extrusion of the Grating--------------------------------------


            Dim oGrateProfile As Profile
            oGrateProfile = oTOGSketch.Profiles.AddForSolid
            Dim oGrateExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oGrateExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oGrateProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oGrateExtrudeDef.SetDistanceExtent(4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oGrateExtrude As ExtrudeFeature
            oGrateExtrude = oDef.Features.ExtrudeFeatures.Add(oGrateExtrudeDef)
            '  oGrateExtrude.Name = "Grating"
            '--------------------------------------------------------------------------
            ProgressBar4.Value = 100

            oGroundPlane.Visible = False
            oTOSPlane.Visible = False
            oTOGPlane.Visible = False

        End If



        If oPlatformState = "CIRC_LEGS_FOOT" Then
            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)
            oGroundPlane.Name = "Bottom Of FootPlates"

            Dim oFootPlateSketch As PlanarSketch
            oFootPlateSketch = oDef.Sketches.Add(oGroundPlane)

            Dim oWp01 As SketchPoints
            oWp01 = oFootPlateSketch.SketchPoints


            ProgressBar4.Value = 40

            '------------Sketch points Footplate----------------------

            Call oWp01.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) - 5, -12), False)
            Call oWp01.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) - 5, +12), False)
            Call oWp01.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20 + 19), +12), False)
            Call oWp01.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20 + 19), -12), False)

            '-------------Draw lines Footplate----------------------------------------------------

            Dim oLn01 As SketchLine
            oLn01 = oFootPlateSketch.SketchLines.AddByTwoPoints(oWp01(1), oWp01(2))
            Dim oLn02 As SketchLine
            oLn02 = oFootPlateSketch.SketchLines.AddByTwoPoints(oWp01(2), oWp01(3))
            Dim oLn03 As SketchLine
            oLn03 = oFootPlateSketch.SketchLines.AddByTwoPoints(oWp01(3), oWp01(4))
            Dim oLn04 As SketchLine
            oLn04 = oFootPlateSketch.SketchLines.AddByTwoPoints(oWp01(4), oWp01(1))

            ProgressBar4.Value = 50

            '--------------Extrusion of the Footplate--------------------------------------


            Dim oFootPlateProfile As Profile
            oFootPlateProfile = oFootPlateSketch.Profiles.AddForSolid
            Dim oExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oFootPlateProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oExtrudeDef.SetDistanceExtent(2, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oExtrude As ExtrudeFeature
            oExtrude = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)

            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(5)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------


            '  oExtrude.Name = "FootPlate"
            '--------------------------------------------------------------------------

            '------------Sketch points Column----------------------

            '--------------------------------------------------------------------------

            Dim oColumnSketch As PlanarSketch
            oColumnSketch = oDef.Sketches.Add(oGroundPlane)
            oGroundPlane.Visible = False
            Dim oWp02 As SketchPoints
            oWp02 = oColumnSketch.SketchPoints

            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20), -7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20), +7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 1, +7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 1, +0.25), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 13, +0.25), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 13, +7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 14, +7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 14, -7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 13, -7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 13, -0.25), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 1, -0.25), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 1, -7), False)

            '-------------Draw lines Footplate----------------------------------------------------

            Dim oLn05 As SketchLine
            oLn05 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(1), oWp02(2))
            Dim oLn06 As SketchLine
            oLn06 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(2), oWp02(3))
            Dim oLn07 As SketchLine
            oLn07 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(3), oWp02(4))
            Dim oLn08 As SketchLine
            oLn08 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(4), oWp02(5))
            Dim oLn09 As SketchLine
            oLn09 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(5), oWp02(6))
            Dim oLn10 As SketchLine
            oLn10 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(6), oWp02(7))
            Dim oLn11 As SketchLine
            oLn11 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(7), oWp02(8))
            Dim oLn12 As SketchLine
            oLn12 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(8), oWp02(9))
            Dim oLn13 As SketchLine
            oLn13 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(9), oWp02(10))
            Dim oLn14 As SketchLine
            oLn14 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(10), oWp02(11))
            Dim oLn15 As SketchLine
            oLn15 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(11), oWp02(12))
            Dim oLn16 As SketchLine
            oLn16 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(12), oWp02(1))

            ProgressBar4.Value = 60

            '--------------Extrusion of the Column--------------------------------------

            Dim oColumnProfile As Profile
            oColumnProfile = oColumnSketch.Profiles.AddForSolid
            Dim oColumnExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oColumnExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oColumnProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oColumnExtrudeDef.SetDistanceExtent(oPlatformHeight / 10 - 24, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oColumnExtrude As ExtrudeFeature
            oColumnExtrude = oDef.Features.ExtrudeFeatures.Add(oColumnExtrudeDef)
            '  oColumnExtrude.Name = "Column"
            '--------------------------------------------------------------------------


            '-------------------------------------Pattern the Column-------------------------------------------------
            Dim oCLMNCol As ObjectCollection

            ' Adding the extrusions, revolves and ... to the object Collection.

            oCLMNCol = (oExtrude.Definition.AffectedBodies)
            oCLMNCol = (oColumnExtrude.Definition.AffectedBodies)
            Dim oRPattern As CircularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing

            ' A Rectangle Pattern in two directions

            Dim oDiaNr As Integer

            If oPlatformDiameter <= 5000 Then
                oDiaNr = 4
            End If
            If CboPlatformLegNumber.Text = "6" Then
                oDiaNr = 6
            End If
            If CboPlatformLegNumber.Text = "8" Then
                oDiaNr = 8
            End If
            If oPlatformDiameter > 5000 And oPlatformDiameter <= 7500 Then
                oDiaNr = 6
            End If
            If CboPlatformLegNumber.Text = "8" Then
                oDiaNr = 8
            End If
            If oPlatformDiameter > 7500 Then
                oDiaNr = 8
            End If

            oRPattern = oDef.Features.CircularPatternFeatures.Add(oCLMNCol, oDef.WorkAxes.Item(3), False, oDiaNr, 360 * pi / 180, True,)
            ' oRPattern.Name = "Column Pattern"




            ProgressBar4.Value = 70

            Dim oTOSPlane As WorkPlane
            oTOSPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10 - 4, False)
            oTOSPlane.Name = "TOS Platform"
            Dim oTOSSketch As PlanarSketch
            oTOSSketch = oDef.Sketches.Add(oTOSPlane)
            Dim oCircle1 As SketchCircle
            Dim oCircle2 As SketchCircle
            oCircle1 = oTOSSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20)
            oCircle2 = oTOSSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20 - 6.5)

            '--------------Extrusion of the Platform steel--------------------------------------

            Dim oSteelProfile As Profile
            oSteelProfile = oTOSSketch.Profiles.AddForSolid
            Dim oSteelExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oSteelExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSteelProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSteelExtrudeDef.SetDistanceExtent(20, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oSteelExtrude As ExtrudeFeature
            oSteelExtrude = oDef.Features.ExtrudeFeatures.Add(oSteelExtrudeDef)
            '  oSteelExtrude.Name = "Steel Frame"
            '--------------------------------------------------------------------------

            ProgressBar4.Value = 80

            Dim oTOGPlane As WorkPlane
            oTOGPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10, False)
            oTOGPlane.Name = "TOG"
            Dim oTOGSketch As PlanarSketch
            oTOGSketch = oDef.Sketches.Add(oTOGPlane)
            Dim oCircle3 As SketchCircle
            oCircle3 = oTOGSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20 - 4)


            ProgressBar4.Value = 90

            '--------------Extrusion of the Grating--------------------------------------


            Dim oGrateProfile As Profile
            oGrateProfile = oTOGSketch.Profiles.AddForSolid
            Dim oGrateExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oGrateExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oGrateProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oGrateExtrudeDef.SetDistanceExtent(4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oGrateExtrude As ExtrudeFeature
            oGrateExtrude = oDef.Features.ExtrudeFeatures.Add(oGrateExtrudeDef)
            '  oGrateExtrude.Name = "Grating"
            '--------------------------------------------------------------------------
            ProgressBar4.Value = 100

            oGroundPlane.Visible = False
            oTOSPlane.Visible = False
            oTOGPlane.Visible = False

        End If

        If oPlatformState = "CIRC_LEGS" Then
            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)

            '--------------------------------------------------------------------------

            '------------Sketch points Column----------------------

            '--------------------------------------------------------------------------

            Dim oColumnSketch As PlanarSketch
            oColumnSketch = oDef.Sketches.Add(oGroundPlane)
            oGroundPlane.Visible = False
            Dim oWp02 As SketchPoints
            oWp02 = oColumnSketch.SketchPoints

            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20), -7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20), +7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 1, +7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 1, +0.25), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 13, +0.25), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 13, +7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 14, +7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 14, -7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 13, -7), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 13, -0.25), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 1, -0.25), False)
            Call oWp02.Add(oTG.CreatePoint2d((-oPlatformDiameter / 20) + 1, -7), False)

            '-------------Draw lines Footplate----------------------------------------------------

            Dim oLn05 As SketchLine
            oLn05 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(1), oWp02(2))
            Dim oLn06 As SketchLine
            oLn06 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(2), oWp02(3))
            Dim oLn07 As SketchLine
            oLn07 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(3), oWp02(4))
            Dim oLn08 As SketchLine
            oLn08 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(4), oWp02(5))
            Dim oLn09 As SketchLine
            oLn09 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(5), oWp02(6))
            Dim oLn10 As SketchLine
            oLn10 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(6), oWp02(7))
            Dim oLn11 As SketchLine
            oLn11 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(7), oWp02(8))
            Dim oLn12 As SketchLine
            oLn12 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(8), oWp02(9))
            Dim oLn13 As SketchLine
            oLn13 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(9), oWp02(10))
            Dim oLn14 As SketchLine
            oLn14 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(10), oWp02(11))
            Dim oLn15 As SketchLine
            oLn15 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(11), oWp02(12))
            Dim oLn16 As SketchLine
            oLn16 = oColumnSketch.SketchLines.AddByTwoPoints(oWp02(12), oWp02(1))

            ProgressBar4.Value = 60

            '--------------Extrusion of the Column--------------------------------------

            Dim oColumnProfile As Profile
            oColumnProfile = oColumnSketch.Profiles.AddForSolid
            Dim oColumnExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oColumnExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oColumnProfile, PartFeatureOperationEnum.kJoinOperation)
            Call oColumnExtrudeDef.SetDistanceExtent(oPlatformHeight / 10 - 24, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
            Dim oColumnExtrude As ExtrudeFeature
            oColumnExtrude = oDef.Features.ExtrudeFeatures.Add(oColumnExtrudeDef)


            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(13)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------



            ' oColumnExtrude.Name = "Column"
            '--------------------------------------------------------------------------


            '-------------------------------------Pattern the Column-------------------------------------------------
            Dim oCLMNCol As ObjectCollection

            ' Adding the extrusions, revolves and ... to the object Collection.


            oCLMNCol = (oColumnExtrude.Definition.AffectedBodies)
            Dim oRPattern As CircularPatternFeature
            Dim kDefault As PatternSpacingTypeEnum = Nothing

            ' A Rectangle Pattern in two directions



            If oPlatformDiameter <= 5000 Then
                oDiaNr = 4
            End If
            If CboPlatformLegNumber.Text = "6" Then
                oDiaNr = 6
            End If
            If CboPlatformLegNumber.Text = "8" Then
                oDiaNr = 8
            End If
            If oPlatformDiameter > 5000 And oPlatformDiameter <= 7500 Then
                oDiaNr = 6
            End If
            If CboPlatformLegNumber.Text = "8" Then
                oDiaNr = 8
            End If
            If oPlatformDiameter > 7500 Then
                oDiaNr = 8
            End If



            oRPattern = oDef.Features.CircularPatternFeatures.Add(oCLMNCol, oDef.WorkAxes.Item(3), False, oDiaNr, 360 * pi / 180, True,)
            '  oRPattern.Name = "Column Pattern"




            ProgressBar4.Value = 70

            Dim oTOSPlane As WorkPlane
            oTOSPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10 - 4, False)
            '  oTOSPlane.Name = "TOS Platform"
            Dim oTOSSketch As PlanarSketch
            oTOSSketch = oDef.Sketches.Add(oTOSPlane)
            Dim oCircle1 As SketchCircle
            Dim oCircle2 As SketchCircle
            oCircle1 = oTOSSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20)
            oCircle2 = oTOSSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20 - 6.5)

            '--------------Extrusion of the Platform steel--------------------------------------

            Dim oSteelProfile As Profile
            oSteelProfile = oTOSSketch.Profiles.AddForSolid
            Dim oSteelExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oSteelExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSteelProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSteelExtrudeDef.SetDistanceExtent(20, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oSteelExtrude As ExtrudeFeature
            oSteelExtrude = oDef.Features.ExtrudeFeatures.Add(oSteelExtrudeDef)
            ' oSteelExtrude.Name = "Steel Frame"
            '--------------------------------------------------------------------------

            ProgressBar4.Value = 80

            Dim oTOGPlane As WorkPlane
            oTOGPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10, False)
            ' oTOGPlane.Name = "TOG"
            Dim oTOGSketch As PlanarSketch
            oTOGSketch = oDef.Sketches.Add(oTOGPlane)
            Dim oCircle3 As SketchCircle
            oCircle3 = oTOGSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20 - 4)


            ProgressBar4.Value = 90

            '--------------Extrusion of the Grating--------------------------------------


            Dim oGrateProfile As Profile
            oGrateProfile = oTOGSketch.Profiles.AddForSolid
            Dim oGrateExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oGrateExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oGrateProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oGrateExtrudeDef.SetDistanceExtent(4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oGrateExtrude As ExtrudeFeature
            oGrateExtrude = oDef.Features.ExtrudeFeatures.Add(oGrateExtrudeDef)
            ' oGrateExtrude.Name = "Grating"
            '--------------------------------------------------------------------------
            ProgressBar4.Value = 100

            oGroundPlane.Visible = False
            oTOSPlane.Visible = False
            oTOGPlane.Visible = False

        End If




        If oPlatformState = "CIRC" Then
            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)





            ProgressBar4.Value = 70

            Dim oTOSPlane As WorkPlane
            oTOSPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10 - 4, False)
            oTOSPlane.Name = "TOS Platform"
            Dim oTOSSketch As PlanarSketch
            oTOSSketch = oDef.Sketches.Add(oTOSPlane)
            Dim oCircle1 As SketchCircle
            Dim oCircle2 As SketchCircle
            oCircle1 = oTOSSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20)
            oCircle2 = oTOSSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20 - 6.5)

            '--------------Extrusion of the Platform steel--------------------------------------

            Dim oSteelProfile As Profile
            oSteelProfile = oTOSSketch.Profiles.AddForSolid
            Dim oSteelExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oSteelExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSteelProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSteelExtrudeDef.SetDistanceExtent(20, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oSteelExtrude As ExtrudeFeature
            oSteelExtrude = oDef.Features.ExtrudeFeatures.Add(oSteelExtrudeDef)


            'start---------Maak een AttributeSet "ConnectFace" ----------

            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(3)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------



            '  oSteelExtrude.Name = "Steel Frame"
            '--------------------------------------------------------------------------

            ProgressBar4.Value = 80

            Dim oTOGPlane As WorkPlane
            oTOGPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10, False)
            '  oTOGPlane.Name = "TOG"
            Dim oTOGSketch As PlanarSketch
            oTOGSketch = oDef.Sketches.Add(oTOGPlane)
            Dim oCircle3 As SketchCircle
            oCircle3 = oTOGSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20 - 4)


            ProgressBar4.Value = 90

            '--------------Extrusion of the Grating--------------------------------------


            Dim oGrateProfile As Profile
            oGrateProfile = oTOGSketch.Profiles.AddForSolid
            Dim oGrateExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oGrateExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oGrateProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oGrateExtrudeDef.SetDistanceExtent(4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oGrateExtrude As ExtrudeFeature
            oGrateExtrude = oDef.Features.ExtrudeFeatures.Add(oGrateExtrudeDef)
            '  oGrateExtrude.Name = "Grating"
            '--------------------------------------------------------------------------
            ProgressBar4.Value = 100

            oGroundPlane.Visible = False
            oTOSPlane.Visible = False
            oTOGPlane.Visible = False

        End If


        If oPlatformState = "CURVED" Then
            Dim oGroundPlane As WorkPlane
            oGroundPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), False)

            ProgressBar4.Value = 70

            Dim oTOSPlane As WorkPlane
            oTOSPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10 - 4, False)
            '  oTOSPlane.Name = "TOS Platform"
            Dim oTOSSketch As PlanarSketch
            oTOSSketch = oDef.Sketches.Add(oTOSPlane)
            Dim oLine01 As SketchLine
            Dim oLine02 As SketchLine
            Dim oArc1 As SketchArc
            Dim oArc2 As SketchArc

            oArc1 = oTOSSketch.SketchArcs.AddByCenterStartSweepAngle(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20, 0, oPlatformAngle * pi / 180)
            oArc2 = oTOSSketch.SketchArcs.AddByCenterStartSweepAngle(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20 - oPlatformWidth / 10, 0, oPlatformAngle * pi / 180)
            oLine01 = oTOSSketch.SketchLines.AddByTwoPoints(oArc1.StartSketchPoint, oArc2.StartSketchPoint)
            oLine02 = oTOSSketch.SketchLines.AddByTwoPoints(oArc1.EndSketchPoint, oArc2.EndSketchPoint)
            Dim oCollection As ObjectCollection
            oCollection = oInvApp.TransientObjects.CreateObjectCollection
            oCollection.Add(oLine01)
            Dim oOffsetPoint As Point2d
            oOffsetPoint = oTG.CreatePoint2d(oPlatformDiameter / 20, 6.5)
            Call oTOSSketch.OffsetSketchEntitiesUsingDistance(oCollection, 6, True, True,)

            '--------------Extrusion of the Platform steel--------------------------------------

            Dim oSteelProfile As Profile
            oSteelProfile = oTOSSketch.Profiles.AddForSolid
            Dim oSteelExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oSteelExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oSteelProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oSteelExtrudeDef.SetDistanceExtent(20, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oSteelExtrude As ExtrudeFeature
            oSteelExtrude = oDef.Features.ExtrudeFeatures.Add(oSteelExtrudeDef)


            Dim entity As Face
            entity = oDef.SurfaceBodies.Item(1).Faces.Item(10)
            Dim attribSets As AttributeSets
            attribSets = entity.AttributeSets
            Dim attribSet As AttributeSet
            attribSet = attribSets.Add("ConnectFace")

            'end---------Maak een AttributeSet "ConnectFace" ----------



            ' oSteelExtrude.Name = "Steel Frame"
            '--------------------------------------------------------------------------

            ProgressBar4.Value = 80

            Dim oTOGPlane As WorkPlane
            oTOGPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes.Item(3), oPlatformHeight / 10, False)
            '   oTOGPlane.Name = "TOG Platform"
            Dim oTOGSketch As PlanarSketch
            oTOGSketch = oDef.Sketches.Add(oTOGPlane)
            Dim oLine03 As SketchLine
            Dim oLine04 As SketchLine
            Dim oArc3 As SketchArc
            Dim oArc4 As SketchArc

            oArc3 = oTOGSketch.SketchArcs.AddByCenterStartSweepAngle(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20, 0, oPlatformAngle * pi / 180)
            oArc4 = oTOGSketch.SketchArcs.AddByCenterStartSweepAngle(oTG.CreatePoint2d(0, 0), oPlatformDiameter / 20 - oPlatformWidth / 10, 0, oPlatformAngle * pi / 180)
            oLine03 = oTOGSketch.SketchLines.AddByTwoPoints(oArc3.StartSketchPoint, oArc4.StartSketchPoint)
            oLine04 = oTOGSketch.SketchLines.AddByTwoPoints(oArc3.EndSketchPoint, oArc4.EndSketchPoint)
            Dim oCollection2 As ObjectCollection
            oCollection2 = oInvApp.TransientObjects.CreateObjectCollection
            oCollection2.Add(oLine03)
            Dim oOffsetPoint2 As Point2d
            oOffsetPoint2 = oTG.CreatePoint2d(oPlatformDiameter / 20, 4)
            Call oTOGSketch.OffsetSketchEntitiesUsingDistance(oCollection2, 4, True, True,)

            oLine03.Delete()
            oLine04.Delete()
            oArc3.Delete()
            oArc4.Delete()



            ProgressBar4.Value = 90

            '--------------Extrusion of the Grating--------------------------------------


            Dim oGrateProfile As Profile
            oGrateProfile = oTOGSketch.Profiles.AddForSolid
            Dim oGrateExtrudeDef As ExtrudeDefinition
            'Dim kJoinOperation As Object = Nothing
            oGrateExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oGrateProfile, PartFeatureOperationEnum.kNewBodyOperation)
            Call oGrateExtrudeDef.SetDistanceExtent(4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
            Dim oGrateExtrude As ExtrudeFeature
            oGrateExtrude = oDef.Features.ExtrudeFeatures.Add(oGrateExtrudeDef)
            '  oGrateExtrude.Name = "Grating"
            '--------------------------------------------------------------------------
            ProgressBar4.Value = 100

            oGroundPlane.Visible = False
            oTOSPlane.Visible = False
            oTOGPlane.Visible = False

        End If

        'Hide()

        'Return view to Home view
        oInvApp.CommandManager.ControlDefinitions.Item("AppViewCubeHomeCmd").Execute()
        'zoom all
        oInvApp.ActiveView.GoHome()
        ProgressBar4.Value = 0
        oStairsDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits
        'Show()

        Try
            oStairsDoc.SaveAs(lblPlatformPath.Text & oPlatformFilename, False)
        Catch ex As Exception
            MsgBox("Could not save the part." & vbCrLf & "Try again or check your credentials", vbOKOnly + "4064", "Warning")
            oInvApp.ScreenUpdating = True
            GoTo EndRoutine
        End Try



        oStairsDoc.Close()
        oInvApp.ScreenUpdating = True








        '--------------Place the part interactive--------------------------------------
PlacePart:

        ' Get the active assembly.

        asmDoc = oInvApp.ActiveDocument

        ' If assembly empty place stairs grounded at origin.

        If asmDoc.ComponentDefinition.ImmediateReferencedDefinitions.Count < 1 Then
            Dim trans As Matrix
            trans = oInvApp.TransientGeometry.CreateMatrix
            stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblPlatformPath.Text & oPlatformFilename), trans)
            stairsOcc.Grounded = True
            GoTo EndRoutine
        End If

        'Else place stairs on a floor surface.

        Hide()
        oFloorFace = oInvApp.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, "Select a floor to place the stairs.")
        If oFloorFace Is Nothing Then
            GoTo EndRoutine
        End If

        stairsOcc = asmDoc.ComponentDefinition.Occurrences.Add((lblPlatformPath.Text & oPlatformFilename), oInvApp.TransientGeometry.CreateMatrix)
        oStairsDoc = stairsOcc.Definition.Document

        Try
            attrSets = oStairsDoc.AttributeManager.FindAttributeSets("ConnectFace")
            Dim oStairsCon As Face
            oStairsCon = attrSets.Item(1).Parent.Parent
            Dim oStairsFaceProxy As FaceProxy
            oStairsFaceProxy = Nothing
            Call stairsOcc.CreateGeometryProxy(oStairsCon, oStairsFaceProxy)
            Call asmDoc.ComponentDefinition.Constraints.AddMateConstraint(oFloorFace, oStairsFaceProxy, 0)
        Catch ex As Exception

        End Try

EndRoutine:
        Show()
    End Sub


    Private Sub picPBlogo_Click(sender As Object, e As EventArgs)
        Process.Start("explorer.exe", "http://www.pantareinwater.be/en")
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)
        Process.Start("explorer.exe", "http://www.pantareinwater.be/en")
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)
        Process.Start("explorer.exe", "http://www.pantareinwater.be/en")
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)
        Process.Start("explorer.exe", "http://www.pantareinwater.be/en")
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)
        Process.Start("explorer.exe", "http://www.pantareinwater.be/en")
    End Sub


End Class