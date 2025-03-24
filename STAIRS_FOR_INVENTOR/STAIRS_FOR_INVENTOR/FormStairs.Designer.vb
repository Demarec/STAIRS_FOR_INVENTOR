<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Stairs_Class
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Stairs_Class))
        Me.cboHeight = New System.Windows.Forms.ComboBox()
        Me.cboWidth = New System.Windows.Forms.ComboBox()
        Me.cboAngle = New System.Windows.Forms.ComboBox()
        Me.lblHeight = New System.Windows.Forms.Label()
        Me.lblWidth = New System.Windows.Forms.Label()
        Me.lblAngle = New System.Windows.Forms.Label()
        Me.chkLeftHandrail = New System.Windows.Forms.CheckBox()
        Me.chkRightHandrail = New System.Windows.Forms.CheckBox()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.lblPath = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.StairsTab = New System.Windows.Forms.TabPage()
        Me.btnPlaceStairsInAssem = New System.Windows.Forms.Button()
        Me.picRailingBoth = New System.Windows.Forms.PictureBox()
        Me.picRailingNone = New System.Windows.Forms.PictureBox()
        Me.picRailingLeft = New System.Windows.Forms.PictureBox()
        Me.picRailingRight = New System.Windows.Forms.PictureBox()
        Me.CurvedStairsTab = New System.Windows.Forms.TabPage()
        Me.btnPlaceCurvedStairs = New System.Windows.Forms.Button()
        Me.rbnClockWise = New System.Windows.Forms.RadioButton()
        Me.rbnCounterClockWise = New System.Windows.Forms.RadioButton()
        Me.cboCurvedStairsIntRadius = New System.Windows.Forms.ComboBox()
        Me.cboCurvedStairsHeight = New System.Windows.Forms.ComboBox()
        Me.cboCurvedStairsWidth = New System.Windows.Forms.ComboBox()
        Me.ProgressBar5 = New System.Windows.Forms.ProgressBar()
        Me.cboCurvedStairsAngle = New System.Windows.Forms.ComboBox()
        Me.lblCurvedStairsIntRadius = New System.Windows.Forms.Label()
        Me.lblCurvedStairsHeight = New System.Windows.Forms.Label()
        Me.lblCurvedStairsWidth = New System.Windows.Forms.Label()
        Me.lblCurvedStairsAngle = New System.Windows.Forms.Label()
        Me.chkLeftHandrailCurvedStairs = New System.Windows.Forms.CheckBox()
        Me.lblCurvedStairsPath = New System.Windows.Forms.Label()
        Me.chkRightHandrailCurvedStairs = New System.Windows.Forms.CheckBox()
        Me.lblCurvedStairsFile = New System.Windows.Forms.Label()
        Me.picCCW_Both = New System.Windows.Forms.PictureBox()
        Me.picCW_Both = New System.Windows.Forms.PictureBox()
        Me.picCCW_Right = New System.Windows.Forms.PictureBox()
        Me.picCW_Right = New System.Windows.Forms.PictureBox()
        Me.picCCW_Left = New System.Windows.Forms.PictureBox()
        Me.picCW_Left = New System.Windows.Forms.PictureBox()
        Me.picCCW_None = New System.Windows.Forms.PictureBox()
        Me.picCW_None = New System.Windows.Forms.PictureBox()
        Me.LadderTab = New System.Windows.Forms.TabPage()
        Me.btnPlaceLadder = New System.Windows.Forms.Button()
        Me.cboLadderHeight = New System.Windows.Forms.ComboBox()
        Me.lblLadderHeight = New System.Windows.Forms.Label()
        Me.chkCage = New System.Windows.Forms.CheckBox()
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
        Me.lblLadderPath = New System.Windows.Forms.Label()
        Me.lblLadderFile = New System.Windows.Forms.Label()
        Me.picLadderCage = New System.Windows.Forms.PictureBox()
        Me.picLadderNoCage = New System.Windows.Forms.PictureBox()
        Me.RailingTab = New System.Windows.Forms.TabPage()
        Me.btnPlaceRailing = New System.Windows.Forms.Button()
        Me.CboRailAngle = New System.Windows.Forms.ComboBox()
        Me.LblRailAngle = New System.Windows.Forms.Label()
        Me.RbnRailDia = New System.Windows.Forms.RadioButton()
        Me.RbnRailLng5 = New System.Windows.Forms.RadioButton()
        Me.RbnRailLng4 = New System.Windows.Forms.RadioButton()
        Me.RbnRailLng3 = New System.Windows.Forms.RadioButton()
        Me.RbnRailLng2 = New System.Windows.Forms.RadioButton()
        Me.RbnRailLng1 = New System.Windows.Forms.RadioButton()
        Me.CboRailDia = New System.Windows.Forms.ComboBox()
        Me.CboRailLng5 = New System.Windows.Forms.ComboBox()
        Me.CboRailLng4 = New System.Windows.Forms.ComboBox()
        Me.CboRailLng3 = New System.Windows.Forms.ComboBox()
        Me.CboRailLng2 = New System.Windows.Forms.ComboBox()
        Me.CboRailLng1 = New System.Windows.Forms.ComboBox()
        Me.LblRailDia = New System.Windows.Forms.Label()
        Me.LblRailLng5 = New System.Windows.Forms.Label()
        Me.LblRailLng4 = New System.Windows.Forms.Label()
        Me.LblRailLng3 = New System.Windows.Forms.Label()
        Me.LblRailLng2 = New System.Windows.Forms.Label()
        Me.LblRailLng1 = New System.Windows.Forms.Label()
        Me.ProgressBar3 = New System.Windows.Forms.ProgressBar()
        Me.lblRailPath = New System.Windows.Forms.Label()
        Me.lblRailFile = New System.Windows.Forms.Label()
        Me.PicRailR = New System.Windows.Forms.PictureBox()
        Me.PicRail05 = New System.Windows.Forms.PictureBox()
        Me.PicRail04 = New System.Windows.Forms.PictureBox()
        Me.PicRail03 = New System.Windows.Forms.PictureBox()
        Me.PicRail01 = New System.Windows.Forms.PictureBox()
        Me.PicRail02 = New System.Windows.Forms.PictureBox()
        Me.PlatformTab = New System.Windows.Forms.TabPage()
        Me.btnPlacePlatform = New System.Windows.Forms.Button()
        Me.ProgressBar4 = New System.Windows.Forms.ProgressBar()
        Me.lblPlatformPath = New System.Windows.Forms.Label()
        Me.lblPlatformFile = New System.Windows.Forms.Label()
        Me.CbxFoot = New System.Windows.Forms.CheckBox()
        Me.CbxLegs = New System.Windows.Forms.CheckBox()
        Me.CboPlatformDiameter = New System.Windows.Forms.ComboBox()
        Me.CboPlatformWidth = New System.Windows.Forms.ComboBox()
        Me.CboPlatformLegNumber = New System.Windows.Forms.ComboBox()
        Me.CboPlatformAngle = New System.Windows.Forms.ComboBox()
        Me.CboPlatformHeight = New System.Windows.Forms.ComboBox()
        Me.CboPlatformLenght = New System.Windows.Forms.ComboBox()
        Me.lblPlatformDiameter = New System.Windows.Forms.Label()
        Me.lblPlatformWidth = New System.Windows.Forms.Label()
        Me.lblPlatformAngle = New System.Windows.Forms.Label()
        Me.lblPlatformLegNumber = New System.Windows.Forms.Label()
        Me.lblPlatformHeight = New System.Windows.Forms.Label()
        Me.lblPlatformLenght = New System.Windows.Forms.Label()
        Me.RbnPlatformCurved = New System.Windows.Forms.RadioButton()
        Me.RbnPlatformCirc = New System.Windows.Forms.RadioButton()
        Me.RbnPlatformRect = New System.Windows.Forms.RadioButton()
        Me.PicPlatformRectLegsPlates = New System.Windows.Forms.PictureBox()
        Me.PicPlatformRectLegs = New System.Windows.Forms.PictureBox()
        Me.PicPlatformRect = New System.Windows.Forms.PictureBox()
        Me.PicPlatformCircleLegsPlates = New System.Windows.Forms.PictureBox()
        Me.PicPlatformCurved = New System.Windows.Forms.PictureBox()
        Me.PicPlatformCircleLegs = New System.Windows.Forms.PictureBox()
        Me.PicPlatformCircle = New System.Windows.Forms.PictureBox()
        Me.TabControl1.SuspendLayout()
        Me.StairsTab.SuspendLayout()
        CType(Me.picRailingBoth, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picRailingNone, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picRailingLeft, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picRailingRight, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CurvedStairsTab.SuspendLayout()
        CType(Me.picCCW_Both, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCW_Both, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCCW_Right, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCW_Right, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCCW_Left, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCW_Left, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCCW_None, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCW_None, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LadderTab.SuspendLayout()
        CType(Me.picLadderCage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLadderNoCage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RailingTab.SuspendLayout()
        CType(Me.PicRailR, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicRail05, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicRail04, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicRail03, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicRail01, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicRail02, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PlatformTab.SuspendLayout()
        CType(Me.PicPlatformRectLegsPlates, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicPlatformRectLegs, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicPlatformRect, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicPlatformCircleLegsPlates, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicPlatformCurved, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicPlatformCircleLegs, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicPlatformCircle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboHeight
        '
        Me.cboHeight.FormattingEnabled = True
        Me.cboHeight.Items.AddRange(New Object() {"4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1100", "1000", "900"})
        Me.cboHeight.Location = New System.Drawing.Point(8, 26)
        Me.cboHeight.MaxLength = 4
        Me.cboHeight.Name = "cboHeight"
        Me.cboHeight.Size = New System.Drawing.Size(58, 21)
        Me.cboHeight.TabIndex = 0
        Me.cboHeight.Text = "4000"
        '
        'cboWidth
        '
        Me.cboWidth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboWidth.FormattingEnabled = True
        Me.cboWidth.Items.AddRange(New Object() {"600", "650", "700", "750", "800", "850", "900", "950", "1000"})
        Me.cboWidth.Location = New System.Drawing.Point(8, 72)
        Me.cboWidth.Name = "cboWidth"
        Me.cboWidth.Size = New System.Drawing.Size(58, 21)
        Me.cboWidth.TabIndex = 1
        '
        'cboAngle
        '
        Me.cboAngle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAngle.FormattingEnabled = True
        Me.cboAngle.Items.AddRange(New Object() {"47°", "46°", "45°", "44°", "43°", "42°", "41°", "40°", "39°", "38°", "37°"})
        Me.cboAngle.Location = New System.Drawing.Point(8, 118)
        Me.cboAngle.Name = "cboAngle"
        Me.cboAngle.Size = New System.Drawing.Size(58, 21)
        Me.cboAngle.TabIndex = 2
        '
        'lblHeight
        '
        Me.lblHeight.AutoSize = True
        Me.lblHeight.BackColor = System.Drawing.Color.Transparent
        Me.lblHeight.Location = New System.Drawing.Point(8, 10)
        Me.lblHeight.Name = "lblHeight"
        Me.lblHeight.Size = New System.Drawing.Size(38, 13)
        Me.lblHeight.TabIndex = 3
        Me.lblHeight.Text = "Height"
        '
        'lblWidth
        '
        Me.lblWidth.AutoSize = True
        Me.lblWidth.BackColor = System.Drawing.Color.Transparent
        Me.lblWidth.Location = New System.Drawing.Point(8, 56)
        Me.lblWidth.Name = "lblWidth"
        Me.lblWidth.Size = New System.Drawing.Size(35, 13)
        Me.lblWidth.TabIndex = 4
        Me.lblWidth.Text = "Width"
        '
        'lblAngle
        '
        Me.lblAngle.AutoSize = True
        Me.lblAngle.BackColor = System.Drawing.Color.Transparent
        Me.lblAngle.Location = New System.Drawing.Point(8, 102)
        Me.lblAngle.Name = "lblAngle"
        Me.lblAngle.Size = New System.Drawing.Size(34, 13)
        Me.lblAngle.TabIndex = 5
        Me.lblAngle.Text = "Angle"
        '
        'chkLeftHandrail
        '
        Me.chkLeftHandrail.AutoSize = True
        Me.chkLeftHandrail.BackColor = System.Drawing.Color.Transparent
        Me.chkLeftHandrail.Checked = True
        Me.chkLeftHandrail.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkLeftHandrail.Location = New System.Drawing.Point(87, 98)
        Me.chkLeftHandrail.Name = "chkLeftHandrail"
        Me.chkLeftHandrail.Size = New System.Drawing.Size(79, 17)
        Me.chkLeftHandrail.TabIndex = 6
        Me.chkLeftHandrail.Text = "Railing Left"
        Me.chkLeftHandrail.UseVisualStyleBackColor = False
        '
        'chkRightHandrail
        '
        Me.chkRightHandrail.AutoSize = True
        Me.chkRightHandrail.BackColor = System.Drawing.Color.Transparent
        Me.chkRightHandrail.Checked = True
        Me.chkRightHandrail.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRightHandrail.Location = New System.Drawing.Point(87, 122)
        Me.chkRightHandrail.Name = "chkRightHandrail"
        Me.chkRightHandrail.Size = New System.Drawing.Size(86, 17)
        Me.chkRightHandrail.TabIndex = 7
        Me.chkRightHandrail.Text = "Railing Right"
        Me.chkRightHandrail.UseVisualStyleBackColor = False
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.BackColor = System.Drawing.Color.Transparent
        Me.lblFile.Location = New System.Drawing.Point(8, 195)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(49, 13)
        Me.lblFile.TabIndex = 10
        Me.lblFile.Text = "Filename"
        '
        'lblPath
        '
        Me.lblPath.AutoSize = True
        Me.lblPath.BackColor = System.Drawing.Color.Transparent
        Me.lblPath.Location = New System.Drawing.Point(8, 214)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.Size = New System.Drawing.Size(29, 13)
        Me.lblPath.TabIndex = 11
        Me.lblPath.Text = "Path"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.ProgressBar1.Location = New System.Drawing.Point(6, 291)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(230, 15)
        Me.ProgressBar1.TabIndex = 15
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.StairsTab)
        Me.TabControl1.Controls.Add(Me.CurvedStairsTab)
        Me.TabControl1.Controls.Add(Me.LadderTab)
        Me.TabControl1.Controls.Add(Me.RailingTab)
        Me.TabControl1.Controls.Add(Me.PlatformTab)
        Me.TabControl1.Location = New System.Drawing.Point(3, 9)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(440, 338)
        Me.TabControl1.TabIndex = 55
        '
        'StairsTab
        '
        Me.StairsTab.BackColor = System.Drawing.SystemColors.Control
        Me.StairsTab.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.StairsTab.Controls.Add(Me.btnPlaceStairsInAssem)
        Me.StairsTab.Controls.Add(Me.cboHeight)
        Me.StairsTab.Controls.Add(Me.cboWidth)
        Me.StairsTab.Controls.Add(Me.ProgressBar1)
        Me.StairsTab.Controls.Add(Me.cboAngle)
        Me.StairsTab.Controls.Add(Me.picRailingBoth)
        Me.StairsTab.Controls.Add(Me.lblHeight)
        Me.StairsTab.Controls.Add(Me.picRailingNone)
        Me.StairsTab.Controls.Add(Me.lblWidth)
        Me.StairsTab.Controls.Add(Me.picRailingLeft)
        Me.StairsTab.Controls.Add(Me.lblAngle)
        Me.StairsTab.Controls.Add(Me.picRailingRight)
        Me.StairsTab.Controls.Add(Me.chkLeftHandrail)
        Me.StairsTab.Controls.Add(Me.lblPath)
        Me.StairsTab.Controls.Add(Me.chkRightHandrail)
        Me.StairsTab.Controls.Add(Me.lblFile)
        Me.StairsTab.Location = New System.Drawing.Point(4, 22)
        Me.StairsTab.Name = "StairsTab"
        Me.StairsTab.Padding = New System.Windows.Forms.Padding(3)
        Me.StairsTab.Size = New System.Drawing.Size(432, 312)
        Me.StairsTab.TabIndex = 0
        Me.StairsTab.Text = "Stairs"
        '
        'btnPlaceStairsInAssem
        '
        Me.btnPlaceStairsInAssem.Location = New System.Drawing.Point(319, 270)
        Me.btnPlaceStairsInAssem.Name = "btnPlaceStairsInAssem"
        Me.btnPlaceStairsInAssem.Size = New System.Drawing.Size(110, 40)
        Me.btnPlaceStairsInAssem.TabIndex = 17
        Me.btnPlaceStairsInAssem.Text = "PLACE STAIRS"
        Me.btnPlaceStairsInAssem.UseVisualStyleBackColor = True
        '
        'picRailingBoth
        '
        Me.picRailingBoth.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.Type1_LH_RH
        Me.picRailingBoth.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picRailingBoth.Location = New System.Drawing.Point(256, 6)
        Me.picRailingBoth.Name = "picRailingBoth"
        Me.picRailingBoth.Size = New System.Drawing.Size(173, 169)
        Me.picRailingBoth.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.picRailingBoth.TabIndex = 9
        Me.picRailingBoth.TabStop = False
        '
        'picRailingNone
        '
        Me.picRailingNone.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.Type1
        Me.picRailingNone.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picRailingNone.Location = New System.Drawing.Point(256, 6)
        Me.picRailingNone.Name = "picRailingNone"
        Me.picRailingNone.Size = New System.Drawing.Size(173, 169)
        Me.picRailingNone.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picRailingNone.TabIndex = 14
        Me.picRailingNone.TabStop = False
        '
        'picRailingLeft
        '
        Me.picRailingLeft.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.Type1_LH
        Me.picRailingLeft.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picRailingLeft.Location = New System.Drawing.Point(256, 6)
        Me.picRailingLeft.Name = "picRailingLeft"
        Me.picRailingLeft.Size = New System.Drawing.Size(173, 169)
        Me.picRailingLeft.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picRailingLeft.TabIndex = 13
        Me.picRailingLeft.TabStop = False
        '
        'picRailingRight
        '
        Me.picRailingRight.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.Type1_RH
        Me.picRailingRight.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picRailingRight.Location = New System.Drawing.Point(256, 6)
        Me.picRailingRight.Name = "picRailingRight"
        Me.picRailingRight.Size = New System.Drawing.Size(173, 169)
        Me.picRailingRight.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picRailingRight.TabIndex = 12
        Me.picRailingRight.TabStop = False
        '
        'CurvedStairsTab
        '
        Me.CurvedStairsTab.BackColor = System.Drawing.SystemColors.Control
        Me.CurvedStairsTab.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.CurvedStairsTab.Controls.Add(Me.btnPlaceCurvedStairs)
        Me.CurvedStairsTab.Controls.Add(Me.rbnClockWise)
        Me.CurvedStairsTab.Controls.Add(Me.rbnCounterClockWise)
        Me.CurvedStairsTab.Controls.Add(Me.cboCurvedStairsIntRadius)
        Me.CurvedStairsTab.Controls.Add(Me.cboCurvedStairsHeight)
        Me.CurvedStairsTab.Controls.Add(Me.cboCurvedStairsWidth)
        Me.CurvedStairsTab.Controls.Add(Me.ProgressBar5)
        Me.CurvedStairsTab.Controls.Add(Me.cboCurvedStairsAngle)
        Me.CurvedStairsTab.Controls.Add(Me.lblCurvedStairsIntRadius)
        Me.CurvedStairsTab.Controls.Add(Me.lblCurvedStairsHeight)
        Me.CurvedStairsTab.Controls.Add(Me.lblCurvedStairsWidth)
        Me.CurvedStairsTab.Controls.Add(Me.lblCurvedStairsAngle)
        Me.CurvedStairsTab.Controls.Add(Me.chkLeftHandrailCurvedStairs)
        Me.CurvedStairsTab.Controls.Add(Me.lblCurvedStairsPath)
        Me.CurvedStairsTab.Controls.Add(Me.chkRightHandrailCurvedStairs)
        Me.CurvedStairsTab.Controls.Add(Me.lblCurvedStairsFile)
        Me.CurvedStairsTab.Controls.Add(Me.picCCW_Both)
        Me.CurvedStairsTab.Controls.Add(Me.picCW_Both)
        Me.CurvedStairsTab.Controls.Add(Me.picCCW_Right)
        Me.CurvedStairsTab.Controls.Add(Me.picCW_Right)
        Me.CurvedStairsTab.Controls.Add(Me.picCCW_Left)
        Me.CurvedStairsTab.Controls.Add(Me.picCW_Left)
        Me.CurvedStairsTab.Controls.Add(Me.picCCW_None)
        Me.CurvedStairsTab.Controls.Add(Me.picCW_None)
        Me.CurvedStairsTab.Location = New System.Drawing.Point(4, 22)
        Me.CurvedStairsTab.Name = "CurvedStairsTab"
        Me.CurvedStairsTab.Size = New System.Drawing.Size(432, 312)
        Me.CurvedStairsTab.TabIndex = 4
        Me.CurvedStairsTab.Text = "Curved Stairs"
        '
        'btnPlaceCurvedStairs
        '
        Me.btnPlaceCurvedStairs.Location = New System.Drawing.Point(319, 270)
        Me.btnPlaceCurvedStairs.Margin = New System.Windows.Forms.Padding(2)
        Me.btnPlaceCurvedStairs.Name = "btnPlaceCurvedStairs"
        Me.btnPlaceCurvedStairs.Size = New System.Drawing.Size(110, 40)
        Me.btnPlaceCurvedStairs.TabIndex = 34
        Me.btnPlaceCurvedStairs.Text = "PLACE STAIRS"
        Me.btnPlaceCurvedStairs.UseVisualStyleBackColor = True
        '
        'rbnClockWise
        '
        Me.rbnClockWise.AutoSize = True
        Me.rbnClockWise.BackColor = System.Drawing.Color.Transparent
        Me.rbnClockWise.Checked = True
        Me.rbnClockWise.Location = New System.Drawing.Point(87, 56)
        Me.rbnClockWise.Name = "rbnClockWise"
        Me.rbnClockWise.Size = New System.Drawing.Size(43, 17)
        Me.rbnClockWise.TabIndex = 24
        Me.rbnClockWise.TabStop = True
        Me.rbnClockWise.Text = "CW"
        Me.rbnClockWise.UseVisualStyleBackColor = False
        '
        'rbnCounterClockWise
        '
        Me.rbnCounterClockWise.AutoSize = True
        Me.rbnCounterClockWise.BackColor = System.Drawing.Color.Transparent
        Me.rbnCounterClockWise.Location = New System.Drawing.Point(87, 76)
        Me.rbnCounterClockWise.Name = "rbnCounterClockWise"
        Me.rbnCounterClockWise.Size = New System.Drawing.Size(50, 17)
        Me.rbnCounterClockWise.TabIndex = 25
        Me.rbnCounterClockWise.Text = "CCW"
        Me.rbnCounterClockWise.UseVisualStyleBackColor = False
        '
        'cboCurvedStairsIntRadius
        '
        Me.cboCurvedStairsIntRadius.FormattingEnabled = True
        Me.cboCurvedStairsIntRadius.Items.AddRange(New Object() {"4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1100", "1000", "900"})
        Me.cboCurvedStairsIntRadius.Location = New System.Drawing.Point(84, 26)
        Me.cboCurvedStairsIntRadius.MaxLength = 5
        Me.cboCurvedStairsIntRadius.Name = "cboCurvedStairsIntRadius"
        Me.cboCurvedStairsIntRadius.Size = New System.Drawing.Size(58, 21)
        Me.cboCurvedStairsIntRadius.TabIndex = 21
        Me.cboCurvedStairsIntRadius.Text = "4000"
        '
        'cboCurvedStairsHeight
        '
        Me.cboCurvedStairsHeight.FormattingEnabled = True
        Me.cboCurvedStairsHeight.Items.AddRange(New Object() {"4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1100", "1000", "900"})
        Me.cboCurvedStairsHeight.Location = New System.Drawing.Point(8, 26)
        Me.cboCurvedStairsHeight.MaxLength = 4
        Me.cboCurvedStairsHeight.Name = "cboCurvedStairsHeight"
        Me.cboCurvedStairsHeight.Size = New System.Drawing.Size(58, 21)
        Me.cboCurvedStairsHeight.TabIndex = 20
        Me.cboCurvedStairsHeight.Text = "4000"
        '
        'cboCurvedStairsWidth
        '
        Me.cboCurvedStairsWidth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurvedStairsWidth.FormattingEnabled = True
        Me.cboCurvedStairsWidth.Items.AddRange(New Object() {"600", "650", "700", "750", "800", "850", "900", "950", "1000"})
        Me.cboCurvedStairsWidth.Location = New System.Drawing.Point(8, 72)
        Me.cboCurvedStairsWidth.Name = "cboCurvedStairsWidth"
        Me.cboCurvedStairsWidth.Size = New System.Drawing.Size(58, 21)
        Me.cboCurvedStairsWidth.TabIndex = 22
        '
        'ProgressBar5
        '
        Me.ProgressBar5.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.ProgressBar5.Location = New System.Drawing.Point(3, 296)
        Me.ProgressBar5.Name = "ProgressBar5"
        Me.ProgressBar5.Size = New System.Drawing.Size(230, 15)
        Me.ProgressBar5.TabIndex = 32
        '
        'cboCurvedStairsAngle
        '
        Me.cboCurvedStairsAngle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurvedStairsAngle.FormattingEnabled = True
        Me.cboCurvedStairsAngle.Items.AddRange(New Object() {"47°", "46°", "45°", "44°", "43°", "42°", "41°", "40°", "39°", "38°", "37°"})
        Me.cboCurvedStairsAngle.Location = New System.Drawing.Point(8, 118)
        Me.cboCurvedStairsAngle.Name = "cboCurvedStairsAngle"
        Me.cboCurvedStairsAngle.Size = New System.Drawing.Size(58, 21)
        Me.cboCurvedStairsAngle.TabIndex = 23
        '
        'lblCurvedStairsIntRadius
        '
        Me.lblCurvedStairsIntRadius.AutoSize = True
        Me.lblCurvedStairsIntRadius.BackColor = System.Drawing.Color.Transparent
        Me.lblCurvedStairsIntRadius.Location = New System.Drawing.Point(84, 10)
        Me.lblCurvedStairsIntRadius.Name = "lblCurvedStairsIntRadius"
        Me.lblCurvedStairsIntRadius.Size = New System.Drawing.Size(71, 13)
        Me.lblCurvedStairsIntRadius.TabIndex = 20
        Me.lblCurvedStairsIntRadius.Text = "Inside Radius"
        '
        'lblCurvedStairsHeight
        '
        Me.lblCurvedStairsHeight.AutoSize = True
        Me.lblCurvedStairsHeight.BackColor = System.Drawing.Color.Transparent
        Me.lblCurvedStairsHeight.Location = New System.Drawing.Point(8, 10)
        Me.lblCurvedStairsHeight.Name = "lblCurvedStairsHeight"
        Me.lblCurvedStairsHeight.Size = New System.Drawing.Size(38, 13)
        Me.lblCurvedStairsHeight.TabIndex = 20
        Me.lblCurvedStairsHeight.Text = "Height"
        '
        'lblCurvedStairsWidth
        '
        Me.lblCurvedStairsWidth.AutoSize = True
        Me.lblCurvedStairsWidth.BackColor = System.Drawing.Color.Transparent
        Me.lblCurvedStairsWidth.Location = New System.Drawing.Point(8, 56)
        Me.lblCurvedStairsWidth.Name = "lblCurvedStairsWidth"
        Me.lblCurvedStairsWidth.Size = New System.Drawing.Size(35, 13)
        Me.lblCurvedStairsWidth.TabIndex = 21
        Me.lblCurvedStairsWidth.Text = "Width"
        '
        'lblCurvedStairsAngle
        '
        Me.lblCurvedStairsAngle.AutoSize = True
        Me.lblCurvedStairsAngle.BackColor = System.Drawing.Color.Transparent
        Me.lblCurvedStairsAngle.Location = New System.Drawing.Point(8, 102)
        Me.lblCurvedStairsAngle.Name = "lblCurvedStairsAngle"
        Me.lblCurvedStairsAngle.Size = New System.Drawing.Size(34, 13)
        Me.lblCurvedStairsAngle.TabIndex = 22
        Me.lblCurvedStairsAngle.Text = "Angle"
        '
        'chkLeftHandrailCurvedStairs
        '
        Me.chkLeftHandrailCurvedStairs.AutoSize = True
        Me.chkLeftHandrailCurvedStairs.BackColor = System.Drawing.Color.Transparent
        Me.chkLeftHandrailCurvedStairs.Checked = True
        Me.chkLeftHandrailCurvedStairs.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkLeftHandrailCurvedStairs.Location = New System.Drawing.Point(87, 98)
        Me.chkLeftHandrailCurvedStairs.Name = "chkLeftHandrailCurvedStairs"
        Me.chkLeftHandrailCurvedStairs.Size = New System.Drawing.Size(79, 17)
        Me.chkLeftHandrailCurvedStairs.TabIndex = 26
        Me.chkLeftHandrailCurvedStairs.Text = "Railing Left"
        Me.chkLeftHandrailCurvedStairs.UseVisualStyleBackColor = False
        '
        'lblCurvedStairsPath
        '
        Me.lblCurvedStairsPath.AutoSize = True
        Me.lblCurvedStairsPath.BackColor = System.Drawing.Color.Transparent
        Me.lblCurvedStairsPath.Location = New System.Drawing.Point(8, 214)
        Me.lblCurvedStairsPath.Name = "lblCurvedStairsPath"
        Me.lblCurvedStairsPath.Size = New System.Drawing.Size(29, 13)
        Me.lblCurvedStairsPath.TabIndex = 28
        Me.lblCurvedStairsPath.Text = "Path"
        '
        'chkRightHandrailCurvedStairs
        '
        Me.chkRightHandrailCurvedStairs.AutoSize = True
        Me.chkRightHandrailCurvedStairs.BackColor = System.Drawing.Color.Transparent
        Me.chkRightHandrailCurvedStairs.Checked = True
        Me.chkRightHandrailCurvedStairs.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRightHandrailCurvedStairs.Location = New System.Drawing.Point(87, 122)
        Me.chkRightHandrailCurvedStairs.Name = "chkRightHandrailCurvedStairs"
        Me.chkRightHandrailCurvedStairs.Size = New System.Drawing.Size(86, 17)
        Me.chkRightHandrailCurvedStairs.TabIndex = 27
        Me.chkRightHandrailCurvedStairs.Text = "Railing Right"
        Me.chkRightHandrailCurvedStairs.UseVisualStyleBackColor = False
        '
        'lblCurvedStairsFile
        '
        Me.lblCurvedStairsFile.AutoSize = True
        Me.lblCurvedStairsFile.BackColor = System.Drawing.Color.Transparent
        Me.lblCurvedStairsFile.Location = New System.Drawing.Point(8, 195)
        Me.lblCurvedStairsFile.Name = "lblCurvedStairsFile"
        Me.lblCurvedStairsFile.Size = New System.Drawing.Size(49, 13)
        Me.lblCurvedStairsFile.TabIndex = 27
        Me.lblCurvedStairsFile.Text = "Filename"
        '
        'picCCW_Both
        '
        Me.picCCW_Both.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.CCW_Both
        Me.picCCW_Both.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picCCW_Both.Location = New System.Drawing.Point(256, 6)
        Me.picCCW_Both.Name = "picCCW_Both"
        Me.picCCW_Both.Size = New System.Drawing.Size(173, 169)
        Me.picCCW_Both.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.picCCW_Both.TabIndex = 26
        Me.picCCW_Both.TabStop = False
        '
        'picCW_Both
        '
        Me.picCW_Both.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.CW_Both
        Me.picCW_Both.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picCW_Both.Location = New System.Drawing.Point(256, 6)
        Me.picCW_Both.Name = "picCW_Both"
        Me.picCW_Both.Size = New System.Drawing.Size(173, 169)
        Me.picCW_Both.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.picCW_Both.TabIndex = 26
        Me.picCW_Both.TabStop = False
        '
        'picCCW_Right
        '
        Me.picCCW_Right.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.CCW_Right
        Me.picCCW_Right.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picCCW_Right.Location = New System.Drawing.Point(256, 6)
        Me.picCCW_Right.Name = "picCCW_Right"
        Me.picCCW_Right.Size = New System.Drawing.Size(173, 169)
        Me.picCCW_Right.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picCCW_Right.TabIndex = 31
        Me.picCCW_Right.TabStop = False
        '
        'picCW_Right
        '
        Me.picCW_Right.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.CW_Right
        Me.picCW_Right.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picCW_Right.Location = New System.Drawing.Point(256, 6)
        Me.picCW_Right.Name = "picCW_Right"
        Me.picCW_Right.Size = New System.Drawing.Size(173, 169)
        Me.picCW_Right.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picCW_Right.TabIndex = 31
        Me.picCW_Right.TabStop = False
        '
        'picCCW_Left
        '
        Me.picCCW_Left.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.CCW_Left
        Me.picCCW_Left.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picCCW_Left.Location = New System.Drawing.Point(256, 6)
        Me.picCCW_Left.Name = "picCCW_Left"
        Me.picCCW_Left.Size = New System.Drawing.Size(173, 169)
        Me.picCCW_Left.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picCCW_Left.TabIndex = 30
        Me.picCCW_Left.TabStop = False
        '
        'picCW_Left
        '
        Me.picCW_Left.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.CW_Left
        Me.picCW_Left.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picCW_Left.Location = New System.Drawing.Point(256, 6)
        Me.picCW_Left.Name = "picCW_Left"
        Me.picCW_Left.Size = New System.Drawing.Size(173, 169)
        Me.picCW_Left.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picCW_Left.TabIndex = 30
        Me.picCW_Left.TabStop = False
        '
        'picCCW_None
        '
        Me.picCCW_None.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.CCW_None
        Me.picCCW_None.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picCCW_None.Location = New System.Drawing.Point(256, 6)
        Me.picCCW_None.Name = "picCCW_None"
        Me.picCCW_None.Size = New System.Drawing.Size(173, 169)
        Me.picCCW_None.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picCCW_None.TabIndex = 29
        Me.picCCW_None.TabStop = False
        '
        'picCW_None
        '
        Me.picCW_None.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.CW_None
        Me.picCW_None.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picCW_None.Location = New System.Drawing.Point(256, 6)
        Me.picCW_None.Name = "picCW_None"
        Me.picCW_None.Size = New System.Drawing.Size(173, 169)
        Me.picCW_None.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picCW_None.TabIndex = 29
        Me.picCW_None.TabStop = False
        '
        'LadderTab
        '
        Me.LadderTab.BackColor = System.Drawing.SystemColors.Control
        Me.LadderTab.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.LadderTab.Controls.Add(Me.btnPlaceLadder)
        Me.LadderTab.Controls.Add(Me.cboLadderHeight)
        Me.LadderTab.Controls.Add(Me.lblLadderHeight)
        Me.LadderTab.Controls.Add(Me.chkCage)
        Me.LadderTab.Controls.Add(Me.ProgressBar2)
        Me.LadderTab.Controls.Add(Me.lblLadderPath)
        Me.LadderTab.Controls.Add(Me.lblLadderFile)
        Me.LadderTab.Controls.Add(Me.picLadderCage)
        Me.LadderTab.Controls.Add(Me.picLadderNoCage)
        Me.LadderTab.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.LadderTab.Location = New System.Drawing.Point(4, 22)
        Me.LadderTab.Name = "LadderTab"
        Me.LadderTab.Padding = New System.Windows.Forms.Padding(3)
        Me.LadderTab.Size = New System.Drawing.Size(432, 312)
        Me.LadderTab.TabIndex = 1
        Me.LadderTab.Text = "Ladder"
        '
        'btnPlaceLadder
        '
        Me.btnPlaceLadder.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPlaceLadder.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPlaceLadder.Location = New System.Drawing.Point(319, 270)
        Me.btnPlaceLadder.Name = "btnPlaceLadder"
        Me.btnPlaceLadder.Size = New System.Drawing.Size(110, 40)
        Me.btnPlaceLadder.TabIndex = 28
        Me.btnPlaceLadder.Text = "PLACE LADDER"
        Me.btnPlaceLadder.UseVisualStyleBackColor = True
        '
        'cboLadderHeight
        '
        Me.cboLadderHeight.FormattingEnabled = True
        Me.cboLadderHeight.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500"})
        Me.cboLadderHeight.Location = New System.Drawing.Point(8, 26)
        Me.cboLadderHeight.MaxLength = 4
        Me.cboLadderHeight.Name = "cboLadderHeight"
        Me.cboLadderHeight.Size = New System.Drawing.Size(132, 21)
        Me.cboLadderHeight.TabIndex = 22
        Me.cboLadderHeight.Text = "3000"
        '
        'lblLadderHeight
        '
        Me.lblLadderHeight.AutoSize = True
        Me.lblLadderHeight.BackColor = System.Drawing.Color.Transparent
        Me.lblLadderHeight.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLadderHeight.Location = New System.Drawing.Point(8, 10)
        Me.lblLadderHeight.Name = "lblLadderHeight"
        Me.lblLadderHeight.Size = New System.Drawing.Size(38, 13)
        Me.lblLadderHeight.TabIndex = 23
        Me.lblLadderHeight.Text = "Height"
        '
        'chkCage
        '
        Me.chkCage.AutoSize = True
        Me.chkCage.BackColor = System.Drawing.Color.Transparent
        Me.chkCage.Checked = True
        Me.chkCage.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCage.Location = New System.Drawing.Point(8, 57)
        Me.chkCage.Name = "chkCage"
        Me.chkCage.Size = New System.Drawing.Size(84, 17)
        Me.chkCage.TabIndex = 24
        Me.chkCage.Text = "Safety Cage"
        Me.chkCage.UseVisualStyleBackColor = False
        '
        'ProgressBar2
        '
        Me.ProgressBar2.Location = New System.Drawing.Point(3, 296)
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(230, 15)
        Me.ProgressBar2.TabIndex = 20
        '
        'lblLadderPath
        '
        Me.lblLadderPath.AutoSize = True
        Me.lblLadderPath.BackColor = System.Drawing.Color.Transparent
        Me.lblLadderPath.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLadderPath.Location = New System.Drawing.Point(8, 214)
        Me.lblLadderPath.Name = "lblLadderPath"
        Me.lblLadderPath.Size = New System.Drawing.Size(29, 13)
        Me.lblLadderPath.TabIndex = 19
        Me.lblLadderPath.Text = "Path"
        '
        'lblLadderFile
        '
        Me.lblLadderFile.AutoSize = True
        Me.lblLadderFile.BackColor = System.Drawing.Color.Transparent
        Me.lblLadderFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLadderFile.Location = New System.Drawing.Point(8, 195)
        Me.lblLadderFile.Name = "lblLadderFile"
        Me.lblLadderFile.Size = New System.Drawing.Size(49, 13)
        Me.lblLadderFile.TabIndex = 18
        Me.lblLadderFile.Text = "Filename"
        '
        'picLadderCage
        '
        Me.picLadderCage.Image = Global.Stairs_Class.My.Resources.Resources.Ladder_C
        Me.picLadderCage.Location = New System.Drawing.Point(256, 6)
        Me.picLadderCage.Name = "picLadderCage"
        Me.picLadderCage.Size = New System.Drawing.Size(173, 169)
        Me.picLadderCage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLadderCage.TabIndex = 27
        Me.picLadderCage.TabStop = False
        '
        'picLadderNoCage
        '
        Me.picLadderNoCage.Image = Global.Stairs_Class.My.Resources.Resources.Ladder
        Me.picLadderNoCage.Location = New System.Drawing.Point(256, 6)
        Me.picLadderNoCage.Name = "picLadderNoCage"
        Me.picLadderNoCage.Size = New System.Drawing.Size(173, 169)
        Me.picLadderNoCage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLadderNoCage.TabIndex = 26
        Me.picLadderNoCage.TabStop = False
        '
        'RailingTab
        '
        Me.RailingTab.BackColor = System.Drawing.SystemColors.Control
        Me.RailingTab.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.RailingTab.Controls.Add(Me.btnPlaceRailing)
        Me.RailingTab.Controls.Add(Me.CboRailAngle)
        Me.RailingTab.Controls.Add(Me.LblRailAngle)
        Me.RailingTab.Controls.Add(Me.RbnRailDia)
        Me.RailingTab.Controls.Add(Me.RbnRailLng5)
        Me.RailingTab.Controls.Add(Me.RbnRailLng4)
        Me.RailingTab.Controls.Add(Me.RbnRailLng3)
        Me.RailingTab.Controls.Add(Me.RbnRailLng2)
        Me.RailingTab.Controls.Add(Me.RbnRailLng1)
        Me.RailingTab.Controls.Add(Me.CboRailDia)
        Me.RailingTab.Controls.Add(Me.CboRailLng5)
        Me.RailingTab.Controls.Add(Me.CboRailLng4)
        Me.RailingTab.Controls.Add(Me.CboRailLng3)
        Me.RailingTab.Controls.Add(Me.CboRailLng2)
        Me.RailingTab.Controls.Add(Me.CboRailLng1)
        Me.RailingTab.Controls.Add(Me.LblRailDia)
        Me.RailingTab.Controls.Add(Me.LblRailLng5)
        Me.RailingTab.Controls.Add(Me.LblRailLng4)
        Me.RailingTab.Controls.Add(Me.LblRailLng3)
        Me.RailingTab.Controls.Add(Me.LblRailLng2)
        Me.RailingTab.Controls.Add(Me.LblRailLng1)
        Me.RailingTab.Controls.Add(Me.ProgressBar3)
        Me.RailingTab.Controls.Add(Me.lblRailPath)
        Me.RailingTab.Controls.Add(Me.lblRailFile)
        Me.RailingTab.Controls.Add(Me.PicRailR)
        Me.RailingTab.Controls.Add(Me.PicRail05)
        Me.RailingTab.Controls.Add(Me.PicRail04)
        Me.RailingTab.Controls.Add(Me.PicRail03)
        Me.RailingTab.Controls.Add(Me.PicRail01)
        Me.RailingTab.Controls.Add(Me.PicRail02)
        Me.RailingTab.Location = New System.Drawing.Point(4, 22)
        Me.RailingTab.Name = "RailingTab"
        Me.RailingTab.Size = New System.Drawing.Size(432, 312)
        Me.RailingTab.TabIndex = 2
        Me.RailingTab.Text = "Railing"
        '
        'btnPlaceRailing
        '
        Me.btnPlaceRailing.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPlaceRailing.BackColor = System.Drawing.Color.Transparent
        Me.btnPlaceRailing.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPlaceRailing.Location = New System.Drawing.Point(319, 270)
        Me.btnPlaceRailing.Name = "btnPlaceRailing"
        Me.btnPlaceRailing.Size = New System.Drawing.Size(110, 40)
        Me.btnPlaceRailing.TabIndex = 45
        Me.btnPlaceRailing.Text = "PLACE RAILING"
        Me.btnPlaceRailing.UseVisualStyleBackColor = False
        '
        'CboRailAngle
        '
        Me.CboRailAngle.FormattingEnabled = True
        Me.CboRailAngle.Items.AddRange(New Object() {"30", "35", "40", "45", "50", "55", "60", "65", "70", "75", "80", "85", "90", "95", "100", "105", "110", "115", "120", "125", "130", "135", "140", "145", "150", "155", "160", "165", "170", "175", "180", "185", "190", "195", "200", "205", "210", "215", "220", "225", "230", "235", "240", "245", "250", "255", "260", "265", "270", "275", "280", "290", "295", "300", "305", "310", "315", "320", "325", "330", "335", "340"})
        Me.CboRailAngle.Location = New System.Drawing.Point(142, 148)
        Me.CboRailAngle.MaxLength = 3
        Me.CboRailAngle.Name = "CboRailAngle"
        Me.CboRailAngle.Size = New System.Drawing.Size(57, 21)
        Me.CboRailAngle.TabIndex = 13
        Me.CboRailAngle.Text = "330"
        '
        'LblRailAngle
        '
        Me.LblRailAngle.AutoSize = True
        Me.LblRailAngle.BackColor = System.Drawing.Color.Transparent
        Me.LblRailAngle.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblRailAngle.Location = New System.Drawing.Point(87, 151)
        Me.LblRailAngle.Name = "LblRailAngle"
        Me.LblRailAngle.Size = New System.Drawing.Size(34, 13)
        Me.LblRailAngle.TabIndex = 44
        Me.LblRailAngle.Text = "Angle"
        '
        'RbnRailDia
        '
        Me.RbnRailDia.AutoSize = True
        Me.RbnRailDia.BackColor = System.Drawing.Color.Transparent
        Me.RbnRailDia.Location = New System.Drawing.Point(8, 125)
        Me.RbnRailDia.Name = "RbnRailDia"
        Me.RbnRailDia.Size = New System.Drawing.Size(57, 17)
        Me.RbnRailDia.TabIndex = 6
        Me.RbnRailDia.Text = "Round"
        Me.RbnRailDia.UseVisualStyleBackColor = False
        '
        'RbnRailLng5
        '
        Me.RbnRailLng5.AutoSize = True
        Me.RbnRailLng5.BackColor = System.Drawing.Color.Transparent
        Me.RbnRailLng5.Location = New System.Drawing.Point(8, 102)
        Me.RbnRailLng5.Name = "RbnRailLng5"
        Me.RbnRailLng5.Size = New System.Drawing.Size(60, 17)
        Me.RbnRailLng5.TabIndex = 5
        Me.RbnRailLng5.Text = "5 Sides"
        Me.RbnRailLng5.UseVisualStyleBackColor = False
        '
        'RbnRailLng4
        '
        Me.RbnRailLng4.AutoSize = True
        Me.RbnRailLng4.BackColor = System.Drawing.Color.Transparent
        Me.RbnRailLng4.Location = New System.Drawing.Point(8, 79)
        Me.RbnRailLng4.Name = "RbnRailLng4"
        Me.RbnRailLng4.Size = New System.Drawing.Size(60, 17)
        Me.RbnRailLng4.TabIndex = 4
        Me.RbnRailLng4.Text = "4 Sides"
        Me.RbnRailLng4.UseVisualStyleBackColor = False
        '
        'RbnRailLng3
        '
        Me.RbnRailLng3.AutoSize = True
        Me.RbnRailLng3.BackColor = System.Drawing.Color.Transparent
        Me.RbnRailLng3.Location = New System.Drawing.Point(8, 56)
        Me.RbnRailLng3.Name = "RbnRailLng3"
        Me.RbnRailLng3.Size = New System.Drawing.Size(60, 17)
        Me.RbnRailLng3.TabIndex = 3
        Me.RbnRailLng3.Text = "3 Sides"
        Me.RbnRailLng3.UseVisualStyleBackColor = False
        '
        'RbnRailLng2
        '
        Me.RbnRailLng2.AutoSize = True
        Me.RbnRailLng2.BackColor = System.Drawing.Color.Transparent
        Me.RbnRailLng2.Location = New System.Drawing.Point(8, 33)
        Me.RbnRailLng2.Name = "RbnRailLng2"
        Me.RbnRailLng2.Size = New System.Drawing.Size(60, 17)
        Me.RbnRailLng2.TabIndex = 2
        Me.RbnRailLng2.Text = "2 Sides"
        Me.RbnRailLng2.UseVisualStyleBackColor = False
        '
        'RbnRailLng1
        '
        Me.RbnRailLng1.AutoSize = True
        Me.RbnRailLng1.BackColor = System.Drawing.Color.Transparent
        Me.RbnRailLng1.Checked = True
        Me.RbnRailLng1.Location = New System.Drawing.Point(8, 10)
        Me.RbnRailLng1.Name = "RbnRailLng1"
        Me.RbnRailLng1.Size = New System.Drawing.Size(55, 17)
        Me.RbnRailLng1.TabIndex = 1
        Me.RbnRailLng1.TabStop = True
        Me.RbnRailLng1.Text = "1 Side"
        Me.RbnRailLng1.UseVisualStyleBackColor = False
        '
        'CboRailDia
        '
        Me.CboRailDia.FormattingEnabled = True
        Me.CboRailDia.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1100", "1000", "900", "800", "700", "600", "500", "400", "300"})
        Me.CboRailDia.Location = New System.Drawing.Point(142, 124)
        Me.CboRailDia.MaxLength = 5
        Me.CboRailDia.Name = "CboRailDia"
        Me.CboRailDia.Size = New System.Drawing.Size(57, 21)
        Me.CboRailDia.TabIndex = 12
        Me.CboRailDia.Text = "3000"
        '
        'CboRailLng5
        '
        Me.CboRailLng5.FormattingEnabled = True
        Me.CboRailLng5.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1100", "1000", "900", "800", "700", "600", "500", "400", "300"})
        Me.CboRailLng5.Location = New System.Drawing.Point(142, 101)
        Me.CboRailLng5.MaxLength = 5
        Me.CboRailLng5.Name = "CboRailLng5"
        Me.CboRailLng5.Size = New System.Drawing.Size(57, 21)
        Me.CboRailLng5.TabIndex = 11
        Me.CboRailLng5.Text = "1500"
        '
        'CboRailLng4
        '
        Me.CboRailLng4.FormattingEnabled = True
        Me.CboRailLng4.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1100", "1000", "900", "800", "700", "600", "500", "400", "300"})
        Me.CboRailLng4.Location = New System.Drawing.Point(142, 78)
        Me.CboRailLng4.MaxLength = 5
        Me.CboRailLng4.Name = "CboRailLng4"
        Me.CboRailLng4.Size = New System.Drawing.Size(57, 21)
        Me.CboRailLng4.TabIndex = 10
        Me.CboRailLng4.Text = "2500"
        '
        'CboRailLng3
        '
        Me.CboRailLng3.FormattingEnabled = True
        Me.CboRailLng3.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1100", "1000", "900", "800", "700", "600", "500", "400", "300"})
        Me.CboRailLng3.Location = New System.Drawing.Point(142, 55)
        Me.CboRailLng3.MaxLength = 5
        Me.CboRailLng3.Name = "CboRailLng3"
        Me.CboRailLng3.Size = New System.Drawing.Size(57, 21)
        Me.CboRailLng3.TabIndex = 9
        Me.CboRailLng3.Text = "4000"
        '
        'CboRailLng2
        '
        Me.CboRailLng2.FormattingEnabled = True
        Me.CboRailLng2.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1100", "1000", "900", "800", "700", "600", "500", "400", "300"})
        Me.CboRailLng2.Location = New System.Drawing.Point(142, 32)
        Me.CboRailLng2.MaxLength = 5
        Me.CboRailLng2.Name = "CboRailLng2"
        Me.CboRailLng2.Size = New System.Drawing.Size(57, 21)
        Me.CboRailLng2.TabIndex = 8
        Me.CboRailLng2.Text = "2500"
        '
        'CboRailLng1
        '
        Me.CboRailLng1.FormattingEnabled = True
        Me.CboRailLng1.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1100", "1000", "900", "800", "700", "600", "500", "400", "300"})
        Me.CboRailLng1.Location = New System.Drawing.Point(142, 8)
        Me.CboRailLng1.MaxLength = 5
        Me.CboRailLng1.Name = "CboRailLng1"
        Me.CboRailLng1.Size = New System.Drawing.Size(57, 21)
        Me.CboRailLng1.TabIndex = 7
        Me.CboRailLng1.Text = "1500"
        '
        'LblRailDia
        '
        Me.LblRailDia.AutoSize = True
        Me.LblRailDia.BackColor = System.Drawing.Color.Transparent
        Me.LblRailDia.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblRailDia.Location = New System.Drawing.Point(87, 127)
        Me.LblRailDia.Name = "LblRailDia"
        Me.LblRailDia.Size = New System.Drawing.Size(49, 13)
        Me.LblRailDia.TabIndex = 34
        Me.LblRailDia.Text = "Diameter"
        '
        'LblRailLng5
        '
        Me.LblRailLng5.AutoSize = True
        Me.LblRailLng5.BackColor = System.Drawing.Color.Transparent
        Me.LblRailLng5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblRailLng5.Location = New System.Drawing.Point(87, 104)
        Me.LblRailLng5.Name = "LblRailLng5"
        Me.LblRailLng5.Size = New System.Drawing.Size(49, 13)
        Me.LblRailLng5.TabIndex = 34
        Me.LblRailLng5.Text = "Lenght 5"
        '
        'LblRailLng4
        '
        Me.LblRailLng4.AutoSize = True
        Me.LblRailLng4.BackColor = System.Drawing.Color.Transparent
        Me.LblRailLng4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblRailLng4.Location = New System.Drawing.Point(87, 81)
        Me.LblRailLng4.Name = "LblRailLng4"
        Me.LblRailLng4.Size = New System.Drawing.Size(49, 13)
        Me.LblRailLng4.TabIndex = 34
        Me.LblRailLng4.Text = "Lenght 4"
        '
        'LblRailLng3
        '
        Me.LblRailLng3.AutoSize = True
        Me.LblRailLng3.BackColor = System.Drawing.Color.Transparent
        Me.LblRailLng3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblRailLng3.Location = New System.Drawing.Point(87, 58)
        Me.LblRailLng3.Name = "LblRailLng3"
        Me.LblRailLng3.Size = New System.Drawing.Size(49, 13)
        Me.LblRailLng3.TabIndex = 34
        Me.LblRailLng3.Text = "Lenght 3"
        '
        'LblRailLng2
        '
        Me.LblRailLng2.AutoSize = True
        Me.LblRailLng2.BackColor = System.Drawing.Color.Transparent
        Me.LblRailLng2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblRailLng2.Location = New System.Drawing.Point(87, 35)
        Me.LblRailLng2.Name = "LblRailLng2"
        Me.LblRailLng2.Size = New System.Drawing.Size(49, 13)
        Me.LblRailLng2.TabIndex = 34
        Me.LblRailLng2.Text = "Lenght 2"
        '
        'LblRailLng1
        '
        Me.LblRailLng1.AutoSize = True
        Me.LblRailLng1.BackColor = System.Drawing.Color.Transparent
        Me.LblRailLng1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblRailLng1.Location = New System.Drawing.Point(87, 12)
        Me.LblRailLng1.Name = "LblRailLng1"
        Me.LblRailLng1.Size = New System.Drawing.Size(49, 13)
        Me.LblRailLng1.TabIndex = 34
        Me.LblRailLng1.Text = "Lenght 1"
        '
        'ProgressBar3
        '
        Me.ProgressBar3.Location = New System.Drawing.Point(3, 296)
        Me.ProgressBar3.Name = "ProgressBar3"
        Me.ProgressBar3.Size = New System.Drawing.Size(230, 15)
        Me.ProgressBar3.TabIndex = 31
        '
        'lblRailPath
        '
        Me.lblRailPath.AutoSize = True
        Me.lblRailPath.BackColor = System.Drawing.Color.Transparent
        Me.lblRailPath.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRailPath.Location = New System.Drawing.Point(8, 214)
        Me.lblRailPath.Name = "lblRailPath"
        Me.lblRailPath.Size = New System.Drawing.Size(29, 13)
        Me.lblRailPath.TabIndex = 30
        Me.lblRailPath.Text = "Path"
        '
        'lblRailFile
        '
        Me.lblRailFile.AutoSize = True
        Me.lblRailFile.BackColor = System.Drawing.Color.Transparent
        Me.lblRailFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRailFile.Location = New System.Drawing.Point(8, 195)
        Me.lblRailFile.Name = "lblRailFile"
        Me.lblRailFile.Size = New System.Drawing.Size(49, 13)
        Me.lblRailFile.TabIndex = 29
        Me.lblRailFile.Text = "Filename"
        '
        'PicRailR
        '
        Me.PicRailR.Image = Global.Stairs_Class.My.Resources.Resources.RailR
        Me.PicRailR.Location = New System.Drawing.Point(256, 6)
        Me.PicRailR.Name = "PicRailR"
        Me.PicRailR.Size = New System.Drawing.Size(173, 169)
        Me.PicRailR.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicRailR.TabIndex = 41
        Me.PicRailR.TabStop = False
        Me.PicRailR.Visible = False
        '
        'PicRail05
        '
        Me.PicRail05.Image = Global.Stairs_Class.My.Resources.Resources.Rail5
        Me.PicRail05.Location = New System.Drawing.Point(256, 6)
        Me.PicRail05.Name = "PicRail05"
        Me.PicRail05.Size = New System.Drawing.Size(173, 169)
        Me.PicRail05.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicRail05.TabIndex = 40
        Me.PicRail05.TabStop = False
        Me.PicRail05.Visible = False
        '
        'PicRail04
        '
        Me.PicRail04.Image = Global.Stairs_Class.My.Resources.Resources.Rail4
        Me.PicRail04.Location = New System.Drawing.Point(256, 6)
        Me.PicRail04.Name = "PicRail04"
        Me.PicRail04.Size = New System.Drawing.Size(173, 169)
        Me.PicRail04.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicRail04.TabIndex = 39
        Me.PicRail04.TabStop = False
        Me.PicRail04.Visible = False
        '
        'PicRail03
        '
        Me.PicRail03.Image = Global.Stairs_Class.My.Resources.Resources.Rail3
        Me.PicRail03.Location = New System.Drawing.Point(256, 6)
        Me.PicRail03.Name = "PicRail03"
        Me.PicRail03.Size = New System.Drawing.Size(173, 169)
        Me.PicRail03.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicRail03.TabIndex = 38
        Me.PicRail03.TabStop = False
        Me.PicRail03.Visible = False
        '
        'PicRail01
        '
        Me.PicRail01.Image = Global.Stairs_Class.My.Resources.Resources.Rail1
        Me.PicRail01.Location = New System.Drawing.Point(256, 6)
        Me.PicRail01.Name = "PicRail01"
        Me.PicRail01.Size = New System.Drawing.Size(173, 169)
        Me.PicRail01.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicRail01.TabIndex = 37
        Me.PicRail01.TabStop = False
        '
        'PicRail02
        '
        Me.PicRail02.Image = Global.Stairs_Class.My.Resources.Resources.Rail2
        Me.PicRail02.Location = New System.Drawing.Point(256, 6)
        Me.PicRail02.Name = "PicRail02"
        Me.PicRail02.Size = New System.Drawing.Size(173, 169)
        Me.PicRail02.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicRail02.TabIndex = 36
        Me.PicRail02.TabStop = False
        Me.PicRail02.Visible = False
        '
        'PlatformTab
        '
        Me.PlatformTab.BackColor = System.Drawing.SystemColors.Control
        Me.PlatformTab.Controls.Add(Me.btnPlacePlatform)
        Me.PlatformTab.Controls.Add(Me.ProgressBar4)
        Me.PlatformTab.Controls.Add(Me.lblPlatformPath)
        Me.PlatformTab.Controls.Add(Me.lblPlatformFile)
        Me.PlatformTab.Controls.Add(Me.CbxFoot)
        Me.PlatformTab.Controls.Add(Me.CbxLegs)
        Me.PlatformTab.Controls.Add(Me.CboPlatformDiameter)
        Me.PlatformTab.Controls.Add(Me.CboPlatformWidth)
        Me.PlatformTab.Controls.Add(Me.CboPlatformLegNumber)
        Me.PlatformTab.Controls.Add(Me.CboPlatformAngle)
        Me.PlatformTab.Controls.Add(Me.CboPlatformHeight)
        Me.PlatformTab.Controls.Add(Me.CboPlatformLenght)
        Me.PlatformTab.Controls.Add(Me.lblPlatformDiameter)
        Me.PlatformTab.Controls.Add(Me.lblPlatformWidth)
        Me.PlatformTab.Controls.Add(Me.lblPlatformAngle)
        Me.PlatformTab.Controls.Add(Me.lblPlatformLegNumber)
        Me.PlatformTab.Controls.Add(Me.lblPlatformHeight)
        Me.PlatformTab.Controls.Add(Me.lblPlatformLenght)
        Me.PlatformTab.Controls.Add(Me.RbnPlatformCurved)
        Me.PlatformTab.Controls.Add(Me.RbnPlatformCirc)
        Me.PlatformTab.Controls.Add(Me.RbnPlatformRect)
        Me.PlatformTab.Controls.Add(Me.PicPlatformRectLegsPlates)
        Me.PlatformTab.Controls.Add(Me.PicPlatformRectLegs)
        Me.PlatformTab.Controls.Add(Me.PicPlatformRect)
        Me.PlatformTab.Controls.Add(Me.PicPlatformCircleLegsPlates)
        Me.PlatformTab.Controls.Add(Me.PicPlatformCurved)
        Me.PlatformTab.Controls.Add(Me.PicPlatformCircleLegs)
        Me.PlatformTab.Controls.Add(Me.PicPlatformCircle)
        Me.PlatformTab.Location = New System.Drawing.Point(4, 22)
        Me.PlatformTab.Name = "PlatformTab"
        Me.PlatformTab.Size = New System.Drawing.Size(432, 312)
        Me.PlatformTab.TabIndex = 3
        Me.PlatformTab.Text = "Platform"
        '
        'btnPlacePlatform
        '
        Me.btnPlacePlatform.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPlacePlatform.Location = New System.Drawing.Point(319, 270)
        Me.btnPlacePlatform.Name = "btnPlacePlatform"
        Me.btnPlacePlatform.Size = New System.Drawing.Size(110, 40)
        Me.btnPlacePlatform.TabIndex = 66
        Me.btnPlacePlatform.Text = "PLACE PLATFORM"
        Me.btnPlacePlatform.UseVisualStyleBackColor = True
        '
        'ProgressBar4
        '
        Me.ProgressBar4.Location = New System.Drawing.Point(3, 296)
        Me.ProgressBar4.Name = "ProgressBar4"
        Me.ProgressBar4.Size = New System.Drawing.Size(230, 15)
        Me.ProgressBar4.TabIndex = 34
        '
        'lblPlatformPath
        '
        Me.lblPlatformPath.AutoSize = True
        Me.lblPlatformPath.BackColor = System.Drawing.Color.Transparent
        Me.lblPlatformPath.Location = New System.Drawing.Point(8, 214)
        Me.lblPlatformPath.Name = "lblPlatformPath"
        Me.lblPlatformPath.Size = New System.Drawing.Size(29, 13)
        Me.lblPlatformPath.TabIndex = 33
        Me.lblPlatformPath.Text = "Path"
        '
        'lblPlatformFile
        '
        Me.lblPlatformFile.AutoSize = True
        Me.lblPlatformFile.BackColor = System.Drawing.Color.Transparent
        Me.lblPlatformFile.Location = New System.Drawing.Point(8, 195)
        Me.lblPlatformFile.Name = "lblPlatformFile"
        Me.lblPlatformFile.Size = New System.Drawing.Size(49, 13)
        Me.lblPlatformFile.TabIndex = 32
        Me.lblPlatformFile.Text = "Filename"
        '
        'CbxFoot
        '
        Me.CbxFoot.AutoSize = True
        Me.CbxFoot.BackColor = System.Drawing.Color.Transparent
        Me.CbxFoot.Checked = True
        Me.CbxFoot.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CbxFoot.Location = New System.Drawing.Point(130, 50)
        Me.CbxFoot.Name = "CbxFoot"
        Me.CbxFoot.Size = New System.Drawing.Size(79, 17)
        Me.CbxFoot.TabIndex = 62
        Me.CbxFoot.Text = "Foot Plates"
        Me.CbxFoot.UseVisualStyleBackColor = False
        '
        'CbxLegs
        '
        Me.CbxLegs.AutoSize = True
        Me.CbxLegs.BackColor = System.Drawing.Color.Transparent
        Me.CbxLegs.Checked = True
        Me.CbxLegs.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CbxLegs.Location = New System.Drawing.Point(130, 30)
        Me.CbxLegs.Name = "CbxLegs"
        Me.CbxLegs.Size = New System.Drawing.Size(49, 17)
        Me.CbxLegs.TabIndex = 61
        Me.CbxLegs.Text = "Legs"
        Me.CbxLegs.UseVisualStyleBackColor = False
        '
        'CboPlatformDiameter
        '
        Me.CboPlatformDiameter.FormattingEnabled = True
        Me.CboPlatformDiameter.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1150", "1100", "1050", "1000", "950", "900", "850", "800"})
        Me.CboPlatformDiameter.Location = New System.Drawing.Point(8, 92)
        Me.CboPlatformDiameter.MaxLength = 5
        Me.CboPlatformDiameter.Name = "CboPlatformDiameter"
        Me.CboPlatformDiameter.Size = New System.Drawing.Size(58, 21)
        Me.CboPlatformDiameter.TabIndex = 59
        Me.CboPlatformDiameter.Text = "3000"
        '
        'CboPlatformWidth
        '
        Me.CboPlatformWidth.FormattingEnabled = True
        Me.CboPlatformWidth.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1200", "1150", "1100", "1050", "1000", "950", "900", "850", "800"})
        Me.CboPlatformWidth.Location = New System.Drawing.Point(8, 137)
        Me.CboPlatformWidth.MaxLength = 5
        Me.CboPlatformWidth.Name = "CboPlatformWidth"
        Me.CboPlatformWidth.Size = New System.Drawing.Size(58, 21)
        Me.CboPlatformWidth.TabIndex = 60
        Me.CboPlatformWidth.Text = "1000"
        '
        'CboPlatformLegNumber
        '
        Me.CboPlatformLegNumber.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPlatformLegNumber.FormattingEnabled = True
        Me.CboPlatformLegNumber.Items.AddRange(New Object() {"4", "6", "8"})
        Me.CboPlatformLegNumber.Location = New System.Drawing.Point(130, 137)
        Me.CboPlatformLegNumber.MaxLength = 1
        Me.CboPlatformLegNumber.Name = "CboPlatformLegNumber"
        Me.CboPlatformLegNumber.Size = New System.Drawing.Size(58, 21)
        Me.CboPlatformLegNumber.TabIndex = 64
        '
        'CboPlatformAngle
        '
        Me.CboPlatformAngle.FormattingEnabled = True
        Me.CboPlatformAngle.Items.AddRange(New Object() {"30", "35", "40", "45", "50", "55", "60", "65", "70", "75", "80", "85", "90", "95", "100", "105", "110", "115", "120", "125", "130", "135", "140", "145", "150", "155", "160", "165", "170", "175", "180", "185", "190", "195", "200", "205", "210", "215", "220", "225", "230", "235", "240", "245", "250", "255", "260", "265", "270", "275", "280", "290", "295", "300", "305", "310", "315", "320", "325", "330", "335", "340"})
        Me.CboPlatformAngle.Location = New System.Drawing.Point(130, 92)
        Me.CboPlatformAngle.MaxLength = 4
        Me.CboPlatformAngle.Name = "CboPlatformAngle"
        Me.CboPlatformAngle.Size = New System.Drawing.Size(58, 21)
        Me.CboPlatformAngle.TabIndex = 63
        Me.CboPlatformAngle.Text = "60"
        '
        'CboPlatformHeight
        '
        Me.CboPlatformHeight.FormattingEnabled = True
        Me.CboPlatformHeight.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1200", "1150", "1100", "1050", "1000", "950", "900", "850", "800"})
        Me.CboPlatformHeight.Location = New System.Drawing.Point(130, 92)
        Me.CboPlatformHeight.MaxLength = 4
        Me.CboPlatformHeight.Name = "CboPlatformHeight"
        Me.CboPlatformHeight.Size = New System.Drawing.Size(58, 21)
        Me.CboPlatformHeight.TabIndex = 63
        Me.CboPlatformHeight.Text = "3000"
        '
        'CboPlatformLenght
        '
        Me.CboPlatformLenght.FormattingEnabled = True
        Me.CboPlatformLenght.Items.AddRange(New Object() {"6000", "5900", "5800", "5700", "5600", "5500", "5400", "5300", "5200", "5100", "5000", "4900", "4800", "4700", "4600", "4500", "4400", "4300", "4200", "4100", "4000", "3900", "3800", "3700", "3600", "3500", "3400", "3300", "3200", "3100", "3000", "2900", "2800", "2700", "2600", "2500", "2400", "2300", "2200", "2100", "2000", "1900", "1800", "1700", "1600", "1500", "1400", "1300", "1200", "1150", "1100", "1050", "1000", "950", "900", "850", "800"})
        Me.CboPlatformLenght.Location = New System.Drawing.Point(8, 92)
        Me.CboPlatformLenght.MaxLength = 5
        Me.CboPlatformLenght.Name = "CboPlatformLenght"
        Me.CboPlatformLenght.Size = New System.Drawing.Size(58, 21)
        Me.CboPlatformLenght.TabIndex = 59
        Me.CboPlatformLenght.Text = "3000"
        '
        'lblPlatformDiameter
        '
        Me.lblPlatformDiameter.AutoSize = True
        Me.lblPlatformDiameter.BackColor = System.Drawing.Color.Transparent
        Me.lblPlatformDiameter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlatformDiameter.Location = New System.Drawing.Point(8, 75)
        Me.lblPlatformDiameter.Name = "lblPlatformDiameter"
        Me.lblPlatformDiameter.Size = New System.Drawing.Size(49, 13)
        Me.lblPlatformDiameter.TabIndex = 28
        Me.lblPlatformDiameter.Text = "Diameter"
        '
        'lblPlatformWidth
        '
        Me.lblPlatformWidth.AutoSize = True
        Me.lblPlatformWidth.BackColor = System.Drawing.Color.Transparent
        Me.lblPlatformWidth.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlatformWidth.Location = New System.Drawing.Point(8, 120)
        Me.lblPlatformWidth.Name = "lblPlatformWidth"
        Me.lblPlatformWidth.Size = New System.Drawing.Size(35, 13)
        Me.lblPlatformWidth.TabIndex = 28
        Me.lblPlatformWidth.Text = "Width"
        '
        'lblPlatformAngle
        '
        Me.lblPlatformAngle.AutoSize = True
        Me.lblPlatformAngle.BackColor = System.Drawing.Color.Transparent
        Me.lblPlatformAngle.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlatformAngle.Location = New System.Drawing.Point(130, 75)
        Me.lblPlatformAngle.Name = "lblPlatformAngle"
        Me.lblPlatformAngle.Size = New System.Drawing.Size(34, 13)
        Me.lblPlatformAngle.TabIndex = 28
        Me.lblPlatformAngle.Text = "Angle"
        '
        'lblPlatformLegNumber
        '
        Me.lblPlatformLegNumber.AutoSize = True
        Me.lblPlatformLegNumber.BackColor = System.Drawing.Color.Transparent
        Me.lblPlatformLegNumber.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlatformLegNumber.Location = New System.Drawing.Point(130, 120)
        Me.lblPlatformLegNumber.Name = "lblPlatformLegNumber"
        Me.lblPlatformLegNumber.Size = New System.Drawing.Size(82, 13)
        Me.lblPlatformLegNumber.TabIndex = 28
        Me.lblPlatformLegNumber.Text = "Number of Legs"
        '
        'lblPlatformHeight
        '
        Me.lblPlatformHeight.AutoSize = True
        Me.lblPlatformHeight.BackColor = System.Drawing.Color.Transparent
        Me.lblPlatformHeight.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlatformHeight.Location = New System.Drawing.Point(130, 75)
        Me.lblPlatformHeight.Name = "lblPlatformHeight"
        Me.lblPlatformHeight.Size = New System.Drawing.Size(38, 13)
        Me.lblPlatformHeight.TabIndex = 28
        Me.lblPlatformHeight.Text = "Height"
        '
        'lblPlatformLenght
        '
        Me.lblPlatformLenght.AutoSize = True
        Me.lblPlatformLenght.BackColor = System.Drawing.Color.Transparent
        Me.lblPlatformLenght.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlatformLenght.Location = New System.Drawing.Point(8, 75)
        Me.lblPlatformLenght.Name = "lblPlatformLenght"
        Me.lblPlatformLenght.Size = New System.Drawing.Size(40, 13)
        Me.lblPlatformLenght.TabIndex = 28
        Me.lblPlatformLenght.Text = "Lenght"
        '
        'RbnPlatformCurved
        '
        Me.RbnPlatformCurved.AutoSize = True
        Me.RbnPlatformCurved.BackColor = System.Drawing.Color.Transparent
        Me.RbnPlatformCurved.Location = New System.Drawing.Point(8, 50)
        Me.RbnPlatformCurved.Name = "RbnPlatformCurved"
        Me.RbnPlatformCurved.Size = New System.Drawing.Size(100, 17)
        Me.RbnPlatformCurved.TabIndex = 58
        Me.RbnPlatformCurved.Text = "Curved Platform"
        Me.RbnPlatformCurved.UseVisualStyleBackColor = False
        '
        'RbnPlatformCirc
        '
        Me.RbnPlatformCirc.AutoSize = True
        Me.RbnPlatformCirc.BackColor = System.Drawing.Color.Transparent
        Me.RbnPlatformCirc.Location = New System.Drawing.Point(8, 30)
        Me.RbnPlatformCirc.Name = "RbnPlatformCirc"
        Me.RbnPlatformCirc.Size = New System.Drawing.Size(98, 17)
        Me.RbnPlatformCirc.TabIndex = 57
        Me.RbnPlatformCirc.Tag = ""
        Me.RbnPlatformCirc.Text = "Round Platform"
        Me.RbnPlatformCirc.UseVisualStyleBackColor = False
        '
        'RbnPlatformRect
        '
        Me.RbnPlatformRect.AutoSize = True
        Me.RbnPlatformRect.BackColor = System.Drawing.Color.Transparent
        Me.RbnPlatformRect.Checked = True
        Me.RbnPlatformRect.Location = New System.Drawing.Point(8, 10)
        Me.RbnPlatformRect.Name = "RbnPlatformRect"
        Me.RbnPlatformRect.Size = New System.Drawing.Size(115, 17)
        Me.RbnPlatformRect.TabIndex = 56
        Me.RbnPlatformRect.TabStop = True
        Me.RbnPlatformRect.Text = "Rectangle Platform"
        Me.RbnPlatformRect.UseVisualStyleBackColor = False
        '
        'PicPlatformRectLegsPlates
        '
        Me.PicPlatformRectLegsPlates.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.PlatformRectLegs
        Me.PicPlatformRectLegsPlates.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PicPlatformRectLegsPlates.Location = New System.Drawing.Point(256, 6)
        Me.PicPlatformRectLegsPlates.Name = "PicPlatformRectLegsPlates"
        Me.PicPlatformRectLegsPlates.Size = New System.Drawing.Size(173, 169)
        Me.PicPlatformRectLegsPlates.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PicPlatformRectLegsPlates.TabIndex = 15
        Me.PicPlatformRectLegsPlates.TabStop = False
        '
        'PicPlatformRectLegs
        '
        Me.PicPlatformRectLegs.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.PlatformRectLegsNoFootPlates
        Me.PicPlatformRectLegs.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PicPlatformRectLegs.Location = New System.Drawing.Point(256, 6)
        Me.PicPlatformRectLegs.Name = "PicPlatformRectLegs"
        Me.PicPlatformRectLegs.Size = New System.Drawing.Size(173, 169)
        Me.PicPlatformRectLegs.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicPlatformRectLegs.TabIndex = 18
        Me.PicPlatformRectLegs.TabStop = False
        '
        'PicPlatformRect
        '
        Me.PicPlatformRect.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.PlatformRectNoLegs
        Me.PicPlatformRect.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PicPlatformRect.Location = New System.Drawing.Point(256, 6)
        Me.PicPlatformRect.Name = "PicPlatformRect"
        Me.PicPlatformRect.Size = New System.Drawing.Size(173, 169)
        Me.PicPlatformRect.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicPlatformRect.TabIndex = 17
        Me.PicPlatformRect.TabStop = False
        '
        'PicPlatformCircleLegsPlates
        '
        Me.PicPlatformCircleLegsPlates.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.PlatformCircLegs
        Me.PicPlatformCircleLegsPlates.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PicPlatformCircleLegsPlates.Location = New System.Drawing.Point(256, 6)
        Me.PicPlatformCircleLegsPlates.Name = "PicPlatformCircleLegsPlates"
        Me.PicPlatformCircleLegsPlates.Size = New System.Drawing.Size(173, 169)
        Me.PicPlatformCircleLegsPlates.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicPlatformCircleLegsPlates.TabIndex = 16
        Me.PicPlatformCircleLegsPlates.TabStop = False
        '
        'PicPlatformCurved
        '
        Me.PicPlatformCurved.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.CurvedPlatform
        Me.PicPlatformCurved.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PicPlatformCurved.Location = New System.Drawing.Point(256, 6)
        Me.PicPlatformCurved.Name = "PicPlatformCurved"
        Me.PicPlatformCurved.Size = New System.Drawing.Size(173, 169)
        Me.PicPlatformCurved.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicPlatformCurved.TabIndex = 21
        Me.PicPlatformCurved.TabStop = False
        '
        'PicPlatformCircleLegs
        '
        Me.PicPlatformCircleLegs.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.PlatformCircLegsNoFootplates
        Me.PicPlatformCircleLegs.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PicPlatformCircleLegs.Location = New System.Drawing.Point(256, 6)
        Me.PicPlatformCircleLegs.Name = "PicPlatformCircleLegs"
        Me.PicPlatformCircleLegs.Size = New System.Drawing.Size(173, 169)
        Me.PicPlatformCircleLegs.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicPlatformCircleLegs.TabIndex = 20
        Me.PicPlatformCircleLegs.TabStop = False
        '
        'PicPlatformCircle
        '
        Me.PicPlatformCircle.BackgroundImage = Global.Stairs_Class.My.Resources.Resources.PlatformCircNoLegs
        Me.PicPlatformCircle.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PicPlatformCircle.Location = New System.Drawing.Point(256, 6)
        Me.PicPlatformCircle.Name = "PicPlatformCircle"
        Me.PicPlatformCircle.Size = New System.Drawing.Size(173, 169)
        Me.PicPlatformCircle.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicPlatformCircle.TabIndex = 19
        Me.PicPlatformCircle.TabStop = False
        '
        'Stairs_Class
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(456, 365)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(472, 404)
        Me.MinimumSize = New System.Drawing.Size(472, 404)
        Me.Name = "Stairs_Class"
        Me.Text = "Stairs and platforms"
        Me.TopMost = True
        Me.TabControl1.ResumeLayout(False)
        Me.StairsTab.ResumeLayout(False)
        Me.StairsTab.PerformLayout()
        CType(Me.picRailingBoth, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picRailingNone, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picRailingLeft, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picRailingRight, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CurvedStairsTab.ResumeLayout(False)
        Me.CurvedStairsTab.PerformLayout()
        CType(Me.picCCW_Both, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCW_Both, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCCW_Right, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCW_Right, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCCW_Left, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCW_Left, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCCW_None, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCW_None, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LadderTab.ResumeLayout(False)
        Me.LadderTab.PerformLayout()
        CType(Me.picLadderCage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLadderNoCage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RailingTab.ResumeLayout(False)
        Me.RailingTab.PerformLayout()
        CType(Me.PicRailR, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicRail05, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicRail04, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicRail03, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicRail01, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicRail02, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PlatformTab.ResumeLayout(False)
        Me.PlatformTab.PerformLayout()
        CType(Me.PicPlatformRectLegsPlates, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicPlatformRectLegs, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicPlatformRect, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicPlatformCircleLegsPlates, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicPlatformCurved, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicPlatformCircleLegs, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicPlatformCircle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub



    Friend WithEvents cboHeight As ComboBox
    Friend WithEvents cboWidth As ComboBox
    Friend WithEvents cboAngle As ComboBox
    Friend WithEvents lblHeight As Label
    Friend WithEvents lblWidth As Label
    Friend WithEvents lblAngle As Label
    Friend WithEvents chkLeftHandrail As CheckBox
    Friend WithEvents chkRightHandrail As CheckBox
    Friend WithEvents picRailingBoth As PictureBox
    Friend WithEvents lblFile As Label
    Friend WithEvents lblPath As Label
    Friend WithEvents picRailingRight As PictureBox
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents StairsTab As TabPage
    Friend WithEvents LadderTab As TabPage
    Friend WithEvents RailingTab As TabPage
    Friend WithEvents PlatformTab As TabPage
    Friend WithEvents ProgressBar2 As ProgressBar
    Friend WithEvents lblLadderPath As Label
    Friend WithEvents lblLadderFile As Label
    Friend WithEvents cboLadderHeight As ComboBox
    Friend WithEvents lblLadderHeight As Label
    Friend WithEvents chkCage As CheckBox
    Friend WithEvents picRailingNone As PictureBox
    Friend WithEvents picRailingLeft As PictureBox
    Friend WithEvents picLadderCage As PictureBox
    Friend WithEvents picLadderNoCage As PictureBox
    Friend WithEvents PicRail01 As PictureBox
    Friend WithEvents PicRail02 As PictureBox
    Friend WithEvents CboRailLng1 As ComboBox
    Friend WithEvents LblRailLng1 As Label
    Friend WithEvents ProgressBar3 As ProgressBar
    Friend WithEvents lblRailPath As Label
    Friend WithEvents lblRailFile As Label
    Friend WithEvents PicRailR As PictureBox
    Friend WithEvents PicRail05 As PictureBox
    Friend WithEvents PicRail04 As PictureBox
    Friend WithEvents PicRail03 As PictureBox
    Friend WithEvents RbnRailDia As RadioButton
    Friend WithEvents RbnRailLng5 As RadioButton
    Friend WithEvents RbnRailLng4 As RadioButton
    Friend WithEvents RbnRailLng3 As RadioButton
    Friend WithEvents RbnRailLng2 As RadioButton
    Friend WithEvents RbnRailLng1 As RadioButton
    Friend WithEvents CboRailDia As ComboBox
    Friend WithEvents CboRailLng5 As ComboBox
    Friend WithEvents CboRailLng4 As ComboBox
    Friend WithEvents CboRailLng3 As ComboBox
    Friend WithEvents CboRailLng2 As ComboBox
    Friend WithEvents LblRailDia As Label
    Friend WithEvents LblRailLng5 As Label
    Friend WithEvents LblRailLng4 As Label
    Friend WithEvents LblRailLng3 As Label
    Friend WithEvents LblRailLng2 As Label
    Friend WithEvents CboRailAngle As ComboBox
    Friend WithEvents LblRailAngle As Label
    Friend WithEvents PicPlatformRectLegsPlates As PictureBox
    Friend WithEvents PicPlatformRectLegs As PictureBox
    Friend WithEvents PicPlatformRect As PictureBox
    Friend WithEvents PicPlatformCircleLegsPlates As PictureBox
    Friend WithEvents PicPlatformCircleLegs As PictureBox
    Friend WithEvents PicPlatformCircle As PictureBox
    Friend WithEvents PicPlatformCurved As PictureBox
    Friend WithEvents RbnPlatformCurved As RadioButton
    Friend WithEvents RbnPlatformCirc As RadioButton
    Friend WithEvents RbnPlatformRect As RadioButton
    Friend WithEvents CbxFoot As CheckBox
    Friend WithEvents CbxLegs As CheckBox
    Friend WithEvents CboPlatformDiameter As ComboBox
    Friend WithEvents CboPlatformWidth As ComboBox
    Friend WithEvents CboPlatformLegNumber As ComboBox
    Friend WithEvents CboPlatformHeight As ComboBox
    Friend WithEvents CboPlatformLenght As ComboBox
    Friend WithEvents lblPlatformDiameter As Label
    Friend WithEvents lblPlatformWidth As Label
    Friend WithEvents lblPlatformLegNumber As Label
    Friend WithEvents lblPlatformHeight As Label
    Friend WithEvents lblPlatformLenght As Label
    Friend WithEvents ProgressBar4 As ProgressBar
    Friend WithEvents lblPlatformPath As Label
    Friend WithEvents lblPlatformFile As Label
    Friend WithEvents CboPlatformAngle As ComboBox
    Friend WithEvents lblPlatformAngle As Label
    Friend WithEvents CurvedStairsTab As TabPage
    Friend WithEvents cboCurvedStairsHeight As ComboBox
    Friend WithEvents cboCurvedStairsWidth As ComboBox
    Friend WithEvents ProgressBar5 As ProgressBar
    Friend WithEvents cboCurvedStairsAngle As ComboBox
    Friend WithEvents picCW_Both As PictureBox
    Friend WithEvents lblCurvedStairsHeight As Label
    Friend WithEvents picCW_Right As PictureBox
    Friend WithEvents lblCurvedStairsWidth As Label
    Friend WithEvents picCW_Left As PictureBox
    Friend WithEvents lblCurvedStairsAngle As Label
    Friend WithEvents picCW_None As PictureBox
    Friend WithEvents chkLeftHandrailCurvedStairs As CheckBox
    Friend WithEvents lblCurvedStairsPath As Label
    Friend WithEvents chkRightHandrailCurvedStairs As CheckBox
    Friend WithEvents lblCurvedStairsFile As Label
    Friend WithEvents rbnClockWise As RadioButton
    Friend WithEvents rbnCounterClockWise As RadioButton
    Friend WithEvents cboCurvedStairsIntRadius As ComboBox
    Friend WithEvents lblCurvedStairsIntRadius As Label
    Friend WithEvents picCCW_Both As PictureBox
    Friend WithEvents picCCW_Right As PictureBox
    Friend WithEvents picCCW_Left As PictureBox
    Friend WithEvents picCCW_None As PictureBox
    Friend WithEvents btnPlaceStairsInAssem As Button
    Friend WithEvents btnPlaceCurvedStairs As Button
    Friend WithEvents btnPlaceLadder As Button
    Friend WithEvents btnPlaceRailing As Button
    Friend WithEvents btnPlacePlatform As Button
End Class
