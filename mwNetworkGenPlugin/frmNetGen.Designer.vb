<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNetGen
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
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
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container
    Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
    Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
    Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
    Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNetGen))
    Me.TabMain = New System.Windows.Forms.TabControl
    Me.TabPageNetworkGeneration = New System.Windows.Forms.TabPage
    Me.TabControl1 = New System.Windows.Forms.TabControl
    Me.TabPage6 = New System.Windows.Forms.TabPage
    Me.grdNode = New System.Windows.Forms.DataGridView
    Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.Column8 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.nodeTYP = New System.Windows.Forms.DataGridViewComboBoxColumn
    Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.Column23 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.TabPage7 = New System.Windows.Forms.TabPage
    Me.grdLink = New System.Windows.Forms.DataGridView
    Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.cmdInitialNET = New System.Windows.Forms.Button
    Me.chkUseExistingNode = New System.Windows.Forms.CheckBox
    Me.cmdAutoGen = New System.Windows.Forms.Button
    Me.cmdSelectNode = New System.Windows.Forms.Button
    Me.cmdCreateNode = New System.Windows.Forms.Button
    Me.TabPageNetworkEditor = New System.Windows.Forms.TabPage
    Me.lblErrorLink = New System.Windows.Forms.Label
    Me.GroupBox2 = New System.Windows.Forms.GroupBox
    Me.cmdSaveUnLinkFile = New System.Windows.Forms.Button
    Me.cmdOpenUnLinkFile = New System.Windows.Forms.Button
    Me.chkDrawCheckingPath = New System.Windows.Forms.CheckBox
    Me.cmdNetworkCheck = New System.Windows.Forms.Button
    Me.grdUnconnectedNode = New System.Windows.Forms.DataGridView
    Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn
    Me.cmdDrawIncompletedNetwork = New System.Windows.Forms.Button
    Me.cmdClearDrawing1 = New System.Windows.Forms.Button
    Me.RadioButton2 = New System.Windows.Forms.RadioButton
    Me.RadioButton1 = New System.Windows.Forms.RadioButton
    Me.ListBox1 = New System.Windows.Forms.ListBox
    Me.txtOrigin = New System.Windows.Forms.TextBox
    Me.txtDestination = New System.Windows.Forms.TextBox
    Me.cmdPath = New System.Windows.Forms.Button
    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.UpdateProgress = New System.Windows.Forms.ToolStripProgressBar
    Me.ToolStripSplitButton1 = New System.Windows.Forms.ToolStripSplitButton
    Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
    Me.mainMenu = New System.Windows.Forms.MenuStrip
    Me.Popup1ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
    Me.RemoveAllNodeSourceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
    Me.ShowSourceNodeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
    Me.lblLyr2 = New System.Windows.Forms.Label
    Me.cmbNetwork = New System.Windows.Forms.ComboBox
    Me.cmbDEM = New System.Windows.Forms.ComboBox
    Me.lblLyr7 = New System.Windows.Forms.Label
    Me.TabMain.SuspendLayout()
    Me.TabPageNetworkGeneration.SuspendLayout()
    Me.TabControl1.SuspendLayout()
    Me.TabPage6.SuspendLayout()
    CType(Me.grdNode, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.TabPage7.SuspendLayout()
    CType(Me.grdLink, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.TabPageNetworkEditor.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    CType(Me.grdUnconnectedNode, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.StatusStrip1.SuspendLayout()
    Me.mainMenu.SuspendLayout()
    Me.SuspendLayout()
    '
    'TabMain
    '
    Me.TabMain.Controls.Add(Me.TabPageNetworkGeneration)
    Me.TabMain.Controls.Add(Me.TabPageNetworkEditor)
    Me.TabMain.Location = New System.Drawing.Point(0, 53)
    Me.TabMain.Name = "TabMain"
    Me.TabMain.SelectedIndex = 0
    Me.TabMain.Size = New System.Drawing.Size(387, 341)
    Me.TabMain.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight
    Me.TabMain.TabIndex = 0
    '
    'TabPageNetworkGeneration
    '
    Me.TabPageNetworkGeneration.BackColor = System.Drawing.Color.White
    Me.TabPageNetworkGeneration.Controls.Add(Me.TabControl1)
    Me.TabPageNetworkGeneration.Controls.Add(Me.cmdInitialNET)
    Me.TabPageNetworkGeneration.Controls.Add(Me.chkUseExistingNode)
    Me.TabPageNetworkGeneration.Controls.Add(Me.cmdAutoGen)
    Me.TabPageNetworkGeneration.Controls.Add(Me.cmdSelectNode)
    Me.TabPageNetworkGeneration.Controls.Add(Me.cmdCreateNode)
    Me.TabPageNetworkGeneration.ImageIndex = 2
    Me.TabPageNetworkGeneration.Location = New System.Drawing.Point(4, 22)
    Me.TabPageNetworkGeneration.Name = "TabPageNetworkGeneration"
    Me.TabPageNetworkGeneration.Size = New System.Drawing.Size(379, 315)
    Me.TabPageNetworkGeneration.TabIndex = 1
    Me.TabPageNetworkGeneration.Text = "Network"
    Me.TabPageNetworkGeneration.UseVisualStyleBackColor = True
    '
    'TabControl1
    '
    Me.TabControl1.Controls.Add(Me.TabPage6)
    Me.TabControl1.Controls.Add(Me.TabPage7)
    Me.TabControl1.Location = New System.Drawing.Point(0, 27)
    Me.TabControl1.Name = "TabControl1"
    Me.TabControl1.SelectedIndex = 0
    Me.TabControl1.Size = New System.Drawing.Size(380, 288)
    Me.TabControl1.TabIndex = 3
    '
    'TabPage6
    '
    Me.TabPage6.Controls.Add(Me.grdNode)
    Me.TabPage6.Location = New System.Drawing.Point(4, 22)
    Me.TabPage6.Name = "TabPage6"
    Me.TabPage6.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage6.Size = New System.Drawing.Size(372, 262)
    Me.TabPage6.TabIndex = 0
    Me.TabPage6.Text = "Node data"
    Me.TabPage6.UseVisualStyleBackColor = True
    '
    'grdNode
    '
    DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
    DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
    DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
    DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
    DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    Me.grdNode.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
    Me.grdNode.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.grdNode.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column6, Me.Column7, Me.Column8, Me.nodeTYP, Me.Column10, Me.Column23})
    DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
    DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
    DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
    DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
    DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
    Me.grdNode.DefaultCellStyle = DataGridViewCellStyle2
    Me.grdNode.Location = New System.Drawing.Point(3, 3)
    Me.grdNode.Name = "grdNode"
    Me.grdNode.Size = New System.Drawing.Size(366, 256)
    Me.grdNode.TabIndex = 0
    '
    'Column6
    '
    Me.Column6.HeaderText = "ID"
    Me.Column6.Name = "Column6"
    Me.Column6.Width = 50
    '
    'Column7
    '
    Me.Column7.HeaderText = "X"
    Me.Column7.Name = "Column7"
    Me.Column7.Width = 80
    '
    'Column8
    '
    Me.Column8.HeaderText = "Y"
    Me.Column8.Name = "Column8"
    Me.Column8.Width = 80
    '
    'nodeTYP
    '
    Me.nodeTYP.HeaderText = "Type"
    Me.nodeTYP.Items.AddRange(New Object() {"Source Node", "Junction Node"})
    Me.nodeTYP.Name = "nodeTYP"
    Me.nodeTYP.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
    Me.nodeTYP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
    '
    'Column10
    '
    Me.Column10.HeaderText = "Demand"
    Me.Column10.Name = "Column10"
    '
    'Column23
    '
    Me.Column23.HeaderText = "Z"
    Me.Column23.Name = "Column23"
    '
    'TabPage7
    '
    Me.TabPage7.Controls.Add(Me.grdLink)
    Me.TabPage7.Location = New System.Drawing.Point(4, 22)
    Me.TabPage7.Name = "TabPage7"
    Me.TabPage7.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage7.Size = New System.Drawing.Size(372, 262)
    Me.TabPage7.TabIndex = 1
    Me.TabPage7.Text = "Link data"
    Me.TabPage7.UseVisualStyleBackColor = True
    '
    'grdLink
    '
    DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
    DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
    DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
    DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
    DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    Me.grdLink.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
    Me.grdLink.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.grdLink.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5})
    DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
    DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
    DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
    DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
    DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
    DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
    Me.grdLink.DefaultCellStyle = DataGridViewCellStyle4
    Me.grdLink.Dock = System.Windows.Forms.DockStyle.Fill
    Me.grdLink.Location = New System.Drawing.Point(3, 3)
    Me.grdLink.Name = "grdLink"
    Me.grdLink.Size = New System.Drawing.Size(366, 256)
    Me.grdLink.TabIndex = 1
    '
    'DataGridViewTextBoxColumn1
    '
    Me.DataGridViewTextBoxColumn1.HeaderText = "ID"
    Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
    Me.DataGridViewTextBoxColumn1.Width = 50
    '
    'DataGridViewTextBoxColumn2
    '
    Me.DataGridViewTextBoxColumn2.HeaderText = "F"
    Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
    Me.DataGridViewTextBoxColumn2.Width = 50
    '
    'DataGridViewTextBoxColumn3
    '
    Me.DataGridViewTextBoxColumn3.HeaderText = "T"
    Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
    Me.DataGridViewTextBoxColumn3.Width = 50
    '
    'DataGridViewTextBoxColumn4
    '
    Me.DataGridViewTextBoxColumn4.HeaderText = "L (m)"
    Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
    '
    'DataGridViewTextBoxColumn5
    '
    Me.DataGridViewTextBoxColumn5.HeaderText = "D (mm)"
    Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
    '
    'cmdInitialNET
    '
    Me.cmdInitialNET.BackgroundImage = CType(resources.GetObject("cmdInitialNET.BackgroundImage"), System.Drawing.Image)
    Me.cmdInitialNET.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
    Me.cmdInitialNET.Location = New System.Drawing.Point(76, 2)
    Me.cmdInitialNET.Name = "cmdInitialNET"
    Me.cmdInitialNET.Size = New System.Drawing.Size(24, 24)
    Me.cmdInitialNET.TabIndex = 68
    Me.cmdInitialNET.Text = "ini"
    Me.ToolTip1.SetToolTip(Me.cmdInitialNET, "สร้างข้อมูล Network")
    Me.cmdInitialNET.UseVisualStyleBackColor = True
    '
    'chkUseExistingNode
    '
    Me.chkUseExistingNode.AutoSize = True
    Me.chkUseExistingNode.BackColor = System.Drawing.Color.Transparent
    Me.chkUseExistingNode.Location = New System.Drawing.Point(7, 299)
    Me.chkUseExistingNode.Name = "chkUseExistingNode"
    Me.chkUseExistingNode.Size = New System.Drawing.Size(110, 17)
    Me.chkUseExistingNode.TabIndex = 69
    Me.chkUseExistingNode.Text = "Use existing node"
    Me.chkUseExistingNode.UseVisualStyleBackColor = False
    '
    'cmdAutoGen
    '
    Me.cmdAutoGen.BackgroundImage = CType(resources.GetObject("cmdAutoGen.BackgroundImage"), System.Drawing.Image)
    Me.cmdAutoGen.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
    Me.cmdAutoGen.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.cmdAutoGen.Location = New System.Drawing.Point(3, 2)
    Me.cmdAutoGen.Name = "cmdAutoGen"
    Me.cmdAutoGen.Size = New System.Drawing.Size(24, 24)
    Me.cmdAutoGen.TabIndex = 61
    Me.ToolTip1.SetToolTip(Me.cmdAutoGen, "สร้างตาราง Node & Link")
    Me.cmdAutoGen.UseVisualStyleBackColor = True
    '
    'cmdSelectNode
    '
    Me.cmdSelectNode.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
    Me.cmdSelectNode.Image = CType(resources.GetObject("cmdSelectNode.Image"), System.Drawing.Image)
    Me.cmdSelectNode.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.cmdSelectNode.Location = New System.Drawing.Point(28, 2)
    Me.cmdSelectNode.Name = "cmdSelectNode"
    Me.cmdSelectNode.Size = New System.Drawing.Size(24, 24)
    Me.cmdSelectNode.TabIndex = 61
    Me.ToolTip1.SetToolTip(Me.cmdSelectNode, "Select Node")
    Me.cmdSelectNode.UseVisualStyleBackColor = True
    '
    'cmdCreateNode
    '
    Me.cmdCreateNode.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
    Me.cmdCreateNode.Image = CType(resources.GetObject("cmdCreateNode.Image"), System.Drawing.Image)
    Me.cmdCreateNode.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.cmdCreateNode.Location = New System.Drawing.Point(52, 2)
    Me.cmdCreateNode.Name = "cmdCreateNode"
    Me.cmdCreateNode.Size = New System.Drawing.Size(24, 24)
    Me.cmdCreateNode.TabIndex = 61
    Me.ToolTip1.SetToolTip(Me.cmdCreateNode, "สร้าง Node shapefile")
    Me.cmdCreateNode.UseVisualStyleBackColor = True
    '
    'TabPageNetworkEditor
    '
    Me.TabPageNetworkEditor.BackColor = System.Drawing.Color.White
    Me.TabPageNetworkEditor.Controls.Add(Me.lblErrorLink)
    Me.TabPageNetworkEditor.Controls.Add(Me.GroupBox2)
    Me.TabPageNetworkEditor.Location = New System.Drawing.Point(4, 22)
    Me.TabPageNetworkEditor.Name = "TabPageNetworkEditor"
    Me.TabPageNetworkEditor.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPageNetworkEditor.Size = New System.Drawing.Size(379, 315)
    Me.TabPageNetworkEditor.TabIndex = 6
    Me.TabPageNetworkEditor.Text = "Edit"
    Me.TabPageNetworkEditor.UseVisualStyleBackColor = True
    '
    'lblErrorLink
    '
    Me.lblErrorLink.AutoSize = True
    Me.lblErrorLink.Location = New System.Drawing.Point(5, 270)
    Me.lblErrorLink.Name = "lblErrorLink"
    Me.lblErrorLink.Size = New System.Drawing.Size(105, 13)
    Me.lblErrorLink.TabIndex = 75
    Me.lblErrorLink.Text = "Number of unlinkage"
    '
    'GroupBox2
    '
    Me.GroupBox2.Controls.Add(Me.cmdSaveUnLinkFile)
    Me.GroupBox2.Controls.Add(Me.cmdOpenUnLinkFile)
    Me.GroupBox2.Controls.Add(Me.chkDrawCheckingPath)
    Me.GroupBox2.Controls.Add(Me.cmdNetworkCheck)
    Me.GroupBox2.Controls.Add(Me.grdUnconnectedNode)
    Me.GroupBox2.Controls.Add(Me.cmdDrawIncompletedNetwork)
    Me.GroupBox2.Controls.Add(Me.cmdClearDrawing1)
    Me.GroupBox2.Location = New System.Drawing.Point(3, 6)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(179, 261)
    Me.GroupBox2.TabIndex = 73
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Network Check"
    '
    'cmdSaveUnLinkFile
    '
    Me.cmdSaveUnLinkFile.BackColor = System.Drawing.Color.White
    Me.cmdSaveUnLinkFile.BackgroundImage = CType(resources.GetObject("cmdSaveUnLinkFile.BackgroundImage"), System.Drawing.Image)
    Me.cmdSaveUnLinkFile.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
    Me.cmdSaveUnLinkFile.Location = New System.Drawing.Point(53, 19)
    Me.cmdSaveUnLinkFile.Name = "cmdSaveUnLinkFile"
    Me.cmdSaveUnLinkFile.Size = New System.Drawing.Size(24, 24)
    Me.cmdSaveUnLinkFile.TabIndex = 50
    Me.ToolTip1.SetToolTip(Me.cmdSaveUnLinkFile, "Save as default project")
    Me.cmdSaveUnLinkFile.UseVisualStyleBackColor = False
    '
    'cmdOpenUnLinkFile
    '
    Me.cmdOpenUnLinkFile.BackColor = System.Drawing.Color.White
    Me.cmdOpenUnLinkFile.BackgroundImage = CType(resources.GetObject("cmdOpenUnLinkFile.BackgroundImage"), System.Drawing.Image)
    Me.cmdOpenUnLinkFile.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
    Me.cmdOpenUnLinkFile.Location = New System.Drawing.Point(29, 19)
    Me.cmdOpenUnLinkFile.Name = "cmdOpenUnLinkFile"
    Me.cmdOpenUnLinkFile.Size = New System.Drawing.Size(24, 24)
    Me.cmdOpenUnLinkFile.TabIndex = 49
    Me.ToolTip1.SetToolTip(Me.cmdOpenUnLinkFile, "Open project")
    Me.cmdOpenUnLinkFile.UseVisualStyleBackColor = False
    '
    'chkDrawCheckingPath
    '
    Me.chkDrawCheckingPath.AutoSize = True
    Me.chkDrawCheckingPath.Location = New System.Drawing.Point(5, 43)
    Me.chkDrawCheckingPath.Name = "chkDrawCheckingPath"
    Me.chkDrawCheckingPath.Size = New System.Drawing.Size(97, 17)
    Me.chkDrawCheckingPath.TabIndex = 47
    Me.chkDrawCheckingPath.Text = "Draw flow path"
    Me.chkDrawCheckingPath.UseVisualStyleBackColor = True
    '
    'cmdNetworkCheck
    '
    Me.cmdNetworkCheck.BackColor = System.Drawing.Color.White
    Me.cmdNetworkCheck.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
    Me.cmdNetworkCheck.Enabled = False
    Me.cmdNetworkCheck.Image = CType(resources.GetObject("cmdNetworkCheck.Image"), System.Drawing.Image)
    Me.cmdNetworkCheck.ImageAlign = System.Drawing.ContentAlignment.TopLeft
    Me.cmdNetworkCheck.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.cmdNetworkCheck.Location = New System.Drawing.Point(5, 19)
    Me.cmdNetworkCheck.Name = "cmdNetworkCheck"
    Me.cmdNetworkCheck.Size = New System.Drawing.Size(24, 24)
    Me.cmdNetworkCheck.TabIndex = 45
    Me.ToolTip1.SetToolTip(Me.cmdNetworkCheck, "ตรวจสอบเส้นทางการไหล")
    Me.cmdNetworkCheck.UseVisualStyleBackColor = False
    '
    'grdUnconnectedNode
    '
    Me.grdUnconnectedNode.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.grdUnconnectedNode.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2})
    Me.grdUnconnectedNode.Location = New System.Drawing.Point(3, 61)
    Me.grdUnconnectedNode.Name = "grdUnconnectedNode"
    Me.grdUnconnectedNode.Size = New System.Drawing.Size(170, 193)
    Me.grdUnconnectedNode.TabIndex = 46
    '
    'Column1
    '
    Me.Column1.HeaderText = "ID"
    Me.Column1.Name = "Column1"
    Me.Column1.Width = 50
    '
    'Column2
    '
    Me.Column2.HeaderText = "Status"
    Me.Column2.Name = "Column2"
    Me.Column2.Width = 60
    '
    'cmdDrawIncompletedNetwork
    '
    Me.cmdDrawIncompletedNetwork.BackColor = System.Drawing.Color.White
    Me.cmdDrawIncompletedNetwork.BackgroundImage = CType(resources.GetObject("cmdDrawIncompletedNetwork.BackgroundImage"), System.Drawing.Image)
    Me.cmdDrawIncompletedNetwork.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
    Me.cmdDrawIncompletedNetwork.ImageAlign = System.Drawing.ContentAlignment.TopLeft
    Me.cmdDrawIncompletedNetwork.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.cmdDrawIncompletedNetwork.Location = New System.Drawing.Point(78, 19)
    Me.cmdDrawIncompletedNetwork.Name = "cmdDrawIncompletedNetwork"
    Me.cmdDrawIncompletedNetwork.Size = New System.Drawing.Size(24, 24)
    Me.cmdDrawIncompletedNetwork.TabIndex = 44
    Me.cmdDrawIncompletedNetwork.UseVisualStyleBackColor = False
    '
    'cmdClearDrawing1
    '
    Me.cmdClearDrawing1.BackColor = System.Drawing.Color.White
    Me.cmdClearDrawing1.BackgroundImage = CType(resources.GetObject("cmdClearDrawing1.BackgroundImage"), System.Drawing.Image)
    Me.cmdClearDrawing1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
    Me.cmdClearDrawing1.ImageAlign = System.Drawing.ContentAlignment.TopLeft
    Me.cmdClearDrawing1.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.cmdClearDrawing1.Location = New System.Drawing.Point(103, 19)
    Me.cmdClearDrawing1.Name = "cmdClearDrawing1"
    Me.cmdClearDrawing1.Size = New System.Drawing.Size(24, 24)
    Me.cmdClearDrawing1.TabIndex = 45
    Me.cmdClearDrawing1.UseVisualStyleBackColor = False
    '
    'RadioButton2
    '
    Me.RadioButton2.AutoSize = True
    Me.RadioButton2.Location = New System.Drawing.Point(405, 179)
    Me.RadioButton2.Name = "RadioButton2"
    Me.RadioButton2.Size = New System.Drawing.Size(80, 17)
    Me.RadioButton2.TabIndex = 24
    Me.RadioButton2.TabStop = True
    Me.RadioButton2.Text = "CUS. Meter"
    Me.RadioButton2.UseVisualStyleBackColor = True
    '
    'RadioButton1
    '
    Me.RadioButton1.AutoSize = True
    Me.RadioButton1.Location = New System.Drawing.Point(405, 158)
    Me.RadioButton1.Name = "RadioButton1"
    Me.RadioButton1.Size = New System.Drawing.Size(88, 17)
    Me.RadioButton1.TabIndex = 23
    Me.RadioButton1.TabStop = True
    Me.RadioButton1.Text = "Source Node"
    Me.RadioButton1.UseVisualStyleBackColor = True
    '
    'ListBox1
    '
    Me.ListBox1.FormattingEnabled = True
    Me.ListBox1.Location = New System.Drawing.Point(403, 237)
    Me.ListBox1.Margin = New System.Windows.Forms.Padding(2)
    Me.ListBox1.Name = "ListBox1"
    Me.ListBox1.Size = New System.Drawing.Size(127, 43)
    Me.ListBox1.TabIndex = 22
    '
    'txtOrigin
    '
    Me.txtOrigin.Location = New System.Drawing.Point(498, 158)
    Me.txtOrigin.Margin = New System.Windows.Forms.Padding(2)
    Me.txtOrigin.Name = "txtOrigin"
    Me.txtOrigin.Size = New System.Drawing.Size(33, 20)
    Me.txtOrigin.TabIndex = 20
    Me.txtOrigin.Text = "1"
    '
    'txtDestination
    '
    Me.txtDestination.Location = New System.Drawing.Point(498, 178)
    Me.txtDestination.Margin = New System.Windows.Forms.Padding(2)
    Me.txtDestination.Name = "txtDestination"
    Me.txtDestination.Size = New System.Drawing.Size(33, 20)
    Me.txtDestination.TabIndex = 21
    Me.txtDestination.Text = "20"
    '
    'cmdPath
    '
    Me.cmdPath.Location = New System.Drawing.Point(396, 205)
    Me.cmdPath.Margin = New System.Windows.Forms.Padding(2)
    Me.cmdPath.Name = "cmdPath"
    Me.cmdPath.Size = New System.Drawing.Size(134, 28)
    Me.cmdPath.TabIndex = 19
    Me.cmdPath.Text = "Path"
    Me.cmdPath.UseVisualStyleBackColor = True
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UpdateProgress, Me.ToolStripSplitButton1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 381)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(726, 22)
    Me.StatusStrip1.TabIndex = 1
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'UpdateProgress
    '
    Me.UpdateProgress.Name = "UpdateProgress"
    Me.UpdateProgress.Size = New System.Drawing.Size(300, 16)
    Me.UpdateProgress.Visible = False
    '
    'ToolStripSplitButton1
    '
    Me.ToolStripSplitButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
    Me.ToolStripSplitButton1.Image = CType(resources.GetObject("ToolStripSplitButton1.Image"), System.Drawing.Image)
    Me.ToolStripSplitButton1.ImageTransparentColor = System.Drawing.Color.Magenta
    Me.ToolStripSplitButton1.Name = "ToolStripSplitButton1"
    Me.ToolStripSplitButton1.Size = New System.Drawing.Size(32, 20)
    Me.ToolStripSplitButton1.Text = "ToolStripSplitButton1"
    Me.ToolStripSplitButton1.ToolTipText = "Show detail"
    '
    'ImageList1
    '
    Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
    Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
    Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
    '
    'mainMenu
    '
    Me.mainMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Popup1ToolStripMenuItem})
    Me.mainMenu.Location = New System.Drawing.Point(0, 0)
    Me.mainMenu.Name = "mainMenu"
    Me.mainMenu.Size = New System.Drawing.Size(726, 24)
    Me.mainMenu.TabIndex = 26
    Me.mainMenu.Text = "MenuStrip1"
    Me.mainMenu.Visible = False
    '
    'Popup1ToolStripMenuItem
    '
    Me.Popup1ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RemoveAllNodeSourceToolStripMenuItem, Me.ShowSourceNodeToolStripMenuItem})
    Me.Popup1ToolStripMenuItem.Name = "Popup1ToolStripMenuItem"
    Me.Popup1ToolStripMenuItem.Size = New System.Drawing.Size(60, 20)
    Me.Popup1ToolStripMenuItem.Text = "popup1"
    '
    'RemoveAllNodeSourceToolStripMenuItem
    '
    Me.RemoveAllNodeSourceToolStripMenuItem.Name = "RemoveAllNodeSourceToolStripMenuItem"
    Me.RemoveAllNodeSourceToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
    Me.RemoveAllNodeSourceToolStripMenuItem.Text = "Remove all node source"
    '
    'ShowSourceNodeToolStripMenuItem
    '
    Me.ShowSourceNodeToolStripMenuItem.Name = "ShowSourceNodeToolStripMenuItem"
    Me.ShowSourceNodeToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
    Me.ShowSourceNodeToolStripMenuItem.Text = "Show source node"
    '
    'lblLyr2
    '
    Me.lblLyr2.AutoSize = True
    Me.lblLyr2.Location = New System.Drawing.Point(1, 6)
    Me.lblLyr2.Name = "lblLyr2"
    Me.lblLyr2.Size = New System.Drawing.Size(28, 13)
    Me.lblLyr2.TabIndex = 29
    Me.lblLyr2.Text = "Pipe"
    '
    'cmbNetwork
    '
    Me.cmbNetwork.FormattingEnabled = True
    Me.cmbNetwork.Location = New System.Drawing.Point(38, 3)
    Me.cmbNetwork.Name = "cmbNetwork"
    Me.cmbNetwork.Size = New System.Drawing.Size(183, 21)
    Me.cmbNetwork.TabIndex = 30
    '
    'cmbDEM
    '
    Me.cmbDEM.FormattingEnabled = True
    Me.cmbDEM.Location = New System.Drawing.Point(38, 26)
    Me.cmbDEM.Name = "cmbDEM"
    Me.cmbDEM.Size = New System.Drawing.Size(183, 21)
    Me.cmbDEM.TabIndex = 28
    '
    'lblLyr7
    '
    Me.lblLyr7.AutoSize = True
    Me.lblLyr7.Location = New System.Drawing.Point(1, 26)
    Me.lblLyr7.Name = "lblLyr7"
    Me.lblLyr7.Size = New System.Drawing.Size(31, 13)
    Me.lblLyr7.TabIndex = 27
    Me.lblLyr7.Text = "DEM"
    '
    'frmNetGen
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
    Me.ClientSize = New System.Drawing.Size(726, 403)
    Me.Controls.Add(Me.lblLyr2)
    Me.Controls.Add(Me.cmbNetwork)
    Me.Controls.Add(Me.cmbDEM)
    Me.Controls.Add(Me.lblLyr7)
    Me.Controls.Add(Me.RadioButton2)
    Me.Controls.Add(Me.TabMain)
    Me.Controls.Add(Me.RadioButton1)
    Me.Controls.Add(Me.ListBox1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.mainMenu)
    Me.Controls.Add(Me.txtOrigin)
    Me.Controls.Add(Me.txtDestination)
    Me.Controls.Add(Me.cmdPath)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
    Me.MainMenuStrip = Me.mainMenu
    Me.Name = "frmNetGen"
    Me.TopMost = True
    Me.TabMain.ResumeLayout(False)
    Me.TabPageNetworkGeneration.ResumeLayout(False)
    Me.TabPageNetworkGeneration.PerformLayout()
    Me.TabControl1.ResumeLayout(False)
    Me.TabPage6.ResumeLayout(False)
    CType(Me.grdNode, System.ComponentModel.ISupportInitialize).EndInit()
    Me.TabPage7.ResumeLayout(False)
    CType(Me.grdLink, System.ComponentModel.ISupportInitialize).EndInit()
    Me.TabPageNetworkEditor.ResumeLayout(False)
    Me.TabPageNetworkEditor.PerformLayout()
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    CType(Me.grdUnconnectedNode, System.ComponentModel.ISupportInitialize).EndInit()
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.mainMenu.ResumeLayout(False)
    Me.mainMenu.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TabMain As System.Windows.Forms.TabControl
  Friend WithEvents TabPageNetworkGeneration As System.Windows.Forms.TabPage
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
  Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
  Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
  Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents UpdateProgress As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents grdNode As System.Windows.Forms.DataGridView
  Friend WithEvents grdLink As System.Windows.Forms.DataGridView
  Friend WithEvents cmdCreateNode As System.Windows.Forms.Button
  Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents cmdAutoGen As System.Windows.Forms.Button
    Friend WithEvents cmdInitialNET As System.Windows.Forms.Button
  Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
  Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
  Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
  Friend WithEvents txtOrigin As System.Windows.Forms.TextBox
  Friend WithEvents txtDestination As System.Windows.Forms.TextBox
  Friend WithEvents cmdPath As System.Windows.Forms.Button
  Friend WithEvents chkUseExistingNode As System.Windows.Forms.CheckBox
  Friend WithEvents ToolStripSplitButton1 As System.Windows.Forms.ToolStripSplitButton
    Friend WithEvents cmdNetworkCheck As System.Windows.Forms.Button
    Friend WithEvents TabPageNetworkEditor As System.Windows.Forms.TabPage
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents grdUnconnectedNode As System.Windows.Forms.DataGridView
  Friend WithEvents chkDrawCheckingPath As System.Windows.Forms.CheckBox
  Friend WithEvents cmdSaveUnLinkFile As System.Windows.Forms.Button
  Friend WithEvents cmdOpenUnLinkFile As System.Windows.Forms.Button
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents Column8 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents nodeTYP As System.Windows.Forms.DataGridViewComboBoxColumn
  Friend WithEvents Column10 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents Column23 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents cmdSelectNode As System.Windows.Forms.Button
  Friend WithEvents cmdDrawIncompletedNetwork As System.Windows.Forms.Button
  Friend WithEvents cmdClearDrawing1 As System.Windows.Forms.Button
  Friend WithEvents lblErrorLink As System.Windows.Forms.Label
  Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
  Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents mainMenu As System.Windows.Forms.MenuStrip
  Friend WithEvents Popup1ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents RemoveAllNodeSourceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ShowSourceNodeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblLyr2 As System.Windows.Forms.Label
  Friend WithEvents cmbNetwork As System.Windows.Forms.ComboBox
  Friend WithEvents cmbDEM As System.Windows.Forms.ComboBox
  Friend WithEvents lblLyr7 As System.Windows.Forms.Label

End Class
