Imports System.Windows.Forms
Imports System.IO
Imports System.Collections.Generic


Public Class frmMain
   Inherits System.Windows.Forms.Form
   '-----------------------------------------------------------
   '
   'Dim HashNode As Hashtable = New Hashtable()
   'Dim HashNode As DictionaryEntry

   Private m_Connection As OleDb.OleDbConnection
   Private m_Command As OleDb.OleDbCommand
   Private m_CommandBuilder As OleDb.OleDbCommandBuilder
   Private m_Adapter As OleDb.OleDbDataAdapter
   Private m_selected As DataTable
   Private m_sf As MapWinGIS.Shapefile
   Private m_owner As Windows.Forms.Form
   Private m_oldRow, m_oldCol As Integer
   Friend m_TrulyChanged As Boolean = False
   Private Loaded As Boolean = False
   Private resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
   Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
   'Private m_Map As MapWindow.Interfaces.IMapWin
   Private m_MapWindowForm As Windows.Forms.Form


   Dim grdTopo As New MapWinGIS.Grid
   Dim grdCU As New MapWinGIS.Grid
   Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
   Friend WithEvents Label5 As System.Windows.Forms.Label
   Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
   Friend WithEvents cmdReport As System.Windows.Forms.Button
   Friend WithEvents cmdLinkMeter2Buld As System.Windows.Forms.Button
   Friend WithEvents cmbShowDiffernce As System.Windows.Forms.Button
   Friend WithEvents cmbSwitchDIR As System.Windows.Forms.Button
   Friend WithEvents cmdLinkBld As System.Windows.Forms.Button
   Friend WithEvents cmdMoveNode As System.Windows.Forms.Button
   Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
   Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
   Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
   Friend WithEvents cmbMeter4Model As System.Windows.Forms.ComboBox
   Friend WithEvents cmbMeterField4Bld As System.Windows.Forms.ComboBox
   Friend WithEvents cmbMeterField4Pipe As System.Windows.Forms.ComboBox
   Friend WithEvents Label13 As System.Windows.Forms.Label
   Friend WithEvents Label11 As System.Windows.Forms.Label
   Friend WithEvents cmbPipe4Model As System.Windows.Forms.ComboBox
   Friend WithEvents cmbBuilding As System.Windows.Forms.ComboBox
   Friend WithEvents cmbPipeField As System.Windows.Forms.ComboBox
   Friend WithEvents cmbBuildingField As System.Windows.Forms.ComboBox
   Friend WithEvents Label12 As System.Windows.Forms.Label
   Friend WithEvents Label10 As System.Windows.Forms.Label
   Friend WithEvents CmbValve As System.Windows.Forms.ComboBox
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents cmbDemand_Node As System.Windows.Forms.ComboBox
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
   Friend WithEvents TabControl3 As System.Windows.Forms.TabControl
   Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
   Friend WithEvents grdNodeList As System.Windows.Forms.DataGridView
   Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Column8 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Column15 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents TabPage9 As System.Windows.Forms.TabPage
   Friend WithEvents grdLinkList As System.Windows.Forms.DataGridView
   Friend WithEvents Column9 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Column10 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Column11 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Column12 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Column13 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents Column14 As System.Windows.Forms.DataGridViewTextBoxColumn
   Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
   Friend WithEvents cmbZonal_Ref As System.Windows.Forms.ComboBox
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents cmdAddJunction As System.Windows.Forms.Button
   Friend WithEvents cmdDeleteJunction As System.Windows.Forms.Button
   Friend WithEvents cmdClearDrawing As System.Windows.Forms.Button
   Friend WithEvents cmdDrawWettedArea As System.Windows.Forms.Button
   Friend WithEvents cmdPlotDemandRef As System.Windows.Forms.Button
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents cmbZone As System.Windows.Forms.ComboBox
   Friend WithEvents cmdMeterCount As System.Windows.Forms.Button
   Friend WithEvents cmdNLtable As System.Windows.Forms.Button
   Friend WithEvents cmdGetLineEdge As System.Windows.Forms.Button
   Friend WithEvents cmdSplitLink As System.Windows.Forms.Button
   Friend WithEvents cmdSplitPt As System.Windows.Forms.Button
   Friend WithEvents cdmPipeID As System.Windows.Forms.Button
   Friend WithEvents cmdGetpoint As System.Windows.Forms.Button
   Friend WithEvents Button1 As System.Windows.Forms.Button
   'As Double
   Dim TopoFilename As String = ""


   <CLSCompliant(False)> _
   Public Sub New(ByRef Shapefile As MapWinGIS.Shapefile, ByVal OwnerForm As Windows.Forms.Form)
      MyBase.New()

      'This call is required by the Windows Form Designer.
      InitializeComponent()

      'Add any initialization after the InitializeComponent() call
      m_sf = Shapefile
      m_owner = OwnerForm
      m_oldRow = -1
      m_oldCol = -1
      Initialize(m_sf, m_owner)
   End Sub

#Region " Windows Form Designer generated code "

   'Form overrides dispose to clean up the component list.
   Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
      If disposing Then
         If Not (components Is Nothing) Then
            components.Dispose()
         End If
      End If
      MyBase.Dispose(disposing)
   End Sub

   'Required by the Windows Form Designer
   Private components As System.ComponentModel.IContainer

   'NOTE: The following procedure is required by the Windows Form Designer
   'It can be modified using the Windows Form Designer.  
   'Do not modify it using the code editor.
   Friend WithEvents LayoutEditorHelp As System.Windows.Forms.HelpProvider
   Friend WithEvents UpdateProgress As System.Windows.Forms.ProgressBar
   <System.Diagnostics.DebuggerStepThrough()> _
   Private Sub InitializeComponent()
      Me.components = New System.ComponentModel.Container
      Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
      Me.LayoutEditorHelp = New System.Windows.Forms.HelpProvider
      Me.UpdateProgress = New System.Windows.Forms.ProgressBar
      Me.cmdReport = New System.Windows.Forms.Button
      Me.PictureBox3 = New System.Windows.Forms.PictureBox
      Me.cmdLinkMeter2Buld = New System.Windows.Forms.Button
      Me.cmbShowDiffernce = New System.Windows.Forms.Button
      Me.cmbSwitchDIR = New System.Windows.Forms.Button
      Me.cmdLinkBld = New System.Windows.Forms.Button
      Me.cmdMoveNode = New System.Windows.Forms.Button
      Me.TabPage4 = New System.Windows.Forms.TabPage
      Me.TabPage1 = New System.Windows.Forms.TabPage
      Me.GroupBox1 = New System.Windows.Forms.GroupBox
      Me.cmbMeter4Model = New System.Windows.Forms.ComboBox
      Me.cmbMeterField4Bld = New System.Windows.Forms.ComboBox
      Me.cmbMeterField4Pipe = New System.Windows.Forms.ComboBox
      Me.Label13 = New System.Windows.Forms.Label
      Me.Label11 = New System.Windows.Forms.Label
      Me.cmbPipe4Model = New System.Windows.Forms.ComboBox
      Me.cmbBuilding = New System.Windows.Forms.ComboBox
      Me.cmbPipeField = New System.Windows.Forms.ComboBox
      Me.cmbBuildingField = New System.Windows.Forms.ComboBox
      Me.Label12 = New System.Windows.Forms.Label
      Me.Label10 = New System.Windows.Forms.Label
      Me.cmbZonal_Ref = New System.Windows.Forms.ComboBox
      Me.Label2 = New System.Windows.Forms.Label
      Me.Label1 = New System.Windows.Forms.Label
      Me.cmbZone = New System.Windows.Forms.ComboBox
      Me.CmbValve = New System.Windows.Forms.ComboBox
      Me.Label4 = New System.Windows.Forms.Label
      Me.cmbDemand_Node = New System.Windows.Forms.ComboBox
      Me.Label3 = New System.Windows.Forms.Label
      Me.TabPage7 = New System.Windows.Forms.TabPage
      Me.Button1 = New System.Windows.Forms.Button
      Me.cmdGetpoint = New System.Windows.Forms.Button
      Me.cdmPipeID = New System.Windows.Forms.Button
      Me.cmdSplitPt = New System.Windows.Forms.Button
      Me.cmdSplitLink = New System.Windows.Forms.Button
      Me.cmdGetLineEdge = New System.Windows.Forms.Button
      Me.cmdNLtable = New System.Windows.Forms.Button
      Me.cmdClearDrawing = New System.Windows.Forms.Button
      Me.cmdDrawWettedArea = New System.Windows.Forms.Button
      Me.cmdAddJunction = New System.Windows.Forms.Button
      Me.cmdDeleteJunction = New System.Windows.Forms.Button
      Me.TabControl3 = New System.Windows.Forms.TabControl
      Me.TabPage8 = New System.Windows.Forms.TabPage
      Me.grdNodeList = New System.Windows.Forms.DataGridView
      Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.Column8 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.TabPage9 = New System.Windows.Forms.TabPage
      Me.grdLinkList = New System.Windows.Forms.DataGridView
      Me.Column9 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.Column11 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.Column12 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.Column13 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn
      Me.cmdMeterCount = New System.Windows.Forms.Button
      Me.cmdPlotDemandRef = New System.Windows.Forms.Button
      Me.TabControl1 = New System.Windows.Forms.TabControl
      Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
      Me.PictureBox1 = New System.Windows.Forms.PictureBox
      Me.Label5 = New System.Windows.Forms.Label
      CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.TabPage1.SuspendLayout()
      Me.GroupBox1.SuspendLayout()
      Me.TabPage7.SuspendLayout()
      Me.TabControl3.SuspendLayout()
      Me.TabPage8.SuspendLayout()
      CType(Me.grdNodeList, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.TabPage9.SuspendLayout()
      CType(Me.grdLinkList, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.TabControl1.SuspendLayout()
      CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'UpdateProgress
      '
      resources.ApplyResources(Me.UpdateProgress, "UpdateProgress")
      Me.UpdateProgress.Name = "UpdateProgress"
      Me.LayoutEditorHelp.SetShowHelp(Me.UpdateProgress, CType(resources.GetObject("UpdateProgress.ShowHelp"), Boolean))
      '
      'cmdReport
      '
      resources.ApplyResources(Me.cmdReport, "cmdReport")
      Me.cmdReport.Name = "cmdReport"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdReport, CType(resources.GetObject("cmdReport.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdReport, resources.GetString("cmdReport.ToolTip"))
      Me.cmdReport.UseVisualStyleBackColor = True
      '
      'PictureBox3
      '
      resources.ApplyResources(Me.PictureBox3, "PictureBox3")
      Me.PictureBox3.Name = "PictureBox3"
      Me.LayoutEditorHelp.SetShowHelp(Me.PictureBox3, CType(resources.GetObject("PictureBox3.ShowHelp"), Boolean))
      Me.PictureBox3.TabStop = False
      '
      'cmdLinkMeter2Buld
      '
      resources.ApplyResources(Me.cmdLinkMeter2Buld, "cmdLinkMeter2Buld")
      Me.cmdLinkMeter2Buld.Name = "cmdLinkMeter2Buld"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdLinkMeter2Buld, CType(resources.GetObject("cmdLinkMeter2Buld.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdLinkMeter2Buld, resources.GetString("cmdLinkMeter2Buld.ToolTip"))
      Me.cmdLinkMeter2Buld.UseVisualStyleBackColor = True
      '
      'cmbShowDiffernce
      '
      resources.ApplyResources(Me.cmbShowDiffernce, "cmbShowDiffernce")
      Me.cmbShowDiffernce.Name = "cmbShowDiffernce"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbShowDiffernce, CType(resources.GetObject("cmbShowDiffernce.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmbShowDiffernce, resources.GetString("cmbShowDiffernce.ToolTip"))
      Me.cmbShowDiffernce.UseVisualStyleBackColor = True
      '
      'cmbSwitchDIR
      '
      resources.ApplyResources(Me.cmbSwitchDIR, "cmbSwitchDIR")
      Me.cmbSwitchDIR.Name = "cmbSwitchDIR"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbSwitchDIR, CType(resources.GetObject("cmbSwitchDIR.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmbSwitchDIR, resources.GetString("cmbSwitchDIR.ToolTip"))
      Me.cmbSwitchDIR.UseVisualStyleBackColor = True
      '
      'cmdLinkBld
      '
      resources.ApplyResources(Me.cmdLinkBld, "cmdLinkBld")
      Me.cmdLinkBld.Name = "cmdLinkBld"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdLinkBld, CType(resources.GetObject("cmdLinkBld.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdLinkBld, resources.GetString("cmdLinkBld.ToolTip"))
      Me.cmdLinkBld.UseVisualStyleBackColor = True
      '
      'cmdMoveNode
      '
      Me.cmdMoveNode.BackColor = System.Drawing.Color.White
      resources.ApplyResources(Me.cmdMoveNode, "cmdMoveNode")
      Me.cmdMoveNode.Name = "cmdMoveNode"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdMoveNode, CType(resources.GetObject("cmdMoveNode.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdMoveNode, resources.GetString("cmdMoveNode.ToolTip"))
      Me.cmdMoveNode.UseVisualStyleBackColor = False
      '
      'TabPage4
      '
      resources.ApplyResources(Me.TabPage4, "TabPage4")
      Me.TabPage4.Name = "TabPage4"
      Me.LayoutEditorHelp.SetShowHelp(Me.TabPage4, CType(resources.GetObject("TabPage4.ShowHelp"), Boolean))
      Me.TabPage4.UseVisualStyleBackColor = True
      '
      'TabPage1
      '
      Me.TabPage1.Controls.Add(Me.GroupBox1)
      resources.ApplyResources(Me.TabPage1, "TabPage1")
      Me.TabPage1.Name = "TabPage1"
      Me.LayoutEditorHelp.SetShowHelp(Me.TabPage1, CType(resources.GetObject("TabPage1.ShowHelp"), Boolean))
      Me.TabPage1.UseVisualStyleBackColor = True
      '
      'GroupBox1
      '
      Me.GroupBox1.Controls.Add(Me.cmbMeter4Model)
      Me.GroupBox1.Controls.Add(Me.cmbMeterField4Bld)
      Me.GroupBox1.Controls.Add(Me.cmbMeterField4Pipe)
      Me.GroupBox1.Controls.Add(Me.Label13)
      Me.GroupBox1.Controls.Add(Me.Label11)
      Me.GroupBox1.Controls.Add(Me.cmbPipe4Model)
      Me.GroupBox1.Controls.Add(Me.cmbBuilding)
      Me.GroupBox1.Controls.Add(Me.cmbPipeField)
      Me.GroupBox1.Controls.Add(Me.cmbBuildingField)
      Me.GroupBox1.Controls.Add(Me.Label12)
      Me.GroupBox1.Controls.Add(Me.Label10)
      Me.GroupBox1.Controls.Add(Me.cmbZonal_Ref)
      Me.GroupBox1.Controls.Add(Me.Label2)
      Me.GroupBox1.Controls.Add(Me.Label1)
      Me.GroupBox1.Controls.Add(Me.cmbZone)
      Me.GroupBox1.Controls.Add(Me.CmbValve)
      Me.GroupBox1.Controls.Add(Me.Label4)
      Me.GroupBox1.Controls.Add(Me.cmbDemand_Node)
      Me.GroupBox1.Controls.Add(Me.Label3)
      resources.ApplyResources(Me.GroupBox1, "GroupBox1")
      Me.GroupBox1.Name = "GroupBox1"
      Me.LayoutEditorHelp.SetShowHelp(Me.GroupBox1, CType(resources.GetObject("GroupBox1.ShowHelp"), Boolean))
      Me.GroupBox1.TabStop = False
      '
      'cmbMeter4Model
      '
      Me.cmbMeter4Model.FormattingEnabled = True
      resources.ApplyResources(Me.cmbMeter4Model, "cmbMeter4Model")
      Me.cmbMeter4Model.Name = "cmbMeter4Model"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbMeter4Model, CType(resources.GetObject("cmbMeter4Model.ShowHelp"), Boolean))
      '
      'cmbMeterField4Bld
      '
      Me.cmbMeterField4Bld.FormattingEnabled = True
      resources.ApplyResources(Me.cmbMeterField4Bld, "cmbMeterField4Bld")
      Me.cmbMeterField4Bld.Name = "cmbMeterField4Bld"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbMeterField4Bld, CType(resources.GetObject("cmbMeterField4Bld.ShowHelp"), Boolean))
      '
      'cmbMeterField4Pipe
      '
      Me.cmbMeterField4Pipe.FormattingEnabled = True
      resources.ApplyResources(Me.cmbMeterField4Pipe, "cmbMeterField4Pipe")
      Me.cmbMeterField4Pipe.Name = "cmbMeterField4Pipe"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbMeterField4Pipe, CType(resources.GetObject("cmbMeterField4Pipe.ShowHelp"), Boolean))
      '
      'Label13
      '
      resources.ApplyResources(Me.Label13, "Label13")
      Me.Label13.Name = "Label13"
      Me.LayoutEditorHelp.SetShowHelp(Me.Label13, CType(resources.GetObject("Label13.ShowHelp"), Boolean))
      '
      'Label11
      '
      resources.ApplyResources(Me.Label11, "Label11")
      Me.Label11.Name = "Label11"
      Me.LayoutEditorHelp.SetShowHelp(Me.Label11, CType(resources.GetObject("Label11.ShowHelp"), Boolean))
      '
      'cmbPipe4Model
      '
      Me.cmbPipe4Model.FormattingEnabled = True
      resources.ApplyResources(Me.cmbPipe4Model, "cmbPipe4Model")
      Me.cmbPipe4Model.Name = "cmbPipe4Model"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbPipe4Model, CType(resources.GetObject("cmbPipe4Model.ShowHelp"), Boolean))
      '
      'cmbBuilding
      '
      Me.cmbBuilding.FormattingEnabled = True
      resources.ApplyResources(Me.cmbBuilding, "cmbBuilding")
      Me.cmbBuilding.Name = "cmbBuilding"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbBuilding, CType(resources.GetObject("cmbBuilding.ShowHelp"), Boolean))
      '
      'cmbPipeField
      '
      Me.cmbPipeField.FormattingEnabled = True
      resources.ApplyResources(Me.cmbPipeField, "cmbPipeField")
      Me.cmbPipeField.Name = "cmbPipeField"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbPipeField, CType(resources.GetObject("cmbPipeField.ShowHelp"), Boolean))
      '
      'cmbBuildingField
      '
      Me.cmbBuildingField.FormattingEnabled = True
      resources.ApplyResources(Me.cmbBuildingField, "cmbBuildingField")
      Me.cmbBuildingField.Name = "cmbBuildingField"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbBuildingField, CType(resources.GetObject("cmbBuildingField.ShowHelp"), Boolean))
      '
      'Label12
      '
      resources.ApplyResources(Me.Label12, "Label12")
      Me.Label12.Name = "Label12"
      Me.LayoutEditorHelp.SetShowHelp(Me.Label12, CType(resources.GetObject("Label12.ShowHelp"), Boolean))
      '
      'Label10
      '
      resources.ApplyResources(Me.Label10, "Label10")
      Me.Label10.Name = "Label10"
      Me.LayoutEditorHelp.SetShowHelp(Me.Label10, CType(resources.GetObject("Label10.ShowHelp"), Boolean))
      '
      'cmbZonal_Ref
      '
      Me.cmbZonal_Ref.FormattingEnabled = True
      resources.ApplyResources(Me.cmbZonal_Ref, "cmbZonal_Ref")
      Me.cmbZonal_Ref.Name = "cmbZonal_Ref"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbZonal_Ref, CType(resources.GetObject("cmbZonal_Ref.ShowHelp"), Boolean))
      '
      'Label2
      '
      resources.ApplyResources(Me.Label2, "Label2")
      Me.Label2.Name = "Label2"
      Me.LayoutEditorHelp.SetShowHelp(Me.Label2, CType(resources.GetObject("Label2.ShowHelp"), Boolean))
      '
      'Label1
      '
      resources.ApplyResources(Me.Label1, "Label1")
      Me.Label1.Name = "Label1"
      Me.LayoutEditorHelp.SetShowHelp(Me.Label1, CType(resources.GetObject("Label1.ShowHelp"), Boolean))
      '
      'cmbZone
      '
      Me.cmbZone.FormattingEnabled = True
      resources.ApplyResources(Me.cmbZone, "cmbZone")
      Me.cmbZone.Name = "cmbZone"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbZone, CType(resources.GetObject("cmbZone.ShowHelp"), Boolean))
      '
      'CmbValve
      '
      Me.CmbValve.FormattingEnabled = True
      resources.ApplyResources(Me.CmbValve, "CmbValve")
      Me.CmbValve.Name = "CmbValve"
      Me.LayoutEditorHelp.SetShowHelp(Me.CmbValve, CType(resources.GetObject("CmbValve.ShowHelp"), Boolean))
      '
      'Label4
      '
      resources.ApplyResources(Me.Label4, "Label4")
      Me.Label4.Name = "Label4"
      Me.LayoutEditorHelp.SetShowHelp(Me.Label4, CType(resources.GetObject("Label4.ShowHelp"), Boolean))
      '
      'cmbDemand_Node
      '
      Me.cmbDemand_Node.FormattingEnabled = True
      resources.ApplyResources(Me.cmbDemand_Node, "cmbDemand_Node")
      Me.cmbDemand_Node.Name = "cmbDemand_Node"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmbDemand_Node, CType(resources.GetObject("cmbDemand_Node.ShowHelp"), Boolean))
      '
      'Label3
      '
      resources.ApplyResources(Me.Label3, "Label3")
      Me.Label3.Name = "Label3"
      Me.LayoutEditorHelp.SetShowHelp(Me.Label3, CType(resources.GetObject("Label3.ShowHelp"), Boolean))
      '
      'TabPage7
      '
      Me.TabPage7.Controls.Add(Me.Button1)
      Me.TabPage7.Controls.Add(Me.cmdGetpoint)
      Me.TabPage7.Controls.Add(Me.cdmPipeID)
      Me.TabPage7.Controls.Add(Me.cmdSplitPt)
      Me.TabPage7.Controls.Add(Me.cmdSplitLink)
      Me.TabPage7.Controls.Add(Me.cmdGetLineEdge)
      Me.TabPage7.Controls.Add(Me.cmdNLtable)
      Me.TabPage7.Controls.Add(Me.cmdClearDrawing)
      Me.TabPage7.Controls.Add(Me.cmdDrawWettedArea)
      Me.TabPage7.Controls.Add(Me.cmdAddJunction)
      Me.TabPage7.Controls.Add(Me.cmdDeleteJunction)
      Me.TabPage7.Controls.Add(Me.cmdMoveNode)
      Me.TabPage7.Controls.Add(Me.TabControl3)
      Me.TabPage7.Controls.Add(Me.cmdLinkMeter2Buld)
      Me.TabPage7.Controls.Add(Me.cmbShowDiffernce)
      Me.TabPage7.Controls.Add(Me.cmdReport)
      Me.TabPage7.Controls.Add(Me.cmbSwitchDIR)
      Me.TabPage7.Controls.Add(Me.cmdMeterCount)
      Me.TabPage7.Controls.Add(Me.cmdPlotDemandRef)
      Me.TabPage7.Controls.Add(Me.cmdLinkBld)
      resources.ApplyResources(Me.TabPage7, "TabPage7")
      Me.TabPage7.Name = "TabPage7"
      Me.LayoutEditorHelp.SetShowHelp(Me.TabPage7, CType(resources.GetObject("TabPage7.ShowHelp"), Boolean))
      Me.TabPage7.UseVisualStyleBackColor = True
      '
      'Button1
      '
      resources.ApplyResources(Me.Button1, "Button1")
      Me.Button1.Name = "Button1"
      Me.Button1.UseVisualStyleBackColor = True
      '
      'cmdGetpoint
      '
      resources.ApplyResources(Me.cmdGetpoint, "cmdGetpoint")
      Me.cmdGetpoint.Name = "cmdGetpoint"
      Me.cmdGetpoint.UseVisualStyleBackColor = True
      '
      'cdmPipeID
      '
      resources.ApplyResources(Me.cdmPipeID, "cdmPipeID")
      Me.cdmPipeID.Name = "cdmPipeID"
      Me.cdmPipeID.UseVisualStyleBackColor = True
      '
      'cmdSplitPt
      '
      resources.ApplyResources(Me.cmdSplitPt, "cmdSplitPt")
      Me.cmdSplitPt.Name = "cmdSplitPt"
      Me.cmdSplitPt.UseVisualStyleBackColor = True
      '
      'cmdSplitLink
      '
      resources.ApplyResources(Me.cmdSplitLink, "cmdSplitLink")
      Me.cmdSplitLink.Name = "cmdSplitLink"
      Me.cmdSplitLink.UseVisualStyleBackColor = True
      '
      'cmdGetLineEdge
      '
      resources.ApplyResources(Me.cmdGetLineEdge, "cmdGetLineEdge")
      Me.cmdGetLineEdge.Name = "cmdGetLineEdge"
      Me.cmdGetLineEdge.UseVisualStyleBackColor = True
      '
      'cmdNLtable
      '
      Me.cmdNLtable.BackColor = System.Drawing.Color.White
      resources.ApplyResources(Me.cmdNLtable, "cmdNLtable")
      Me.cmdNLtable.Name = "cmdNLtable"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdNLtable, CType(resources.GetObject("cmdNLtable.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdNLtable, resources.GetString("cmdNLtable.ToolTip"))
      Me.cmdNLtable.UseVisualStyleBackColor = False
      '
      'cmdClearDrawing
      '
      Me.cmdClearDrawing.BackColor = System.Drawing.Color.White
      resources.ApplyResources(Me.cmdClearDrawing, "cmdClearDrawing")
      Me.cmdClearDrawing.Name = "cmdClearDrawing"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdClearDrawing, CType(resources.GetObject("cmdClearDrawing.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdClearDrawing, resources.GetString("cmdClearDrawing.ToolTip"))
      Me.cmdClearDrawing.UseVisualStyleBackColor = False
      '
      'cmdDrawWettedArea
      '
      Me.cmdDrawWettedArea.BackColor = System.Drawing.Color.White
      resources.ApplyResources(Me.cmdDrawWettedArea, "cmdDrawWettedArea")
      Me.cmdDrawWettedArea.Name = "cmdDrawWettedArea"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdDrawWettedArea, CType(resources.GetObject("cmdDrawWettedArea.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdDrawWettedArea, resources.GetString("cmdDrawWettedArea.ToolTip"))
      Me.cmdDrawWettedArea.UseVisualStyleBackColor = False
      '
      'cmdAddJunction
      '
      Me.cmdAddJunction.BackColor = System.Drawing.Color.White
      resources.ApplyResources(Me.cmdAddJunction, "cmdAddJunction")
      Me.cmdAddJunction.Name = "cmdAddJunction"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdAddJunction, CType(resources.GetObject("cmdAddJunction.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdAddJunction, resources.GetString("cmdAddJunction.ToolTip"))
      Me.cmdAddJunction.UseVisualStyleBackColor = False
      '
      'cmdDeleteJunction
      '
      Me.cmdDeleteJunction.BackColor = System.Drawing.Color.White
      resources.ApplyResources(Me.cmdDeleteJunction, "cmdDeleteJunction")
      Me.cmdDeleteJunction.Name = "cmdDeleteJunction"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdDeleteJunction, CType(resources.GetObject("cmdDeleteJunction.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdDeleteJunction, resources.GetString("cmdDeleteJunction.ToolTip"))
      Me.cmdDeleteJunction.UseVisualStyleBackColor = False
      '
      'TabControl3
      '
      Me.TabControl3.Controls.Add(Me.TabPage8)
      Me.TabControl3.Controls.Add(Me.TabPage9)
      resources.ApplyResources(Me.TabControl3, "TabControl3")
      Me.TabControl3.Name = "TabControl3"
      Me.TabControl3.SelectedIndex = 0
      Me.LayoutEditorHelp.SetShowHelp(Me.TabControl3, CType(resources.GetObject("TabControl3.ShowHelp"), Boolean))
      '
      'TabPage8
      '
      Me.TabPage8.Controls.Add(Me.grdNodeList)
      resources.ApplyResources(Me.TabPage8, "TabPage8")
      Me.TabPage8.Name = "TabPage8"
      Me.LayoutEditorHelp.SetShowHelp(Me.TabPage8, CType(resources.GetObject("TabPage8.ShowHelp"), Boolean))
      Me.TabPage8.UseVisualStyleBackColor = True
      '
      'grdNodeList
      '
      Me.grdNodeList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
      Me.grdNodeList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column4, Me.Column5, Me.Column6, Me.Column7, Me.Column8, Me.Column15})
      resources.ApplyResources(Me.grdNodeList, "grdNodeList")
      Me.grdNodeList.Name = "grdNodeList"
      Me.LayoutEditorHelp.SetShowHelp(Me.grdNodeList, CType(resources.GetObject("grdNodeList.ShowHelp"), Boolean))
      '
      'Column4
      '
      resources.ApplyResources(Me.Column4, "Column4")
      Me.Column4.Name = "Column4"
      '
      'Column5
      '
      resources.ApplyResources(Me.Column5, "Column5")
      Me.Column5.Name = "Column5"
      '
      'Column6
      '
      resources.ApplyResources(Me.Column6, "Column6")
      Me.Column6.Name = "Column6"
      '
      'Column7
      '
      resources.ApplyResources(Me.Column7, "Column7")
      Me.Column7.Name = "Column7"
      '
      'Column8
      '
      resources.ApplyResources(Me.Column8, "Column8")
      Me.Column8.Name = "Column8"
      '
      'Column15
      '
      resources.ApplyResources(Me.Column15, "Column15")
      Me.Column15.Name = "Column15"
      '
      'TabPage9
      '
      Me.TabPage9.Controls.Add(Me.grdLinkList)
      resources.ApplyResources(Me.TabPage9, "TabPage9")
      Me.TabPage9.Name = "TabPage9"
      Me.LayoutEditorHelp.SetShowHelp(Me.TabPage9, CType(resources.GetObject("TabPage9.ShowHelp"), Boolean))
      Me.TabPage9.UseVisualStyleBackColor = True
      '
      'grdLinkList
      '
      Me.grdLinkList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
      Me.grdLinkList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column9, Me.Column10, Me.Column11, Me.Column12, Me.Column13, Me.Column14})
      resources.ApplyResources(Me.grdLinkList, "grdLinkList")
      Me.grdLinkList.Name = "grdLinkList"
      Me.LayoutEditorHelp.SetShowHelp(Me.grdLinkList, CType(resources.GetObject("grdLinkList.ShowHelp"), Boolean))
      '
      'Column9
      '
      resources.ApplyResources(Me.Column9, "Column9")
      Me.Column9.Name = "Column9"
      '
      'Column10
      '
      resources.ApplyResources(Me.Column10, "Column10")
      Me.Column10.Name = "Column10"
      '
      'Column11
      '
      resources.ApplyResources(Me.Column11, "Column11")
      Me.Column11.Name = "Column11"
      '
      'Column12
      '
      resources.ApplyResources(Me.Column12, "Column12")
      Me.Column12.Name = "Column12"
      '
      'Column13
      '
      resources.ApplyResources(Me.Column13, "Column13")
      Me.Column13.Name = "Column13"
      '
      'Column14
      '
      resources.ApplyResources(Me.Column14, "Column14")
      Me.Column14.Name = "Column14"
      '
      'cmdMeterCount
      '
      resources.ApplyResources(Me.cmdMeterCount, "cmdMeterCount")
      Me.cmdMeterCount.Name = "cmdMeterCount"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdMeterCount, CType(resources.GetObject("cmdMeterCount.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdMeterCount, resources.GetString("cmdMeterCount.ToolTip"))
      Me.cmdMeterCount.UseVisualStyleBackColor = True
      '
      'cmdPlotDemandRef
      '
      resources.ApplyResources(Me.cmdPlotDemandRef, "cmdPlotDemandRef")
      Me.cmdPlotDemandRef.Name = "cmdPlotDemandRef"
      Me.LayoutEditorHelp.SetShowHelp(Me.cmdPlotDemandRef, CType(resources.GetObject("cmdPlotDemandRef.ShowHelp"), Boolean))
      Me.ToolTip1.SetToolTip(Me.cmdPlotDemandRef, resources.GetString("cmdPlotDemandRef.ToolTip"))
      Me.cmdPlotDemandRef.UseVisualStyleBackColor = True
      '
      'TabControl1
      '
      Me.TabControl1.Controls.Add(Me.TabPage7)
      Me.TabControl1.Controls.Add(Me.TabPage1)
      Me.TabControl1.Controls.Add(Me.TabPage4)
      resources.ApplyResources(Me.TabControl1, "TabControl1")
      Me.TabControl1.Name = "TabControl1"
      Me.TabControl1.SelectedIndex = 0
      Me.LayoutEditorHelp.SetShowHelp(Me.TabControl1, CType(resources.GetObject("TabControl1.ShowHelp"), Boolean))
      '
      'PictureBox1
      '
      resources.ApplyResources(Me.PictureBox1, "PictureBox1")
      Me.PictureBox1.Name = "PictureBox1"
      Me.PictureBox1.TabStop = False
      '
      'Label5
      '
      resources.ApplyResources(Me.Label5, "Label5")
      Me.Label5.BackColor = System.Drawing.Color.Transparent
      Me.Label5.ForeColor = System.Drawing.SystemColors.ControlDarkDark
      Me.Label5.Name = "Label5"
      '
      'frmMain
      '
      resources.ApplyResources(Me, "$this")
      Me.BackColor = System.Drawing.Color.White
      Me.Controls.Add(Me.TabControl1)
      Me.Controls.Add(Me.UpdateProgress)
      Me.Controls.Add(Me.Label5)
      Me.Controls.Add(Me.PictureBox1)
      Me.Controls.Add(Me.PictureBox3)
      Me.DoubleBuffered = True
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
      Me.KeyPreview = True
      Me.Name = "frmMain"
      Me.LayoutEditorHelp.SetShowHelp(Me, CType(resources.GetObject("$this.ShowHelp"), Boolean))
      Me.ShowInTaskbar = False
      Me.TopMost = True
      CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
      Me.TabPage1.ResumeLayout(False)
      Me.GroupBox1.ResumeLayout(False)
      Me.GroupBox1.PerformLayout()
      Me.TabPage7.ResumeLayout(False)
      Me.TabControl3.ResumeLayout(False)
      Me.TabPage8.ResumeLayout(False)
      CType(Me.grdNodeList, System.ComponentModel.ISupportInitialize).EndInit()
      Me.TabPage9.ResumeLayout(False)
      CType(Me.grdLinkList, System.ComponentModel.ISupportInitialize).EndInit()
      Me.TabControl1.ResumeLayout(False)
      CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub

#End Region

   <CLSCompliant(False)> _
   Public Function Initialize(ByRef Shapefile As MapWinGIS.Shapefile, ByVal OwnerForm As Windows.Forms.Form) As Boolean
      Dim oldCursor, ownerCursor As Windows.Forms.Cursor
      oldCursor = Cursor
      ownerCursor = OwnerForm.Cursor
      Cursor = Cursors.WaitCursor
      OwnerForm.Cursor = Cursors.WaitCursor
      OwnerForm.Refresh()
      m_sf = Shapefile

      Me.Owner = OwnerForm
      Cursor = oldCursor
      OwnerForm.Cursor = ownerCursor
      m_TrulyChanged = False
      '
      ' Add Layer List to Combo
      '
      Return True
   End Function

   Private Function FindSafeFieldName(ByVal sf As MapWinGIS.Shapefile, ByVal OldName As String) As String
      Dim ht As New Hashtable()
      Dim i As Integer
      ' build a list of unique names
      For i = 0 To sf.NumFields - 1
         Dim cur As String = sf.Field(i).Name
         If ht.ContainsKey(cur) = False Then
            ht.Add(cur, cur)
         End If
      Next
      ' Start mangling the old name to come up with something new by just appending a _# to it.
      i = 1
      While ht.ContainsKey(OldName & "_" & i)
         i += 1
      End While
      Return OldName & "_" & i
   End Function

#Region "shape layer"


#End Region

   Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
      Me.Close()
   End Sub

   Private Sub frmLayoutEditor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      cmbPipe4Model.Items.Clear()
      cmbMeter4Model.Items.Clear()
      cmbBuilding.Items.Clear()
      cmbDemand_Node.Items.Clear()
      cmbZonal_Ref.Items.Clear()
      CmbValve.Items.Clear()
      cmbZone.Items.Clear()
      '
      cmbPipeField.Items.Clear()
      cmbMeterField4Bld.Items.Clear()
      cmbMeterField4Pipe.Items.Clear()
      cmbPipeField.Items.Clear()



      For i As Integer = 0 To g_MW.Layers.NumLayers - 1
         Dim LyrName As String = LCase(g_MW.Layers(g_MW.Layers.GetHandle(i)).Name)
         If g_MW.Layers(g_MW.Layers.GetHandle(i)).LayerType = MapWindow.Interfaces.eLayerType.LineShapefile Then
            cmbPipe4Model.Items.Add(LyrName)
            If InStr(LyrName, "pipe") > 0 Then cmbPipe4Model.Text = LyrName
         End If
         If g_MW.Layers(g_MW.Layers.GetHandle(i)).LayerType = MapWindow.Interfaces.eLayerType.PointShapefile Then
            '
            cmbMeter4Model.Items.Add(LyrName)
            cmbZonal_Ref.Items.Add(LyrName)
            CmbValve.Items.Add(LyrName)
            cmbDemand_Node.Items.Add(LyrName)
            If InStr(LyrName, "zone") > 0 Then cmbZonal_Ref.Text = LyrName
            If InStr(LyrName, "demand") > 0 Then cmbDemand_Node.Text = LyrName
            If InStr(LyrName, "meter") > 0 Then cmbMeter4Model.Text = LyrName
            If InStr(LyrName, "valve") > 0 Then CmbValve.Text = LyrName
         End If
         If g_MW.Layers(g_MW.Layers.GetHandle(i)).LayerType = MapWindow.Interfaces.eLayerType.PolygonShapefile Then
            cmbBuilding.Items.Add(LyrName)
            cmbZone.Items.Add(LyrName)
            If InStr(LyrName, "bldg") > 0 Then cmbBuilding.Text = LyrName
            If InStr(LyrName, "zone") > 0 Then cmbZone.Text = LyrName
         End If
      Next
      '

      'cmbBuilding.Text = cmbBuilding.Items(0)
      'cmbPipe4Model.Text = cmbPipe4Model.Items(0)
      'cmbMeter4Model.Text = cmbMeter4Model.Items(0)

      ' Adding ListView Columns
      'ListView1.Columns.Add("Emp Name", 100, HorizontalAlignment.Left)
      'ListView1.Columns.Add("Emp Address", 150, HorizontalAlignment.Left)
      'ListView1.Columns.Add("Title", 60, HorizontalAlignment.Left)
      'ListView1.Columns.Add("Salary", 50, HorizontalAlignment.Left)
      'ListView1.Columns.Add("Department", 60, HorizontalAlignment.Left)

      cmbBuildingField.Text = "BLDG_ID"
      cmbMeterField4Bld.Text = "BLDG_ID"
      cmbPipeField.Text = "PIPE_ID"
      cmbMeterField4Pipe.Text = "PIPE_ID"

   End Sub

   Function ItemID(ByVal sh As ComboBox, ByVal LryName As String) As Integer
      For i As Integer = 0 To sh.Items.Count - 1
         If Strings.InStr(UCase(sh.Items(i).ToString), UCase(LryName)) >= 1 Then
            Return i
         End If
      Next
      Return 0
   End Function


   Sub drawShape(ByVal shp As MapWinGIS.Shape, ByVal color As System.Drawing.Color)
      g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
      For i As Integer = 0 To shp.numPoints - 2
         Dim x1 As Double = shp.Point(i).x
         Dim y1 As Double = shp.Point(i).y
         Dim x2 As Double = shp.Point(i + 1).x
         Dim y2 As Double = shp.Point(i + 1).y
         g_MW.View.Draw.DrawLine(x1, y1, x2, y2, 3, color)
      Next
   End Sub

   Private Sub cmbCompareShape_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSwitchDIR.Click

      Dim shpPipe As New MapWinGIS.Shapefile
      Dim iLyr As Integer = getLayerHandle(cmbPipe4Model.Text)
      If iLyr <> -1 Then
         shpPipe = g_MW.Layers(iLyr).GetObject()
      Else
         MsgBox("Error No data")
      End If

      Dim cord() As POINTAPI
      Dim selID As Integer = g_MW.View.SelectedShapes(0).ShapeIndex
      Dim numPoint As Integer = shpPipe.Shape(selID).numPoints - 1
      ReDim cord(numPoint)
      For i As Integer = 0 To numPoint
         cord(i).X = shpPipe.Shape(selID).Point(i).x
         cord(i).Y = shpPipe.Shape(selID).Point(i).y
      Next
      shpPipe.StartEditingShapes()
      For i As Integer = 0 To numPoint
         shpPipe.Shape(selID).Point(i).x = cord(numPoint - i).X
         shpPipe.Shape(selID).Point(i).y = cord(numPoint - i).Y
      Next
      shpPipe.StopEditingShapes()
      shpPipe.Save()
   End Sub


#Region "Utility"
   Function OpenShape(ByVal Fname As String) As MapWinGIS.Shapefile
      Dim shp As New MapWinGIS.Shapefile
      Dim iLyr As Integer = getLayerHandle(Fname)
      If iLyr <> -1 Then
         'shp.Open(g_MW.Layers(iLyr).FileName)
         shp = g_MW.Layers(iLyr).GetObject()
         Return shp
      Else
         Return Nothing
      End If
   End Function

   Sub CreatFieldList(ByRef cmbShapname As ComboBox, ByRef cmbFiled As ComboBox)
      Dim shp As New MapWinGIS.Shapefile
      cmbFiled.Items.Clear()
      shp = OpenShape(cmbShapname.Text)
      If shp Is Nothing Then
         MsgBox("Select building layer before")
      Else
         For i As Integer = 0 To shp.NumFields - 1
            cmbFiled.Items.Add(shp.Field(i).Name)
         Next
         cmbFiled.Text = cmbFiled.Items(0).ToString
      End If
   End Sub

   Function PointInLine(ByVal LLn As MapWinGIS.Shapefile, ByVal x As Double, ByVal y As Double, ByVal pipeID As Integer) As MapWinGIS.Point

      Dim maxDist As Double = 10000000000
      Dim Dist As Double = 0
      Dim xx, yy As Double
      Dim ptInLine As New MapWinGIS.Point
      'Dim Ln As New MapWinGIS.Shape
      ptInLine.x = x
      ptInLine.y = y
      Dim pID As Integer = getColNum("PIPE_ID", LLn)
      If pID >= 0 Then
         Dim j As Integer = pipeID
         'For j As Integer = 0 To LLn.NumShapes - 1
         'Ln = LLn.Shape(j)
         'If j = 36 Then
         '    Dim xxxx
         '    xxxx = 1
         'End If
         Dim p1, p2 As Integer
         For p As Integer = 0 To LLn.Shape(j).NumParts - 1
            If LLn.Shape(j).NumParts = 1 Then
               p1 = 0
               p2 = LLn.Shape(j).numPoints - 1
            Else
               If p = 0 Then p1 = LLn.Shape(j).Part(p)
               If p <= LLn.Shape(j).NumParts - 2 Then
                  p2 = LLn.Shape(j).Part(p + 1) - 1
               Else
                  p2 = LLn.Shape(j).numPoints - 1
               End If
            End If
            For i As Integer = p1 + 1 To p2
               Dim x1 As Double = LLn.Shape(j).Point(i - 1).x
               Dim y1 As Double = LLn.Shape(j).Point(i - 1).y
               Dim x2 As Double = LLn.Shape(j).Point(i).x
               Dim y2 As Double = LLn.Shape(j).Point(i).y
               Dist = DistToSegment(x, y, x1, y1, x2, y2, xx, yy)
               If Dist < maxDist Then
                  ptInLine.x = xx
                  ptInLine.y = yy
                  maxDist = Dist
                  pipeID = LLn.CellValue(pID, j).ToString
               End If
            Next
            p1 = p2 + 1
         Next
         'Next
         'g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
         'g_MW.View.Draw.DrawCircle(ptInLine.x, ptInLine.y, 5, Drawing.Color.BlanchedAlmond, False)
         'g_MW.View.Draw.DrawCircle(x, y, 5, Drawing.Color.BlanchedAlmond, False)
         Return ptInLine
      Else
         Return Nothing
      End If
   End Function

   Function PointInLine(ByVal x As Double, ByVal y As Double, ByRef pipeID As Integer) As MapWinGIS.Point
      Dim maxDist As Double = 10000000000
      Dim Dist As Double = 0
      Dim xx, yy As Double
      Dim ptInLine As New MapWinGIS.Point
      'Dim Ln As New MapWinGIS.Shape
      ptInLine.x = x
      ptInLine.y = y
      Dim p1, p2 As Integer
      Dim ii As Integer = -1

      For iseg As Integer = 1 To UBound(Segment)
         Dim x1 As Double = Segment(iseg).x1
         Dim y1 As Double = Segment(iseg).y1
         Dim x2 As Double = Segment(iseg).x2
         Dim y2 As Double = Segment(iseg).y2

         Dist = DistToSegment(x, y, x1, y1, x2, y2, xx, yy)
         If Dist < maxDist Then
            ptInLine.x = xx
            ptInLine.y = yy
            maxDist = Dist
            pipeID = Segment(iseg).PipeID
            ii = iseg
            If Dist = 0 Then
               Return ptInLine
            End If
         End If
      Next
      If ii = -1 Then
         Return Nothing
      Else
         Return ptInLine
      End If

   End Function

#End Region


   Private Sub cmdLinkMeter2Buld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLinkMeter2Buld.Click
      Dim BldShp As New MapWinGIS.Shapefile
      Dim PipeShp As New MapWinGIS.Shapefile
      Dim MeterShp As New MapWinGIS.Shapefile
      BldShp = OpenShape(cmbBuilding.Text)
      PipeShp = OpenShape(cmbPipe4Model.Text)
      MeterShp = OpenShape(cmbMeter4Model.Text)

      '
      '
      Dim BldCol As Integer = getColNum(cmbBuildingField.Text, BldShp)
      Dim pipeCol As Integer = getColNum(cmbPipeField.Text, PipeShp)
      Dim metBldCol As Integer = getColNum(cmbMeterField4Bld.Text, MeterShp)
      Dim metPipeCol As Integer = getColNum(cmbMeterField4Pipe.Text, MeterShp)
      g_MW.View.Draw.ClearDrawings()
      g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)

      '
      Dim bldID As String
      Dim MeterID As String
      Dim pipeID As String
      UpdateProgress.Minimum = 0
      UpdateProgress.Maximum = BldShp.NumShapes
      UpdateProgress.Visible = True
      '        g_MW.View.LabelsEdit(0)
      For i As Integer = 0 To BldShp.NumShapes - 1

         bldID = BldShp.CellValue(BldCol, i)

         For j As Integer = 0 To MeterShp.NumShapes - 1

            MeterID = MeterShp.CellValue(metBldCol, j)
            If MeterID = bldID Then
               Dim L As Geo_Function.LineType
               L.Pt1.X = MeterShp.Shape(j).Point(0).x
               L.Pt1.Y = MeterShp.Shape(j).Point(0).y
               L.Pt2.X = BldShp.Shape(i).Centroid.x
               L.Pt2.Y = BldShp.Shape(i).Centroid.y
               If Geo_Function.Length(L) <= 300 Then
                  g_MW.View.Draw.DrawLine(MeterShp.Shape(j).Point(0).x, MeterShp.Shape(j).Point(0).y, BldShp.Shape(i).Centroid.x, BldShp.Shape(i).Centroid.y, 1, Drawing.Color.Red)
                  g_MW.View.Draw.DrawPoint(MeterShp.Shape(j).Point(0).x, MeterShp.Shape(j).Point(0).y, 5, Drawing.Color.DarkRed)
                  '
                  g_MW.View.Draw.DrawPoint(BldShp.Shape(i).Centroid.x, BldShp.Shape(i).Centroid.y, 5, Drawing.Color.DarkRed)

                  Exit For
               End If
            End If
         Next
         UpdateProgress.Value = i
      Next
      UpdateProgress.Visible = False

   End Sub

   Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
      Dim x, y As Double
      Dim d As Double = DistToSegment(0, 0, 10, 10, -10, -10, x, y)
      MsgBox(x & "," & y & " d= " & d)
   End Sub

   Private Sub cmdMoveNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveNode.Click
      'MsgBox(getLayerHandle(TextBox1.Text))
      Dim pipeLine As New MapWinGIS.Shapefile
      Dim Node As New MapWinGIS.Shapefile
      Dim Success As Boolean
      ' Get layer id
      Dim LineLayerID As Integer = getLayerHandle(cmbPipe4Model.Text)
      Dim NodeLayerID As Integer = getLayerHandle(cmbDemand_Node.Text)
      'Open Shapefile
      Success = pipeLine.Open(g_MW.Layers(LineLayerID).FileName)
      Node = g_MW.Layers(NodeLayerID).GetObject
      '
      Dim x1, y1, x2, y2, xNode, yNode, xx, yy As Double
      Dim xp, yp As Double
      UpdateProgress.Value = 1
      UpdateProgress.Minimum = 1
      UpdateProgress.Maximum = Node.NumShapes - 1
      UpdateProgress.Visible = True
      '
      '----------------------------------------------------------------------
      Dim nodeID As Integer = 1
      UpdateProgress.Visible = True
      UpdateProgress.Minimum = 1
      UpdateProgress.Maximum = Node.NumShapes
      ' Renumber Pump and Junction
      For i As Integer = 0 To Node.NumShapes - 1
         If Node.CellValue(2, i) = "PUMP" Then
            Node.StartEditingShapes()
            Success = Node.EditCellValue(0, i, 1)
            Node.StopEditingShapes()
         End If
         If Node.CellValue(2, i) = "JUNCTION" Then
            nodeID += 1
            Node.StartEditingShapes()
            Success = Node.EditCellValue(0, i, nodeID)
            Node.StopEditingShapes()
         End If
         UpdateProgress.Value = i + 1
      Next
      '
      'renumer of outlet node
      For i As Integer = 0 To Node.NumShapes - 1
         If Node.CellValue(2, i) = "OUTLET" Then
            nodeID += 1
            Node.StartEditingShapes()
            Success = Node.EditCellValue(0, i, nodeID)
            Node.StopEditingShapes()
         End If
         UpdateProgress.Value = i + 1
      Next
      UpdateProgress.Visible = True
      '----------------------------------------------------------------------
      '
      For i As Integer = 0 To Node.NumShapes - 1
         Dim MinDist As Double = 1000000000000.0
         xNode = Node.Shape(i).Point(0).x
         yNode = Node.Shape(i).Point(0).y
         xp = xNode
         yp = yNode
         For j As Integer = 0 To pipeLine.NumShapes - 1
            '
            ' OUTLET NODE
            '
            'If Node.CellValue(2, i) = "OUTLET" Then
            For k As Integer = 1 To pipeLine.Shape(j).numPoints - 1
               x1 = pipeLine.Shape(j).Point(k - 1).x
               y1 = pipeLine.Shape(j).Point(k - 1).y
               x2 = pipeLine.Shape(j).Point(k).x
               y2 = pipeLine.Shape(j).Point(k).y


               Dim L1 As Geo_Function.LineType
               L1.Pt1.X = x1 : L1.Pt1.Y = y1
               L1.Pt2.X = xNode : L1.Pt2.Y = yNode
               Dim dist1 As Double = Length(L1)
               If MinDist > dist1 And dist1 <= 100 Then
                  MinDist = dist1 : xp = x1 : yp = y1
               End If

               Dim L2 As Geo_Function.LineType
               L2.Pt1.X = xNode : L2.Pt1.Y = yNode
               L2.Pt2.X = x2 : L2.Pt2.Y = y2
               Dim dist2 As Double = Length(L2)
               If MinDist > dist2 And dist2 <= 100 Then
                  MinDist = dist2 : xp = x2 : yp = y2
               End If
               Dim dist As Double = DistancePointLine(xNode, yNode, x1, y1, x2, y2, xx, yy)

               If MinDist > dist And ((xx - xNode) ^ 2 + (yy - yNode) ^ 2) ^ 0.5 <= 10000 Then
                  MinDist = dist : xp = xx : yp = yy
               End If
            Next
            'ElseIf Node.CellValue(2, i) = "PUMP" Then
            '    '
            '    ' PUMP NODE
            '    '
            '    Dim k As Integer = pipeLine.Shape(j).numPoints - 1
            '    x1 = pipeLine.Shape(j).Point(0).x
            '    y1 = pipeLine.Shape(j).Point(0).y
            '    x2 = pipeLine.Shape(j).Point(k).x
            '    y2 = pipeLine.Shape(j).Point(k).y
            '    Dim L1 As LineType
            '    L1.Pt1.X = x1 : L1.Pt1.Y = y1
            '    L1.Pt2.X = xNode : L1.Pt2.Y = yNode
            '    Dim dist1 As Double = Length(L1)
            '    If MinDist > dist1 And dist1 <= 100 Then
            '        MinDist = dist1 : xp = x1 : yp = y1
            '    End If
            '    Dim L2 As LineType
            '    L2.Pt1.X = xNode : L2.Pt1.Y = yNode
            '    L2.Pt2.X = x2 : L2.Pt2.Y = y2
            '    Dim dist2 As Double = Length(L2)
            'End If

         Next
         'If ((xp - xNode) ^ 2 + (yp - yNode) ^ 2) ^ 0.5 >= 10000 Then
         '    Dim mmm As Integer = 1
         'End If
         'PointGIS(xp, yp, 10, Drawing.Color.Red)
         Node.StartEditingShapes()
         Node.Shape(i).Point(0).x = xp
         Node.Shape(i).Point(0).y = yp
         Node.StopEditingShapes()
         UpdateProgress.Value = i + 1
         UpdateProgress.Refresh()
      Next
      UpdateProgress.Visible = False
      'MsgBox(pipeLine.NumShapes)
      ''
      'PlotPotLink()
   End Sub


#Region "Layer combo"
   '------------------------------------------------------------
   'Field list
   '------------------------------------------------------------
   Private Sub cmbBuilding_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbBuilding.SelectedIndexChanged
      CreatFieldList(cmbBuilding, cmbBuildingField)
   End Sub

   Private Sub cmbMeter4Model_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMeter4Model.SelectedIndexChanged
      CreatFieldList(cmbMeter4Model, cmbMeterField4Bld)
      CreatFieldList(cmbMeter4Model, cmbMeterField4Pipe)
   End Sub

   Private Sub cmbPipe4Model_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMeter4Model.SelectedIndexChanged
      CreatFieldList(cmbPipe4Model, cmbPipeField)
   End Sub
#End Region


#Region "Junction tool"

   Private Sub cmdAddJunction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddJunction.Click
      Dim pipeLine As New MapWinGIS.Shapefile
      Dim Success As Boolean
      ' Get layer id
      Dim LineLayerID As Integer = getLayerHandle(cmbPipe4Model.Text)
      'Dim NodeLayerID As Integer = getLayerHandle(cmbHydrantFile.Text)
      'Open Shapefile
      Success = pipeLine.Open(g_MW.Layers(LineLayerID).FileName)
      'Success = Node.Open(g_MW.Layers(NodeLayerID).FileName)
      Dim L1 As Geo_Function.LineType
      Dim L2 As Geo_Function.LineType
      Dim nIntersec As Integer = 0

      Dim maxPoint As Integer = 0
      For i As Integer = 0 To pipeLine.NumShapes - 1
         If maxPoint < pipeLine.Shape(i).numPoints Then
            maxPoint = pipeLine.Shape(i).numPoints - 1
         End If
      Next
      Dim L(pipeLine.NumShapes - 1, maxPoint) As POINTAPI
      For i As Integer = 0 To pipeLine.NumShapes - 1
         For i1 As Integer = 0 To pipeLine.Shape(i).numPoints - 1
            L(i, i1).X = pipeLine.Shape(i).Point(i1).x
            L(i, i1).Y = pipeLine.Shape(i).Point(i1).y
         Next
      Next
      Dim jPoint() As POINTAPI
      ReDim jPoint(1)
      For i As Integer = 0 To pipeLine.NumShapes - 1
         For i1 As Integer = 1 To pipeLine.Shape(i).numPoints - 1
            L1.Pt1.X = L(i, i1 - 1).X
            L1.Pt1.Y = L(i, i1 - 1).Y
            L1.Pt2.X = L(i, i1).X
            L1.Pt2.Y = L(i, i1).Y

            LineGIS(L1.Pt1.X, L1.Pt1.Y, L1.Pt2.X, L1.Pt2.Y, 1, Drawing.Color.Gold)

            For j As Integer = i + 1 To pipeLine.NumShapes - 1
               If i < j Then
                  For j1 As Integer = 1 To pipeLine.Shape(j).numPoints - 1
                     L2.Pt1.X = L(j, j1 - 1).X
                     L2.Pt1.Y = L(j, j1 - 1).Y
                     L2.Pt2.X = L(j, j1).X
                     L2.Pt2.Y = L(j, j1).Y
                     'If Geo_Function.lineIntersection(L1, L2, pt) = True Then
                     '    PointGIS(pt.X, pt.Y, 10, Drawing.Color.LightYellow)
                     '    nIntersec += 1
                     'End If
                     Dim pt As POINTAPI
                     If Geo_Function.SegmentsIntersect(L1, L2, pt) = True Then
                        'Debug.Print(pt.X & "," & pt.Y)

                        Dim chk As Boolean = True
                        For k As Integer = 1 To nIntersec
                           If jPoint(k).X = pt.X And jPoint(k).Y = pt.Y Then
                              chk = False
                           End If
                        Next
                        If chk = True Then
                           nIntersec += 1
                           ReDim Preserve jPoint(nIntersec)
                           jPoint(nIntersec) = pt
                           CreateSp(pt.X, pt.Y, "JUNCTION")
                           PointGIS(pt.X, pt.Y, 5, Drawing.Color.Red)
                        End If
                     End If
                  Next
               End If
            Next
         Next
      Next
      MsgBox("number of intersection point:" & nIntersec)

   End Sub

   Private Sub cmdDeleteJunction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteJunction.Click
      Dim Node As New MapWinGIS.Shapefile
      Dim Success As Boolean
      ' Get layer id
      Dim LineLayerID As Integer = getLayerHandle(cmbPipe4Model.Text)
      Dim NodeLayerID As Integer = getLayerHandle(cmbDemand_Node.Text)
      'Open Shapefile
      Node = g_MW.Layers(NodeLayerID).GetObject
      '

      UpdateProgress.Value = 1
      UpdateProgress.Minimum = 1
      UpdateProgress.Maximum = Node.NumShapes - 1
      UpdateProgress.Visible = True
      '
      '----------------------------------------------------------------------
      Dim nodeID As Integer = 1
      UpdateProgress.Visible = True
      UpdateProgress.Minimum = 1
      UpdateProgress.Maximum = Node.NumShapes
      ' Renumber Pump and Junction
      Node.StartEditingShapes()
      For i As Integer = Node.NumShapes - 1 To 0 Step -1
         If Node.CellValue(2, i) = "JUNCTION" Then
            Success = Node.EditDeleteShape(i)
         End If
         UpdateProgress.Value = i + 1
      Next
      UpdateProgress.Visible = False
      'Node.Close()
      Node.StopEditingShapes()
      Node = Nothing

   End Sub

#End Region

   Private Sub cmdPlotDemandRef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPlotDemandRef.Click
      g_MW.View.Draw.ClearDrawings()
      g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
      Dim Node As New MapWinGIS.Shapefile
      Dim Zone As New MapWinGIS.Shapefile
      ' Get layer id
      Dim zoneLayerID As Integer = getLayerHandle(cmbZonal_Ref.Text)
      Dim NodeLayerID As Integer = getLayerHandle(cmbMeter4Model.Text)
      'Open Shapefile
      Node = g_MW.Layers(NodeLayerID).GetObject
      Zone = g_MW.Layers(zoneLayerID).GetObject

      'Zone in Node reference
      Dim ZoneCol As Integer = getColNum("Zone", Node)
      If ZoneCol = -1 Then
         MsgBox("Don't have filed zone in Naodal shape")
         Exit Sub
      End If
      'Zone in Zone reference
      Dim ZoneRefCol As Integer = getColNum("Zone", Zone)
      Dim countRefCol As Integer = getColNum("NumMeter", Zone)
      If ZoneRefCol = -1 Then
         MsgBox("Don't have filed zone in Zone reference shape")
         Exit Sub
      End If

      UpdateProgress.Maximum = Zone.NumShapes - 1
      UpdateProgress.Minimum = 0
      UpdateProgress.Visible = True

      For i As Integer = 0 To Zone.NumShapes - 1
         UpdateProgress.Value = i
         UpdateProgress.Refresh()

         Dim NumMeter As Integer = 0
         Dim zoneNo As Integer = CInt(Zone.CellValue(ZoneRefCol, i))
         Dim x1 As Double = Zone.Shape(i).Point(0).x
         Dim y1 As Double = Zone.Shape(i).Point(0).y
         For j As Integer = 0 To Node.NumShapes - 1
            Dim NodalzoneNo As Integer = CInt(Node.CellValue(ZoneCol, j))
            If zoneNo = NodalzoneNo Then
               Dim x2 As Double = Node.Shape(j).Point(0).x
               Dim y2 As Double = Node.Shape(j).Point(0).y
               LineGIS(x1, y1, x2, y2, 1, Drawing.Color.Red)
               NumMeter += 1
            End If
         Next
         Zone.StartEditingTable()
         Zone.EditCellValue(countRefCol, i, NumMeter)
         Zone.StopEditingTable()
      Next
      Zone.Save()
      g_MW.View.Redraw()
      Node = Nothing
      Zone = Nothing
      UpdateProgress.Visible = False
   End Sub

   Private Sub cmdClearDrawing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearDrawing.Click
      g_MW.View.Draw.ClearDrawings()
      g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
   End Sub

   Private Sub cmdMeterCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMeterCount.Click
      Dim Node As New MapWinGIS.Shapefile
      Dim Zone As New MapWinGIS.Shapefile
      ' Get layer id
      Dim zoneLayerID As Integer = getLayerHandle(cmbZone.Text)
      Dim NodeLayerID As Integer = getLayerHandle(cmbMeter4Model.Text)
      'Open Shapefile
      Node = g_MW.Layers(NodeLayerID).GetObject
      Zone = g_MW.Layers(zoneLayerID).GetObject

      'Zone in Node reference
      Dim ZoneCol As Integer = getColNum("Zone", Node)
      If ZoneCol = -1 Then
         MsgBox("Don't have filed zone in Naodal shape")
         Exit Sub
      End If
      'Zone in Zone reference
      Dim ZoneRefCol As Integer = getColNum("Zone", Zone)
      Dim countRefCol As Integer = getColNum("NumMeter", Zone)
      If ZoneRefCol = -1 Then
         MsgBox("Don't have filed zone in Block/Zone shape")
         Exit Sub
      End If
      UpdateProgress.Maximum = Zone.NumShapes - 1
      UpdateProgress.Minimum = 0
      UpdateProgress.Visible = True
      For i As Integer = 0 To Zone.NumShapes - 1
         UpdateProgress.Value = i
         UpdateProgress.Refresh()
         Dim NumMeter As Integer = 0
         Dim PolyZone() As Geo_Function.POINTAPI
         ReDim PolyZone(Zone.Shape(i).numPoints - 1)
         For j As Integer = 0 To Zone.Shape(i).numPoints - 1
            PolyZone(j).X = Zone.Shape(i).Point(j).x
            PolyZone(j).Y = Zone.Shape(i).Point(j).y
         Next
         For j As Integer = 0 To Node.NumShapes - 1
            Dim x2 As Double = Node.Shape(j).Point(0).x
            Dim y2 As Double = Node.Shape(j).Point(0).y
            If PointInPolygon(PolyZone, x2, y2) Then
               NumMeter += 1
               PointGIS(x2, y2, 1, Drawing.Color.Magenta)
            End If
            'If Zone.Shape(i).PointInThisPoly(Node.Shape(j).Point(0)) = True Then
            'End If
         Next
         Zone.StartEditingTable()
         Zone.EditCellValue(countRefCol, i, NumMeter)
         Zone.StopEditingTable()
      Next
      Zone.Save()
      Node = Nothing
      Zone = Nothing
      UpdateProgress.Visible = False
   End Sub

  
   Private Sub cmdGetLineEdge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetLineEdge.Click
      Dim T1 As String = Now
      Dim HashLink As New ArrayList
      grdNodeList.Rows.Clear()
      grdLinkList.Rows.Clear()
      Dim shpPipe As New MapWinGIS.Shapefile
      Dim iLyr As Integer = getLayerHandle(cmbPipe4Model.Text)
      If iLyr <> -1 Then
         shpPipe = g_MW.Layers(iLyr).GetObject()
         shpPipe.CreateSpatialIndex(g_MW.Layers(iLyr).FileName)
         shpPipe.UseQTree = False
      Else
         MsgBox("Error No pipe data")
         Exit Sub
      End If

      'HashLink.Clear()
      Dim ii As Integer = 0
      UpdateProgress.Maximum = shpPipe.NumShapes
      UpdateProgress.Visible = True
      Dim nodeList As New ArrayList
      'nodeList.Clear()

      'Dim PIPE_ID As Integer = getColNum("PIPE_ID", shpPipe)
      'If PIPE_ID = -1 Then
      '    MsgBox("NO PIPE_ID Field")
      '    Exit Sub
      'End If


      ReDim Link(shpPipe.NumShapes)
      For i As Integer = 0 To shpPipe.NumShapes - 1
         Dim lastNode As Integer = shpPipe.Shape(i).numPoints - 1
         Dim x1, y1, x2, y2 As Double
         x1 = Math.Round(shpPipe.Shape(i).Point(0).x, 2)
         y1 = Math.Round(shpPipe.Shape(i).Point(0).y, 2)
         x2 = Math.Round(shpPipe.Shape(i).Point(lastNode).x, 2)
         y2 = Math.Round(shpPipe.Shape(i).Point(lastNode).y, 2)

         'US node
         Try
            If dupplicateChk(x1 & "," & y1, nodeList) = False Then ' HashNode.Add(x1 & "," & y1, ii)
               nodeList.Add(x1 & "," & y1)
            End If
         Catch ex As Exception
         End Try

         'DS node
         Try
            If dupplicateChk(x2 & "," & y2, nodeList) = False Then ' HashNode.Add(x1 & "," & y1, ii)
               nodeList.Add(x2 & "," & y2)
            End If
         Catch ex As Exception
         End Try

         Dim N1 As Integer = Array.IndexOf(nodeList.ToArray, x1 & "," & y1)  'getNode(x1 & "," & y1)
         Dim N2 As Integer = Array.IndexOf(nodeList.ToArray, x2 & "," & y2)  'getNode(x2 & "," & y2)
         'Link(i).id = shpPipe.CellValue(PIPE_ID, i)
         Link(i).NodeF = N1
         Link(i).NodeT = N2
         Link(i).L = shpPipe.Shape(i).Length
         UpdateProgress.Value = i
      Next
      'MsgBox("Number of node is " & HatchNode.Count)
      UpdateProgress.Maximum = nodeList.Count
      ReDim Node(nodeList.Count)
      '
      '
      g_MW.View.Draw.ClearDrawings()
      g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
      '
      For i As Integer = 0 To nodeList.Count - 1
         Dim st As String = nodeList.Item(i)
         If Not st Is Nothing Then
            grdNodeList.Rows.Add()
            Dim a As Object = Split(st, ",")
            grdNodeList.Item(0, i).Value = i
            grdNodeList.Item(1, i).Value = a(0)
            grdNodeList.Item(2, i).Value = a(1)
            Node(i).x = a(0)
            Node(i).y = a(1)
            UpdateProgress.Value = i
            g_MW.View.Draw.DrawPoint(Node(i).x, Node(i).y, 3, Drawing.Color.Red)
         End If
      Next

      UpdateProgress.Maximum = shpPipe.NumShapes
      For i As Integer = 0 To shpPipe.NumShapes - 1
         grdLinkList.Rows.Add()
         grdLinkList.Item(0, i).Value = Link(i).id
         grdLinkList.Item(1, i).Value = Link(i).NodeF
         grdLinkList.Item(2, i).Value = Link(i).NodeT
         grdLinkList.Item(3, i).Value = Math.Round(Link(i).L, 2)
         UpdateProgress.Value = i
         g_MW.View.Draw.DrawLine(Node(Link(i).NodeF).x, Node(Link(i).NodeF).y, Node(Link(i).NodeT).x, Node(Link(i).NodeT).y, 1, Drawing.Color.Red)
         Node(Link(i).NodeF).UsedCount += 1
         Node(Link(i).NodeT).UsedCount += 1
      Next
      g_MW.View.Redraw()
      For i As Integer = 0 To nodeList.Count - 1
         'Class junction order
         grdNodeList.Item(3, i).Value = Node(i).UsedCount
         If Node(i).UsedCount = 1 Then g_MW.View.Draw.DrawPoint(Node(i).x, Node(i).y, 4, Drawing.Color.Magenta)
         If Node(i).UsedCount = 2 Then g_MW.View.Draw.DrawPoint(Node(i).x, Node(i).y, 2, Drawing.Color.Green)
         If Node(i).UsedCount = 3 Then g_MW.View.Draw.DrawPoint(Node(i).x, Node(i).y, 3, Drawing.Color.Blue)
         If Node(i).UsedCount = 4 Then g_MW.View.Draw.DrawPoint(Node(i).x, Node(i).y, 4, Drawing.Color.Black)
         If Node(i).UsedCount > 4 Then
            g_MW.View.Draw.DrawPoint(Node(i).x, Node(i).y, 5, Drawing.Color.Red)
            Debug.Print(i & " no. is " & Node(i).UsedCount)
         End If
      Next
      g_MW.View.Redraw()
      UpdateProgress.Visible = False
      MsgBox(T1 & " " & Now)
   End Sub

   Function dupplicateChk(ByVal chk As String, ByVal Ar As ArrayList) As Boolean
      dupplicateChk = False
      For i As Integer = 0 To Ar.Count - 1
         If chk = Ar.Item(i).ToString Then
            dupplicateChk = True
            Exit Function
         End If
      Next
   End Function

   Private Sub cmdSplitLink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSplitLink.Click
      Dim xx, yy As Double
      Dim T1 As String = Now
      Dim fx As Integer = FreeFile()
      'FileOpen(fx, "F:\SEATEC_Project\Web-GIS\2012-08-25\mwNetworkGenPlugin\Node_insert.txt", OpenMode.Output)
      FileOpen(fx, "F:\temp\Node_insert3.txt", OpenMode.Output)
      Dim HashLink As New ArrayList
      grdNodeList.Rows.Clear()
      grdLinkList.Rows.Clear()
      Dim shpPipe As New MapWinGIS.Shapefile
      Dim iLyr As Integer = getLayerHandle(cmbPipe4Model.Text)
      If iLyr <> -1 Then
         shpPipe = g_MW.Layers(iLyr).GetObject()
      Else
         MsgBox("Error No pipe data")
         FileClose(fx)
         Exit Sub
      End If

      'HashLink.Clear()
      Dim ii As Integer = 0
      UpdateProgress.Maximum = shpPipe.NumShapes
      UpdateProgress.Visible = True
      Dim nodeList As New ArrayList
      'nodeList.Clear()

      Dim PIPE_ID As Integer = getColNum("PIPE_ID", shpPipe)
      If PIPE_ID = -1 Then
         MsgBox("The selected layer don't hvae myPIPE_ID Field")
         FileClose(fx)
         Exit Sub
      End If
      '
      g_MW.View.Draw.ClearDrawings()
      Dim hdl As Integer = g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)

      UpdateProgress.Maximum = shpPipe.NumShapes
      UpdateProgress.Visible = True

      ReDim Link(shpPipe.NumShapes)

      For i As Integer = 0 To shpPipe.NumShapes - 1
         'Line segment
         Dim ShapeID() As Int32
         shpPipe.SelectShapes(shpPipe.Shape(i).Extents, 0, MapWinGIS.SelectMode.INTERSECTION, ShapeID)
         'First point/Last point
         Dim x1Actv, y1Actv, x2Actv, y2Actv As Double
         x1Actv = shpPipe.Shape(i).Point(0).x
         y1Actv = shpPipe.Shape(i).Point(0).y
         x2Actv = shpPipe.Shape(i).Point(shpPipe.Shape(i).numPoints - 1).x
         y2Actv = shpPipe.Shape(i).Point(shpPipe.Shape(i).numPoints - 1).y

         If shpPipe.Shape(i).NumParts > 1 Then
            Debug.Print("Pipe id " & i & " have " & shpPipe.Shape(i).NumParts & " number paths")
         End If

         Dim Dist As New ArrayList()

         For seg As Integer = 1 To shpPipe.Shape(i).numPoints - 1
            Dim segX1, segY1, segX2, segY2 As Double
            'Dim ext As New MapWinGIS.Extents
            segX1 = shpPipe.Shape(i).Point(seg - 1).x
            segY1 = shpPipe.Shape(i).Point(seg - 1).y
            segX2 = shpPipe.Shape(i).Point(seg).x
            segY2 = shpPipe.Shape(i).Point(seg).y

            'If segX1 > segX2 Then
            '    Dim xTmp As Double = segX1
            '    segX1 = segX2
            '    segX2 = xTmp
            'End If
            'If segY1 > segY2 Then
            '    Dim yTmp As Double = segY1
            '    segY1 = segY2
            '    segY2 = yTmp
            'End If

            'ext.SetBounds(segX1, segY1, 0, segX2, segY2, 0)

            If UBound(ShapeID) > 0 Then
               Dim LineSegment As Geo_Function.LineType
               LineSegment.Pt1.X = segX1
               LineSegment.Pt1.Y = segY1
               LineSegment.Pt2.X = segX2
               LineSegment.Pt2.Y = segY2

               For jj As Integer = LBound(ShapeID) To UBound(ShapeID)
                  Dim j As Integer = ShapeID(jj)
                  If i <> j Then

                     Dim lastNode As Integer = shpPipe.Shape(j).numPoints - 1
                     Dim x1, y1, x2, y2 As Double
                     x1 = shpPipe.Shape(j).Point(0).x
                     y1 = shpPipe.Shape(j).Point(0).y
                     If ((x1Actv - x1) ^ 2 + (y1Actv - y1) ^ 2) ^ 0.5 > 0.5 Then
                        If ((x2Actv - x1) ^ 2 + (y2Actv - y1) ^ 2) ^ 0.5 > 0.5 Then
                           'If Point_Line(x1, y1, LineSegment) = True Then
                           If Geo_Function.DistToSegment(x1, y1, LineSegment, xx, yy) < 1 Then
                              g_MW.View.Draw.DrawPoint(x1, y1, 5, Drawing.Color.Blue)
                              'PrintLine(fx, "Pipe ID " & vbTab & shpPipe.CellValue(PIPE_ID, i) & vbTab & "Segment ID " & seg & vbTab & "Point " & vbTab & x1 & "," & y1)
                              Dim L As Double = ShapeLength(shpPipe.Shape(i), seg, xx, yy)
                              'g_MW.View.Draw.AddDrawingLabel(hdl, L.ToString, Drawing.Color.Black, xx, yy, MapWinGIS.tkHJustification.hjLeft)
                              Dist.Add(Format(L, "000000000.000") & ", " & seg & ", " & xx & ", " & yy)

                              'If j = 183 Or i = 183 Then
                              '    Debug.Print("Actv-1 :" & ((x1Actv - x1) ^ 2 + (y1Actv - y1) ^ 2) ^ 0.5)
                              '    Debug.Print("Actv-2 :" & ((x2Actv - x1) ^ 2 + (y2Actv - y1) ^ 2) ^ 0.5)
                              '    Dim mmm
                              '    mmm = 0
                              'End If
                           End If
                        End If
                     End If

                     x2 = shpPipe.Shape(j).Point(lastNode).x
                     y2 = shpPipe.Shape(j).Point(lastNode).y
                     If ((x1Actv - x2) ^ 2 + (y1Actv - y2) ^ 2) ^ 0.5 > 0.5 Then
                        If ((x2Actv - x2) ^ 2 + (y2Actv - y2) ^ 2) ^ 0.5 > 0.5 Then
                           'If Point_Line(x2, y2, LineSegment) = True Then
                           If Geo_Function.DistToSegment(x2, y2, LineSegment, xx, yy) < 1 Then
                              g_MW.View.Draw.DrawPoint(x2, y2, 5, Drawing.Color.Blue)
                              'PrintLine(fx, "Pipe ID " & vbTab & shpPipe.CellValue(PIPE_ID, i) & vbTab & "Segment ID " & seg & vbTab & "Point " & vbTab & x2 & "," & y2)
                              Dim L As Double = ShapeLength(shpPipe.Shape(i), seg, xx, yy)
                              'g_MW.View.Draw.AddDrawingLabel(hdl, L.ToString, Drawing.Color.Black, xx, yy, MapWinGIS.tkHJustification.hjLeft)
                              Dist.Add(Format(L, "000000000.000") & ", " & seg & ", " & xx & ", " & yy)
                              'If j = 183 And i = 184 Then
                              '    Debug.Print("Actv-1 :" & ((x1Actv - x2) ^ 2 + (y1Actv - y2) ^ 2) ^ 0.5)
                              '    Debug.Print("Actv-2 :" & ((x2Actv - x2) ^ 2 + (y2Actv - y2) ^ 2) ^ 0.5)
                              '    Dim mmm
                              '    mmm = 0
                              '    g_MW.View.Draw.DrawLine(LineSegment.Pt1.X, LineSegment.Pt1.Y, LineSegment.Pt2.X, LineSegment.Pt2.Y, 1, Drawing.Color.Magenta)
                              '    g_MW.View.Draw.DrawPoint(x2, y2, 8, Drawing.Color.Red)
                              'End If
                              'shpPipe.Shape(i).InsertPoint   
                           End If
                        End If
                     End If
                  End If
               Next
            End If

         Next 'End of Segment

         UpdateProgress.Value = i
         Dist.Sort()
         Dim Seg1 As Integer = 0
         Dim xx0, yy0, xx1, yy1 As Double
         xx0 = shpPipe.Shape(i).Point(0).x
         yy0 = shpPipe.Shape(i).Point(0).y
         For ind As Integer = 0 To Dist.Count - 1
            Dim a As Object = Split(Dist(ind), ",")
            Dim Seg2 As Integer = a(1)
            xx = a(2)
            yy = a(3)
            g_MW.View.Draw.AddDrawingLabel(hdl, "(" & ind & ") " & CDbl(a(0)).ToString, Drawing.Color.Black, xx, yy, MapWinGIS.tkHJustification.hjLeft)
            PrintLine(fx, "ADD_LINE")
            PrintLine(fx, xx0 & "," & yy0)
            For jnd As Integer = Seg1 + 1 To Seg2 - 1
               xx1 = shpPipe.Shape(i).Point(jnd).x
               yy1 = shpPipe.Shape(i).Point(jnd).y
               PrintLine(fx, xx1 & "," & yy1)
               xx0 = xx1
               yy0 = yy1
            Next
            PrintLine(fx, xx & "," & yy)
            xx0 = xx
            yy0 = yy
            Seg1 = Seg2 - 1
            PrintLine(fx, "END_LINE")
            PrintLine(fx, "ADD_ATTRIBUTE, " & i)
         Next
         Dim LastPt As Integer = shpPipe.Shape(i).numPoints - 1
         If xx0 <> shpPipe.Shape(i).Point(LastPt).x And yy0 <> shpPipe.Shape(i).Point(LastPt).y And Dist.Count > 0 Then
            PrintLine(fx, "ADD_LINE")
            PrintLine(fx, xx0 & "," & yy0)
            For jnd As Integer = Seg1 + 1 To LastPt
               xx1 = shpPipe.Shape(i).Point(jnd).x
               yy1 = shpPipe.Shape(i).Point(jnd).y
               PrintLine(fx, xx1 & "," & yy1)
               xx0 = xx1
               yy0 = yy1
            Next
            PrintLine(fx, "END_LINE")
            PrintLine(fx, "ADD_ATTRIBUTE, " & i)
            'For jj As Integer = 0 To shpPipe.NumFields - 1
            '    PrintLine(fx, shpPipe.CellValue(jj, i))
            'Next
            'FileClose(fx)
         End If
         If Dist.Count > 0 Then
            PrintLine(fx, "DELETE_LINE, " & i)
         End If
      Next
      UpdateProgress.Visible = False
      FileClose(fx)
      FileOpen(fx, "F:\temp\Node_insert3.txt", OpenMode.Input)
      Dim nSeg As Integer = 0
      Dim Cord() As POINTAPI
      Dim adCord As Boolean = False
      While (Not EOF(fx))
         Dim st As String = LineInput(fx)
         If st = "END_LINE" Then
            adCord = False
            For i As Integer = 1 To nSeg - 1
               g_MW.View.Draw.DrawLine(Cord(i).X, Cord(i).Y, Cord(i + 1).X, Cord(i + 1).Y, 2, Drawing.Color.Magenta)
            Next
            'Draw node 1
            g_MW.View.Draw.DrawPoint(Cord(1).X, Cord(1).Y, 5, Drawing.Color.Yellow)
            'Draw node 2
            g_MW.View.Draw.DrawPoint(Cord(nSeg).X, Cord(nSeg).Y, 5, Drawing.Color.Yellow)
            nSeg = 0
         End If

         If adCord = True Then
            nSeg += 1
            ReDim Preserve Cord(nSeg)
            Dim a As Object = Split(st, ",")
            Cord(nSeg).X = a(0)
            Cord(nSeg).Y = a(1)
         End If
         If st = "ADD_LINE" Then adCord = True
      End While
      FileClose(fx)
      MsgBox(T1 & " " & Now)
   End Sub

   Function ShapeLength(ByVal shp As MapWinGIS.Shape, ByVal seg As Integer, ByVal x As Double, ByVal y As Double) As Double
      Dim Lseg As Double = 0
      Dim L As Double = 0
      Dim x1, y1, x2, y2 As Double
      x1 = shp.Point(0).x
      y1 = shp.Point(0).y
      If seg > 1 Then
         For i As Integer = 1 To seg - 1
            x1 = shp.Point(i - 1).x
            y1 = shp.Point(i - 1).y
            x2 = shp.Point(i).x
            y2 = shp.Point(i).y
            Lseg = ((x1 - x2) ^ 2 + (y1 - y2) ^ 2) ^ 0.5
            L += Lseg
         Next
         x1 = shp.Point(seg - 1).x
         y1 = shp.Point(seg - 1).y
      End If
      x2 = x
      y2 = y
      Lseg = ((x1 - x2) ^ 2 + (y1 - y2) ^ 2) ^ 0.5
      L += Lseg
      Return Math.Round(L / 1000, 2)
   End Function

   Function SplitShape(ByVal shp As MapWinGIS.Shape, ByVal seg As Integer, ByVal x As Double, ByVal y As Double) As Double
      Dim Lseg As Double = 0
      Dim L As Double = 0
      Dim x1, y1, x2, y2 As Double
      x1 = shp.Point(0).x
      y1 = shp.Point(0).y
      If seg > 1 Then
         For i As Integer = 1 To seg - 1
            x1 = shp.Point(i - 1).x
            y1 = shp.Point(i - 1).y
            x2 = shp.Point(i).x
            y2 = shp.Point(i).y
            Lseg = ((x1 - x2) ^ 2 + (y1 - y2) ^ 2) ^ 0.5
            L += Lseg
         Next
         x1 = shp.Point(seg - 1).x
         y1 = shp.Point(seg - 1).y
      End If
      x2 = x
      y2 = y
      Lseg = ((x1 - x2) ^ 2 + (y1 - y2) ^ 2) ^ 0.5
      L += Lseg
      Return Math.Round(L / 1000, 2)
   End Function



   Private Sub cmdSplitPt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSplitPt.Click
      Dim fx As Integer = FreeFile()

      'FileOpen(fx, "F:\SEATEC_Project\Web-GIS\2012-08-25\mwNetworkGenPlugin\Node_insert-Backup.txt", OpenMode.Input)
      FileOpen(fx, "F:\temp\Node_insert3.txt", OpenMode.Input)
      Dim fx2 As Integer = FreeFile()
      FileOpen(fx2, "F:\temp\Pipe4Delete.txt", OpenMode.Output)
      '
      Dim shpPipe As New MapWinGIS.Shapefile
      Dim iLyr As Integer = getLayerHandle(cmbPipe4Model.Text)
      If iLyr <> -1 Then
         shpPipe = g_MW.Layers(iLyr).GetObject()
      Else
         MsgBox("Error No pipe data")
         Exit Sub
      End If
      g_MW.View.Draw.ClearDrawings()
      g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
      Dim nSeg As Integer = 0
      Dim CordX() As Double
      Dim CordY() As Double
      Dim adCord As Boolean = False
      Dim pt As MapWinGIS.Point
      shpPipe.StartEditingShapes()
      shpPipe.UseQTree = True
      While (Not EOF(fx))
         Dim st As String = LineInput(fx)
         If st = "END_LINE" Then
            adCord = False
         End If
         Dim aa As Object = Split(st, ",")
         If aa(0) = "DELETE_LINE" Then
            PrintLine(fx2, st)
         End If
         If adCord = True Then
            nSeg += 1
            ReDim Preserve CordX(nSeg)
            ReDim Preserve CordY(nSeg)
            Dim a As Object = Split(st, ",")
            CordX(nSeg) = a(0)
            CordY(nSeg) = a(1)
         End If
         If st = "ADD_LINE" Then adCord = True
         If InStr(st, "ADD_ATTRIBUTE") > 0 Then
            Dim b As Object = Split(st, ",")
            Dim id As Integer = CInt(b(1))


            Dim success As Boolean
            Dim shapeIndex As Long
            'objects
            Dim shape As MapWinGIS.Shape
            Dim point As MapWinGIS.Point

            'Start Editing it...
            success = shpPipe.StartEditingShapes(True)
            shapeIndex = shpPipe.NumShapes
            point = New MapWinGIS.Point
            Dim npt As Integer = UBound(CordX)

            'Create and add shape
            shape = New MapWinGIS.Shape

            'Create a new Point shape
            With shape
               success = .Create(MapWinGIS.ShpfileType.SHP_POLYLINE)
               If Not success Then
                  MsgBox("Error in creating shape: " & .ErrorMsg(.LastErrorCode))
               End If
            End With ' shape

            'Set the values for the point to be inserted

            'shape = New MapWinGIS.Shape
            shape.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
            'Set up the points

            'Insert the point into the shape

            For i As Integer = 0 To npt - 1
               point = New MapWinGIS.Point
               point.x = CordX(i + 1)
               point.y = CordY(i + 1)

               'Add the points to a shape
               success = shape.InsertPoint(point, i)
               If Not success Then
                  MsgBox("Error in adding point: " & shape.ErrorMsg(shape.LastErrorCode))
               End If
            Next i

            'Insert the shape into the shapefile
            With shpPipe
               success = .EditInsertShape(shape, shapeIndex)
               If Not success Then
                  MsgBox("Error in adding field: " & .ErrorMsg(.LastErrorCode))
               End If
            End With ' shpPipe

            'Add value to at least one attribute
            'Use shapeindex as dummy value:

            With shpPipe
               For i As Integer = 0 To shpPipe.NumFields - 1
                  success = .EditCellValue(i, shapeIndex, shpPipe.CellValue(i, id))
               Next

               If Not success Then
                  MsgBox("Error in adding field: " & .ErrorMsg(.LastErrorCode))
               End If
            End With ' sf
            'success = sf.Save

            'Stop editing shapes in the shapefile, saving changes to shapes,
            'also stopping editing of the attribute table
            'With shpPipe
            '    success = shpPipe.StopEditingShapes(True, True)
            '    If Not success Then
            '        MsgBox("Error in adding field: " & .ErrorMsg(.LastErrorCode))
            '    End If
            'End With ' sf
Cleanup:
            On Error Resume Next
            shape = Nothing
            point = Nothing
            'newPipeShp(CordX, CordY, id, shpPipe)
            For i As Integer = 1 To nSeg - 1
               g_MW.View.Draw.DrawLine(CordX(i), CordY(i), CordX(i + 1), CordY(i + 1), 2, Drawing.Color.Magenta)
            Next
            'Draw node 1
            g_MW.View.Draw.DrawPoint(CordX(1), CordY(1), 5, Drawing.Color.Yellow)
            'Draw node 2
            g_MW.View.Draw.DrawPoint(CordX(nSeg), CordY(nSeg), 5, Drawing.Color.Yellow)
            nSeg = 0
         End If
      End While
      With shpPipe
         Dim success As Boolean = shpPipe.StopEditingShapes(True, True)
         If Not success Then
            MsgBox("Error in adding field: " & .ErrorMsg(.LastErrorCode))
         End If
      End With ' sf        FileClose(fx2)
      FileClose(fx)
      '--------------------------------------------------------------------------
      'DELETE PIPE
      '
      '--------------------------------------------------------------------------
      FileOpen(fx2, "F:\temp\Pipe4Delete.txt", OpenMode.Input)
      '
      shpPipe.StartEditingShapes()
      shpPipe.UseQTree = True
      While (Not EOF(fx2))
         Dim st As String = LineInput(fx2)
         Dim a As Object = Split(st, ",")
         If a(0) = "DELETE_LINE" Then
            Dim id As Integer = CInt(a(1))
            shpPipe.EditDeleteShape(id)
         End If
      End While

      shpPipe.StopEditingShapes()
      shpPipe.Save()

      FileClose(fx2)
   End Sub

   Function GetShapeID(ByVal shp As MapWinGIS.Shapefile, ByVal FieldID As Integer, ByVal Search As String, ByRef numFound As Integer) As Integer
      Dim nFound As Integer = 0
      For i As Integer = 0 To shp.NumShapes - 1
         If shp.CellValue(FieldID, i).ToString = Search Then
            nFound += 1
            GetShapeID = Search
         End If
      Next
      numFound = nFound
   End Function


   Private Sub cdmPipeID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cdmPipeID.Click
      Dim xx As MapWinGeoProc.Topology2D.LineSegment
      Dim shpPipe As New MapWinGIS.Shapefile

      Dim iLyr As Integer = getLayerHandle(cmbPipe4Model.Text)

      If iLyr <> -1 Then
         shpPipe = g_MW.Layers(iLyr).GetObject()
      Else
         MsgBox("Error No pipe data")
         Exit Sub
      End If

      Dim myPIPE_ID As Integer = getColNum("myPIPE_ID", shpPipe)
      If myPIPE_ID = -1 Then
         MsgBox("Error No myPIPE_ID Filed in the selected layer")
         Exit Sub
      End If

      'shpPipe.StartEditingTable()
      'For i As Integer = 0 To 1 'shpPipe.NumShapes
      '    shpPipe.EditCellValue(myPIPE_ID, i, "P-" & Format(i, "00000"))
      '    For j As Integer = 0 To shpPipe.Shape(i).numPoints - 1
      '        Debug.Print(j & ": key : " & shpPipe.Shape(i).Point(j).Key)
      '    Next
      'Next
      'shpPipe.StopEditingTable()

      shpPipe.StartEditingShapes()
      For i As Integer = 0 To 1 'shpPipe.NumShapes
         For j As Integer = 0 To shpPipe.Shape(i).numPoints - 1
            shpPipe.Shape(i).Point(j).Z = j
         Next
      Next
      shpPipe.StopEditingShapes()
      shpPipe.Save()

      'Report
      For i As Integer = 0 To 1 'shpPipe.NumShapes
         For j As Integer = 0 To shpPipe.Shape(i).numPoints - 1
            Debug.Print(j & ": key : " & shpPipe.Shape(i).Point(j).Z)
         Next
      Next

   End Sub

   Private Sub cmdGetPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetpoint.Click

   End Sub


   Function newPipeShp(ByVal x() As Double, ByVal y() As Double, ByVal AttributeID As Integer, ByRef sf As MapWinGIS.Shapefile) As Boolean
      On Error GoTo Error_handler

      Dim success As Boolean
      Dim shapeIndex As Long
      'objects
      Dim shape As MapWinGIS.Shape
      Dim point As MapWinGIS.Point

      'Start Editing it...
      success = sf.StartEditingShapes(True)
      shapeIndex = sf.NumShapes
      point = New MapWinGIS.Point
      Dim npt As Integer = UBound(x)

      'Create and add shape
      shape = New MapWinGIS.Shape

      'Create a new Point shape
      With shape
         success = .Create(MapWinGIS.ShpfileType.SHP_POLYLINE)
         If Not success Then
            MsgBox("Error in creating shape: " & .ErrorMsg(.LastErrorCode))
            GoTo Error_handler
         End If
      End With ' shape

      'Set the values for the point to be inserted

      'shape = New MapWinGIS.Shape
      shape.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
      'Set up the points

      'Insert the point into the shape

      For i As Integer = 0 To npt - 1
         point = New MapWinGIS.Point
         point.x = x(i + 1)
         point.y = y(i + 1)

         'Add the points to a shape
         success = shape.InsertPoint(point, i)
         If Not success Then
            MsgBox("Error in adding point: " & shape.ErrorMsg(shape.LastErrorCode))
            GoTo Error_handler
         End If
      Next i

      'Insert the shape into the shapefile
      With sf
         success = .EditInsertShape(shape, shapeIndex)
         If Not success Then
            MsgBox("Error in adding field: " & .ErrorMsg(.LastErrorCode))
            GoTo Error_handler
         End If
      End With ' sf

      'Add value to at least one attribute
      'Use shapeindex as dummy value:

      With sf
         For i As Integer = 0 To sf.NumFields - 1
            success = .EditCellValue(i, shapeIndex, sf.CellValue(i, AttributeID))
         Next

         If Not success Then
            MsgBox("Error in adding field: " & .ErrorMsg(.LastErrorCode))
            GoTo Error_handler
         End If
      End With ' sf
      'success = sf.Save

      'Stop editing shapes in the shapefile, saving changes to shapes,
      'also stopping editing of the attribute table
      With sf
         success = sf.StopEditingShapes(True, True)
         If Not success Then
            MsgBox("Error in adding field: " & .ErrorMsg(.LastErrorCode))
            GoTo Error_handler
         End If
      End With ' sf

      newPipeShp = True

Cleanup:
      On Error Resume Next
      shape = Nothing
      point = Nothing
      sf = Nothing

      Exit Function

Error_handler:
      newPipeShp = False
      GoTo Cleanup
   End Function


   Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
      Dim sf As New MapWinGIS.Shapefile
      'sf.CreateSpatialIndex("L1.SHP")
      'sf.Open("L1.SHP")

      'sf.CreateSpatialIndex("PIPE_LINES.SHP")
      sf.Open("PIPE_LINES.SHP")

      'AxMap1.addlayer(sf, True)

      Dim arBrkList(sf.NumShapes - 1) As List(Of brkStruct)
      For i As Integer = 0 To sf.NumShapes - 1
         arBrkList(i) = New List(Of brkStruct)
      Next




      'For i = 0 To sf.NumShapes - 1
      '    Dim c1 As New brkStruct
      '    With c1
      '        .id = i
      '        .x = 0
      '        .y = 0
      '        .brkPart = 0
      '    End With
      '    arBrkList(0).Add(c1)
      'Next

      For i As Integer = 0 To sf.NumShapes - 1
         Dim sft As Object
         sf.SelectShapes(sf.Shape(i).Extents, 0, MapWinGIS.SelectMode.INTERSECTION, sft)
         'sf.SelectByShapefile(sf.Shape(i), MapWinGIS.tkSpatialRelation.srIntersects, False, sft)

         'MsgBox(sft.length)

         Dim shpi As New MapWinGIS.Shape
         shpi = sf.Shape(i)

         Me.Text = i & " : Numpoint = " & shpi.numPoints
         For j As Integer = 0 To sft.length - 1
            If i <> sft(j) Then
               Dim shpj As New MapWinGIS.Shape
               shpj = sf.Shape(sft(j))
               Dim xx() As Double
               Dim yy() As Double
               Dim part() As Integer
               If chkShpIntersect(shpi, shpj, xx, yy, part) Then
                  Debug.Print("intersect at : " & i & ", " & sft(j))
                  For ii As Int16 = 0 To UBound(xx)

                     Dim pt As New MapWinGIS.Point
                     pt.x = xx(ii) : pt.y = yy(ii)
                     Debug.Print("  x=" & pt.x & ", y=" & pt.y)

                     Dim c1 As New brkStruct()
                     With c1
                        .id = i
                        .x = xx(ii)
                        .y = yy(ii)
                        .brkPart = part(ii)
                     End With
                     arBrkList(i).Add(c1)
                  Next

                  'Dim splitShp() As MapWinGIS.Shape
                  'createSplitLineOnePtAr(shpi, xx, yy, part, splitShp)
                  'Debug.Print("shp1:" & splitShp(0).numPoints & ", shp2:" & splitShp(0).numPoints)
                  'Dim ut As New MapWinGeoProc.Utils
                  'ut.PointOnLine(pt, shpi, 0)
                  'Debug.Print("PointInThisPoly=" & ut.PointOnLine(pt, shpi, 0))

               End If
            End If
         Next
      Next

      'test compare/sort
      'Dim Comp As New brkStructComparer
      'Now lets use our interface to sort
      'arBrkList(2).Sort(brkStruct, Comp)

      sf.Close()
      sf = Nothing

   End Sub

   Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click

   End Sub

   Private Sub cmdLinkBld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLinkBld.Click

   End Sub
End Class
