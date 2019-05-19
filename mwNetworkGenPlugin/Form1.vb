Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing.Color
Imports System.IO
Imports ZedGraph
Imports System.Math
Imports System.Threading
Imports MySql
Imports System.Data



Public Class Form1
    Private m_GraphThread As Thread


    Private blInit As Boolean '=True if a graph has been initialized
    Private blCoord As Boolean '=True if node coordinates exist
    Dim LinkUnconnect() As Integer
    Dim Q() As Double

    Declare Sub ENepanet Lib "epanet2.dll" (ByVal f1 As String, ByVal f2 As String, ByVal f3 As String)

    Declare Sub ENopen Lib "epanet2.dll" (ByVal inpfile As String, ByVal reportfile As String, ByVal binaryresultfile As String)
    Declare Sub ENsaveinpfile Lib "epanet2.dll" (ByVal inpfile As String)
    Declare Sub ENclose Lib "epanet2.dll" ()

    Declare Sub ENsolveH Lib "epanet2.dll" ()
    Declare Sub ENsaveH Lib "epanet2.dll" ()
    Declare Sub ENopenH Lib "epanet2.dll" ()
    Declare Sub ENinitH Lib "epanet2.dll" (ByVal timestep As Integer)
    Declare Sub ENrunH Lib "epanet2.dll" (ByRef timestep As Integer)
    Declare Sub ENnextH Lib "epanet2.dll" (ByRef timestep As Integer)
    Declare Sub ENcloseH Lib "epanet2.dll" ()
    Declare Sub ENsavehydfile Lib "epanet2.dll" (ByVal hydfile As String)
    Declare Sub ENusehydfile Lib "epanet2.dll" (ByVal hydfile As String)

    Declare Sub ENsolveQ Lib "epanet2.dll" ()
    Declare Sub ENopenQ Lib "epanet2.dll" ()
    Declare Sub ENinitQ Lib "epanet2.dll" (ByVal timestep As Integer)
    Declare Sub ENrunQ Lib "epanet2.dll" (ByRef timestep As Integer)
    Declare Sub ENstepQ Lib "epanet2.dll" (ByRef timestep As Integer)
    Declare Sub ENcloseQ Lib "epanet2.dll" ()

    Declare Sub ENwriteline Lib "epanet2.dll" (ByVal line As String)
    Declare Sub ENreport Lib "epanet2.dll" ()
    Declare Sub ENresetreport Lib "epanet2.dll" ()
    Declare Sub ENsetreport Lib "epanet2.dll" ()

    Declare Sub ENgetcontrol Lib "epanet2.dll" (ByVal cindex As Integer, ByRef ictype As Integer, ByRef lindex As Integer, ByRef setting As Single, ByRef nindex As Integer, ByRef level As Single)
    Declare Sub ENgetcount Lib "epanet2.dll" (ByVal type As Integer, ByRef count As Integer)
    Declare Sub ENgetoption Lib "epanet2.dll" (ByVal code As Integer, ByRef value As Single)
    Declare Sub ENgettimeparam Lib "epanet2.dll" (ByVal code As Integer, ByRef value As Single)
    Declare Sub ENgetflowunits Lib "epanet2.dll" (ByRef code As Integer)
    Declare Sub ENgetpatternindex Lib "epanet2.dll" (ByVal id As String, ByRef index As Integer)
    Declare Sub ENgetpatternid Lib "epanet2.dll" (ByVal index As Integer, ByRef id As String)
    Declare Sub ENgetpatternlen Lib "epanet2.dll" (ByVal index As Integer, ByRef len As Integer)
    Declare Sub ENgetpatternvalue Lib "epanet2.dll" (ByVal index As Integer, ByVal period As Integer, ByRef value As Single)
    Declare Sub ENgetqualtype Lib "epanet2.dll" (ByRef qualcode As Integer, ByRef tracenode As Integer)
    Declare Sub ENgeterror Lib "epanet2.dll" (ByVal errcode As Integer, ByRef errmsg As String, ByVal n As Integer)

    Declare Sub ENgetnodeindex Lib "epanet2.dll" (ByVal id As String, ByRef index As Integer)
    Declare Sub ENgetnodeid Lib "epanet2.dll" (ByVal index As Integer, ByVal id As String)
    Declare Sub ENgetnodetype Lib "epanet2.dll" (ByVal index As Integer, ByRef code As Integer)
    Declare Sub ENgetnodevalue Lib "epanet2.dll" (ByVal timestep As Integer, ByVal type As Integer, ByRef value As Single)
    Declare Sub ENgetlinkindex Lib "epanet2.dll" (ByVal id As String, ByRef index As Integer)
    Declare Sub ENgetlinkid Lib "epanet2.dll" (ByVal index As Integer, ByVal id As String)
    Declare Sub ENgetlinktype Lib "epanet2.dll" (ByVal index As Integer, ByRef code As Integer)
    Declare Sub ENgetlinknodes Lib "epanet2.dll" (ByVal index As Integer, ByRef node1 As Integer, ByRef node2 As Integer)
    Declare Sub ENgetlinkvalue Lib "epanet2.dll" (ByVal timestep As Integer, ByVal type As Integer, ByRef value As Single)

    Declare Sub ENgetversion Lib "epanet2.dll" (ByRef version As Integer)

    Declare Sub ENsetcontrol Lib "epanet2.dll" (ByVal cindex As Integer, ByVal ictype As Integer, ByVal lindex As Integer, ByVal setting As Single, ByVal nindex As Integer, ByVal level As Single)
    Declare Sub ENsetnodevalue Lib "epanet2.dll" (ByVal index As Integer, ByVal code As Integer, ByVal v As Single)
    Declare Sub ENsetlinkvalue Lib "epanet2.dll" (ByVal index As Integer, ByVal code As Integer, ByVal v As Single)
    Declare Sub ENaddpattern Lib "epanet2.dll" (ByVal id As String)
    Declare Sub ENsetpattern Lib "epanet2.dll" (ByVal index As Integer, ByRef f As Single, ByVal n As Integer)
    Declare Sub ENsetpatternvalue Lib "epanet2.dll" (ByVal index As Integer, ByVal period As Integer, ByVal value As Single)
    Declare Sub ENsettimeparam Lib "epanet2.dll" (ByVal code As Integer, ByVal value As Single)
    Declare Sub ENsetoption Lib "epanet2.dll" (ByVal code As Integer, ByVal v As Single)
    Declare Sub ENsetstatusreport Lib "epanet2.dll" (ByVal code As Integer)
    Declare Sub ENsetqualtype Lib "epanet2.dll" (ByVal qualcode As Integer, ByVal chemname As String, ByVal chemunits As String, ByVal tracenode As String)
    Declare Sub ENgetheadcurve Lib "epanet2.dll" (ByVal index As Integer, ByVal id As String)
    Declare Sub ENgetpumptype Lib "epanet2.dll" (ByVal index As Integer, ByVal type As Integer)

    ' Node parameters
    'Public Const EN_ELEVATION = 0
    'Public Const EN_BASEDEMAND = 1
    'Public Const EN_PATTERN = 2
    'Public Const EN_EMITTER = 3
    'Public Const EN_INITQUAL = 4
    'Public Const EN_SOURCEQUAL = 5
    'Public Const EN_SOURCEPAT = 6
    'Public Const EN_SOURCETYPE = 7
    'Public Const EN_TANKLEVEL = 8
    'Public Const EN_DEMAND = 9
    'Public Const EN_HEAD = 10
    'Public Const EN_PRESSURE = 11
    'Public Const EN_QUALITY = 12
    'Public Const EN_SOURCEMASS = 13
    'Public Const EN_INITVOLUME = 14
    'Public Const EN_MIXMODEL = 15
    'Public Const EN_MIXZONEVOL = 16

    'Public Const EN_TANKDIAM = 17
    'Public Const EN_MINVOLUME = 18
    'Public Const EN_VOLCURVE = 19
    'Public Const EN_MINLEVEL = 20
    'Public Const EN_MAXLEVEL = 21
    'Public Const EN_MIXFRACTION = 22
    'Public Const EN_TANK_KBULK = 23
    ' Link parameters
    'Public Const EN_DIAMETER = 0
    'Public Const EN_LENGTH = 1
    'Public Const EN_ROUGHNESS = 2
    'Public Const EN_MINORLOSS = 3
    'Public Const EN_INITSTATUS = 4
    'Public Const EN_INITSETTING = 5
    'Public Const EN_KBULK = 6
    'Public Const EN_KWALL = 7
    'Public Const EN_FLOW = 8
    'Public Const EN_VELOCITY = 9
    'Public Const EN_HEADLOSS = 10
    'Public Const EN_STATUS = 11
    'Public Const EN_SETTING = 12
    'Public Const EN_ENERGY = 13

    ' Time parameters
    'Public Const EN_DURATION = 0
    'Public Const EN_HYDSTEP = 1
    'Public Const EN_QUALSTEP = 2
    'Public Const EN_PATTERNSTEP = 3
    'Public Const EN_PATTERNSTART = 4
    'Public Const EN_REPORTSTEP = 5
    'Public Const EN_REPORTSTART = 6
    'Public Const EN_RULESTEP = 7
    'Public Const EN_STATISTIC = 8
    'Public Const EN_PERIODS = 9

    'Public Const EN_NODECOUNT = 0
    'Public Const EN_TANKCOUNT = 1
    'Public Const EN_LINKCOUNT = 2
    'Public Const EN_PATCOUNT = 3
    'Public Const EN_CURVECOUNT = 4
    'Public Const EN_CONTROLCOUNT = 5

    'Node types
    'Public Const EN_JUNCTION = 0
    'Public Const EN_RESERVOIR = 1
    'Public Const EN_TANK = 2

    'Link types
    'Public Const EN_CVPIPE = 0
    'Public Const EN_PIPE = 1
    'Public Const EN_PUMP = 2
    'Public Const EN_PRV = 3
    'Public Const EN_PSV = 4
    'Public Const EN_PBV = 5
    'Public Const EN_FCV = 6
    'Public Const EN_TCV = 7
    'Public Const EN_GPV = 8

    ' Quality analysis types
    'Public Const EN_NONE = 0
    'Public Const EN_CHEM = 1
    'Public Const EN_AGE = 2
    'Public Const EN_TRACE = 3

    ' Source quality types
    'Public Const EN_CONCEN = 0
    'Public Const EN_MASS = 1
    'Public Const EN_SETPOINT = 2
    'Public Const EN_FLOWPACED = 3

    ' Flow unit types
    'Public Const EN_CFS = 0
    'Public Const EN_GPM = 1
    'Public Const EN_MGD = 2
    'Public Const EN_IMGD = 3
    'Public Const EN_AFD = 4
    'Public Const EN_LPS = 5
    'Public Const EN_LPM = 6
    'Public Const EN_MLD = 7
    'Public Const EN_CMH = 8
    'Public Const EN_CMD = 9

    ' Mistc, options
    'Public Const EN_TRIALS = 0
    'Public Const EN_ACCURACY = 1
    'Public Const EN_TOLERANCE = 2
    'Public Const EN_EMITEXPON = 3
    'Public Const EN_DEMANDMULT = 4

    ' Control types
    'Public Const EN_LOWLEVEL = 0
    'Public Const EN_HILEVEL = 1
    'Public Const EN_TIMER = 2
    'Public Const EN_TIMEOFDAY = 3

    'Time statistics types
    'Public Const EN_AVERAGE = 1
    'Public Const EN_MINIMUM = 2
    'Public Const EN_MAXIMUM = 3
    'Public Const EN_RANGE = 4

    ' Tank mixing models
    'Public Const EN_MIX1 = 0
    'Public Const EN_MIX2 = 1
    'Public Const EN_FIFO = 2
    'Public Const EN_LIFO = 3

    'Public Const EN_NOSAVE = 0   ' Save-results-to-file flag
    'Public Const EN_SAVE = 1

    'Public Const EN_INITFLOW = 10   ' Re-initialize flows flag 

    'Public Const EN_CONST_HP = 0   ' constant horsepower          
    'Public Const EN_POWER_FUNC = 1 'power function
    'Public Const EN_CUSTOM = 2  ' user-defined custom curve
 


#Region "FormControl"

    Private Sub frmDmaDesign_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
       
        'TODO: This line of code loads data into the 'PumpDataSet.Pump' table. You can move, or remove it, as needed.

        ClearComboLayer()
        loadcomboLayer()

        TabMain.TabPages.Remove(TP_Design)
        TabMain.TabPages.Remove(TP_analysis)
        TC1.TabPages.Remove(TP_path)
        TC1.TabPages(2).Enabled = False
    End Sub

   
    

#End Region

#Region "Form Utilty"
    'Clear Graphic
    Sub clearGraphic()
        g_MW.View.Draw.ClearDrawings()
        g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
    End Sub

    'Clear combobox
    Sub ClearComboLayer()
        cmbDEM.Items.Clear()
        cmbNetwork.Items.Clear()
        CmbHydrant.Items.Clear()
    End Sub

    'Load combobox
    Sub loadcomboLayer()
        ClearComboLayer()
        For i As Integer = 0 To g_MW.Layers.NumLayers - 1
            Try
                'Line type
                If g_MW.Layers(i).LayerType = MapWindow.Interfaces.eLayerType.LineShapefile Then
                    cmbNetwork.Items.Add(g_MW.Layers(i).Name)
                    cmbNetwork.Text = g_MW.Layers(i).Name
                End If
            Catch ex As Exception

            End Try
            Try
                'point type
                If g_MW.Layers(i).LayerType = MapWindow.Interfaces.eLayerType.PointShapefile Then
                    CmbHydrant.Items.Add(g_MW.Layers(i).Name)
                    CmbHydrant.Text = g_MW.Layers(i).Name
                End If
            Catch ex As Exception

            End Try

            Try
                'Grid type
                If g_MW.Layers(i).LayerType = MapWindow.Interfaces.eLayerType.Grid Or _
                    g_MW.Layers(i).LayerType = MapWindow.Interfaces.eLayerType.Image Then
                    cmbDEM.Items.Add(g_MW.Layers(i).Name)
                    cmbDEM.Text = g_MW.Layers(i).Name
                End If
            Catch ex As Exception

            End Try

        Next
    End Sub
    'Set Progress bar
    Sub SetProgressBar(ByVal maxVal As Integer, Optional ByVal status As Boolean = True)
        UpdateProgress.Maximum = maxVal
        UpdateProgress.Minimum = 0
        UpdateProgress.Value = 0
        UpdateProgress.Visible = status
    End Sub

#End Region

#Region "GIS Utilities"



    Private Sub drawShape(ByVal shp As MapWinGIS.Shape, ByVal color As System.Drawing.Color, Optional ByVal LineWidth As Integer = 3) ', Optional ByVal zoomToShp As Boolean = False)
        Try
            g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
            For i As Integer = 0 To shp.numPoints - 2
                Dim x1 As Double = shp.Point(i).x
                Dim y1 As Double = shp.Point(i).y
                Dim x2 As Double = shp.Point(i + 1).x
                Dim y2 As Double = shp.Point(i + 1).y
                g_MW.View.Draw.DrawLine(x1, y1, x2, y2, LineWidth, color)
            Next
        Catch ex As Exception

        End Try
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



    Sub addShapeAttribute(ByRef Shp As MapWinGIS.Shapefile, ByVal SourceID As Integer, ByVal TargetID As Integer)
        Shp.StartEditingTable()
        For i As Integer = 0 To Shp.NumFields - 1
            Shp.EditCellValue(i, TargetID, Shp.CellValue(i, SourceID))
        Next
        Shp.StopEditingTable()

        Shp.Save()
    End Sub



#End Region

#Region "2. Network generation TAB"

    Private Sub cmdAutoGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAutoGen.Click
        ' Try


        ' step 0: Clear tables and set properties
        clearGraphic()
        grdNode.Rows.Clear()
        grdLink.Rows.Clear()




        '--------------------------------------------------------------------------------
        grdNode.GridColor = System.Drawing.Color.Blue
        grdNode.CellBorderStyle = DataGridViewCellBorderStyle.None
        grdNode.BackgroundColor = System.Drawing.Color.LightGray

        grdNode.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.Aquamarine
        grdNode.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Blue

        grdNode.DefaultCellStyle.WrapMode = DataGridViewTriState.[True]

        grdNode.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grdNode.AllowUserToResizeColumns = False

        grdNode.RowsDefaultCellStyle.BackColor = System.Drawing.Color.AliceBlue
        grdNode.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.Beige
        '--------------------------------------------------------------------------------

        grdLink.GridColor = System.Drawing.Color.Blue
        grdLink.CellBorderStyle = DataGridViewCellBorderStyle.None
        grdLink.BackgroundColor = System.Drawing.Color.LightGray

        grdLink.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.Aquamarine
        grdLink.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Blue

        grdLink.DefaultCellStyle.WrapMode = DataGridViewTriState.[True]

        grdLink.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grdLink.AllowUserToResizeColumns = False

        grdLink.RowsDefaultCellStyle.BackColor = System.Drawing.Color.AliceBlue
        grdLink.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.Beige



        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        'step1: read GIS files
        Dim NetworkShp, hydrantShp As New MapWinGIS.Shapefile
        Dim DEM As New MapWinGIS.Grid
        NetworkShp = g_MW.Layers(getLayerHandle(cmbNetwork.Text)).GetObject
        hydrantShp = g_MW.Layers(getLayerHandle(CmbHydrant.Text)).GetObject
        Try
            DEM.Open(g_MW.Layers(getLayerHandle(cmbDEM.Text)).FileName)
        Catch ex As Exception

        End Try


        Dim IdDiam, IdArea, IdhydQ, idrough As Integer


        If CMBArea.Text <> CMBDisch.Text Then
            For i As Integer = 0 To hydrantShp.NumFields - 1
                If hydrantShp.Field(i).Name = CMBArea.Text Then
                    IdArea = i
                End If
                If hydrantShp.Field(i).Name = CMBDisch.Text Then
                    IdhydQ = i
                End If
            Next
        Else
            MessageBox.Show("Area field cannot be the same as Discharge field")
            Exit Sub
        End If


        If CmbDiam.Text <> CMBrough.Text Then
            For i As Integer = 0 To NetworkShp.NumFields - 1
                If NetworkShp.Field(i).Name = CmbDiam.Text Then
                    IdDiam = i
                End If
                If NetworkShp.Field(i).Name = CMBrough.Text Then
                    idrough = i
                End If
            Next
        Else
            MessageBox.Show("Rougness field cannot be the same as Diameters field")
            Exit Sub
        End If



        'step2: Start numbering nodes
        Dim HashNode As New Hashtable
        Dim hashString As String
        Dim NodeF, NodeT As Integer
        Dim NODEid As Integer = 1 ' Start Node at node no. 1

        UpdateProgress.Maximum = NetworkShp.NumShapes
        UpdateProgress.Minimum = 0
        UpdateProgress.Visible = True


        Dim j As Integer

        For i As Integer = 0 To NetworkShp.NumShapes - 1
            Application.DoEvents()
            UpdateProgress.Value = i
            NodeT = -1
            NodeF = -1

            Dim pt1 As New MapWinGIS.Point
            Dim pt2 As New MapWinGIS.Point
            Dim hydpt As New MapWinGIS.Point

            Dim LastPt As Integer = NetworkShp.Shape(i).numPoints - 1
            pt1 = NetworkShp.Shape(i).Point(0)
            pt2 = NetworkShp.Shape(i).Point(LastPt)

            For jj As Integer = 1 To 2
                Try
                    'a) Assign XY coordinate to the node  
                    Dim pt As New MapWinGIS.Point
                    Select Case jj
                        Case 1 'First node
                            pt.x = pt1.x
                            pt.y = pt1.y
                        Case 2 'Terminal node
                            pt.x = pt2.x
                            pt.y = pt2.y
                    End Select
                    hashString = Math.Round(pt.x, 2) & "," & Math.Round(pt.y, 2)
                    HashNode.Add(hashString, NODEid)
                    NODEid += 1

                    grdNode.Rows.Add()
                    Dim ActRow As Integer = grdNode.RowCount - 1
                    Select Case jj
                        Case 1
                            NodeF = HashNode(hashString)
                            grdNode.Item(0, ActRow).Value = NodeF
                        Case 2
                            NodeT = HashNode(hashString)
                            grdNode.Item(0, ActRow).Value = NodeT
                    End Select



                    grdNode.Item(1, ActRow).Value = pt.x
                    grdNode.Item(2, ActRow).Value = pt.y

                    'b) Extract Land eleveation of hydrant from DEM
                    Try
                        Dim ColDem, RowDEM As Integer
                        DEM.ProjToCell(pt.x, pt.y, ColDem, RowDEM)
                        grdNode.Item(5, ActRow).Value = DEM.Value(ColDem, RowDEM)
                    Catch ex As Exception
                        grdNode.Item(5, ActRow).Value = 0
                    End Try


                    'c)  identify node type(3) /Discharge (4) /irrigated area(6)

                    If ActRow = 0 Then ' First point by default is reservoir
                        grdNode.Item(3, ActRow).Value = "Reservoir"
                        grdNode.Item(4, ActRow).Value = 0
                        grdNode.Item(6, ActRow).Value = 0
                        g_MW.View.Draw.DrawWideCircle(pt.x, pt.y, 8, System.Drawing.Color.Black, True, 8)
                    Else
                        Dim ff As Integer = 0 ' differentiate between hydrant (buffer 0.5m) and node
                        For j = 0 To hydrantShp.NumShapes - 1
                            hydpt = hydrantShp.Shape(j).Point(0)
                            ff = j
                            Dim buff As Integer = 0
                            If (hydpt.x <= pt.x + buff) And (hydpt.x >= pt.x - buff) And (hydpt.y <= pt.y + buff) And (hydpt.y >= pt.y - buff) Then 'buffer of 0.5 m from hydrant shapefile
                                grdNode.Item(3, ActRow).Value = "Hydrant"
                                grdNode.Item(4, ActRow).Value = hydrantShp.CellValue(IdhydQ, j)
                                grdNode.Item(6, ActRow).Value = hydrantShp.CellValue(IdArea, j) 'Area
                                g_MW.View.Draw.DrawPoint(pt.x, pt.y, 4, System.Drawing.Color.Blue)
                                Exit For
                            End If
                            If ff = hydrantShp.NumShapes - 1 Then
                                grdNode.Item(3, ActRow).Value = "Node"
                                grdNode.Item(4, ActRow).Value = 0 'Discharge
                                grdNode.Item(6, ActRow).Value = 0 'Area
                                g_MW.View.Draw.DrawPoint(pt.x, pt.y, 2, System.Drawing.Color.Azure)
                            End If
                        Next
                    End If
                Catch ex As Exception
                    Select Case jj
                        Case 1
                            NodeF = HashNode(hashString)
                        Case 2
                            NodeT = HashNode(hashString)
                    End Select
                End Try
            Next 'jj

            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            'step2: Start numbering pipes
            Dim f As Integer = getColNum("F", NetworkShp) 'First node
            Dim t As Integer = getColNum("T", NetworkShp) 'Terminal node
            grdLink.Rows.Add()
            grdLink.Item(0, i).Value = i + 1 'NetworkShp.CellValue(FieldNetwork.LinkID, i) ' Link ID
            grdLink.Item(1, i).Value = NodeF 'First
            grdLink.Item(2, i).Value = NodeT 'Termianl
            grdLink.Item(3, i).Value = Format(Math.Round(NetworkShp.Shape(i).Length, 3), "0.000") 'Pipe length
            ' grdLink.Item(4, i).Value = Q clement column
            grdLink.Item(5, i).Value = NetworkShp.CellValue(IdDiam, i) ' Diameter
            grdLink.Item(6, i).Value = NetworkShp.CellValue(idrough, i) '0.0015 'pipe roughness
        Next
        UpdateProgress.Visible = False
        cmdSaveFile.Visible = True
        cmdcreatenet.Visible = True
        TC1.TabPages(2).Enabled = True

        MsgBox("complete")
        'Catch ex As Exception
        '    MsgBox("Error has occured, check your input file")
        'End Try
    End Sub



#End Region

#Region "3. Network Editor TAB"

#Region "  3.1 Network checking process"





    Private Sub cmdClearDrawing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        clearGraphic()
    End Sub


#End Region



#End Region


    Private Sub cmdInitialNET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdcreatenet.Click
        Try

            TabMain.TabPages.Remove(TP_Design)
            TabMain.TabPages.Remove(TP_analysis)
            TC1.TabPages.Remove(TP_path)


            Dim x1, y1 As Double

            clearGraphic()
            nNode = grdNode.RowCount
            HashNodeID.Clear()
            ReDim Node(grdNode.RowCount)

            UpdateProgress.Visible = True
            UpdateProgress.Minimum = 0
            UpdateProgress.Maximum = grdNode.RowCount - 1


            nres = 0
            CB_Cl_Res.Items.Clear()
            CB_res_design.Items.Clear()
            For i As Integer = 0 To grdNode.RowCount - 1
                Try


                    UpdateProgress.Value = i
                    Node(i + 1).id = grdNode.Item(0, i).Value
                    Node(i + 1).x = grdNode.Item(1, i).Value
                    Node(i + 1).y = grdNode.Item(2, i).Value
                    Node(i + 1).Type = grdNode.Item(3, i).Value
                    Select Case Node(i + 1).Type
                        Case "Reservoir"
                            nres += 1
                            ReDim Preserve ResId(nres)
                            ResId(nres) = Node(i + 1).id
                            CB_Cl_Res.Items.Add(Node(i + 1).id)
                            CB_res_design.Items.Add(Node(i + 1).id)
                            If nres = 1 Then
                                CB_res_design.Text = Node(i + 1).id
                            End If
                    End Select
                    Node(i + 1).Q = grdNode.Item(4, i).Value

                    Node(i + 1).z = grdNode.Item(5, i).Value
                    Node(i + 1).A = grdNode.Item(6, i).Value
                    HashNodeID.Add(Node(i + 1).id, i)
                    ' g_MW.View.Draw.DrawPoint(Node(i + 1).x, Node(i + 1).y, 3, Drawing.Color.Blue)
                Catch ex As Exception
                    MsgBox("Error occured at node " & i & "!")
                    Exit Sub
                End Try
            Next
            If nres > 1 Then
                GB_Clement.Visible = True
                CB_res_design.Visible = True
                Label12.Visible = True
            Else
                GB_Clement.Visible = False
                CB_res_design.Visible = False
                Label12.Visible = False
            End If

            nLink = grdLink.RowCount
            ReDim Link(grdLink.RowCount)
            HashLinkID.Clear()
            UpdateProgress.Maximum = nLink
            UpdateProgress.Minimum = 0

            For i As Integer = 0 To grdLink.RowCount - 1
                Try

                    UpdateProgress.Value = i
                    Dim F As Integer = grdLink.Item(1, i).Value
                    Dim T As Integer = grdLink.Item(2, i).Value

                    x1 = Node(F).x
                    y1 = Node(F).y
                    x1 = Node(T).x
                    y1 = Node(T).y
                    Link(i + 1).NodeF = F
                    Link(i + 1).NodeT = T
                    Link(i + 1).id = i + 1
                    Link(i + 1).L = grdLink.Item(3, i).Value
                    Try
                        Link(i + 1).Discharge = grdLink.Item(4, i).Value 'Qclement
                    Catch ex As Exception

                    End Try
                    Link(i + 1).diam = grdLink.Item(5, i).Value
                    Link(i + 1).rough = grdLink.Item(6, i).Value

                    ' Create HashTable for Link ID and Node ID
                    Try
                        HashLinkID.Add(F & "-" & T, i + 1)
                        HashLinkID.Add(T & "-" & F, i + 1)
                    Catch ex As Exception
                    End Try
                Catch ex As Exception
                    MsgBox("Error occured at pipe " & i & "!")
                    Exit Sub
                End Try
            Next


        npump = DG_pump.Rows.Count
        If npump >= 1 Then
            Dim con As New OleDb.OleDbConnection
            Dim ds As New System.Data.DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim sql As String
            Dim dbProvider As String
            Dim dbSource As String

            dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
            dbSource = "Data Source=" & myStream

            con.ConnectionString = dbProvider & dbSource
            con.Open()

            sql = "select * FROM Pumptbl"
            da = New OleDb.OleDbDataAdapter(sql, con)
            da.Fill(ds, "pump table") ' You can call it "Bacon Sandwish" and will still work
            con.Close() ' ds is now the table with the infor


            ReDim pumpcurve(ds.Tables("pump table").Rows.Count)
            For i As Integer = 1 To ds.Tables("pump table").Rows.Count
                Dim str As String = ds.Tables("pump table").Rows(i - 1).Item(1) & " " & ds.Tables("pump table").Rows(i - 1).Item(2) & " " & ds.Tables("pump table").Rows(i - 1).Item(3)
                pumpcurve(i) = str
            Next
        End If


        UpdateProgress.Visible = False
        MsgBox("Network was created")

        'Other buttons need to be disabled too
        btnpath.Enabled = True ' Draw pathway
        btn_checkpipe.Enabled = True
        btnClear.Visible = True

        TabMain.TabPages.Add(TP_Design)
        TabMain.TabPages.Add(TP_analysis)
        TC1.TabPages.Add(TP_path)
        TC5.TabPages.Remove(TPLoop)
        Catch ex As Exception
            MsgBox("Check that you have opened/created a network")
        End Try
    End Sub









    Sub save_inputfile(infile As String)
        Using sw As StreamWriter = New StreamWriter(infile)
            sw.WriteLine("[Node]")
            sw.WriteLine(grdNode.Rows.Count)

            For i As Integer = 0 To grdNode.Rows.Count - 1
                For j As Integer = 0 To 6
                    If j = 6 Then
                        sw.WriteLine(grdNode.Item(j, i).Value)
                    Else
                        sw.Write(grdNode.Item(j, i).Value & ";")
                    End If
                Next
            Next
            sw.WriteLine("[Pipe]")
            sw.WriteLine(grdLink.Rows.Count)
            For i As Integer = 0 To grdLink.Rows.Count - 1
                For j As Integer = 0 To 6
                    If j = 6 Then
                        sw.WriteLine(grdLink.Item(j, i).Value)
                    Else
                        sw.Write(grdLink.Item(j, i).Value & ";")
                    End If
                Next
            Next

        End Using
    End Sub



    Sub open_inputfile(infile As String)

        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(infile) '(myStream)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(";")
            MyReader.ReadLine() 'Read "[Node]"
            Dim currentRow As String()
            Dim nnode As Integer = MyReader.ReadLine()
            ' nnode = nnode - 1
            Dim currentField As String
            Dim j As Integer
            grdNode.Rows.Clear()
            For i As Integer = 0 To nnode - 1
                grdNode.Rows.Add()
                currentRow = MyReader.ReadFields()
                j = -1
                For Each currentField In currentRow
                    j = j + 1
                    grdNode.Item(j, i).Value = currentField
                Next
            Next

            grdLink.Rows.Clear()
            MyReader.ReadLine() 'Read "[Pipe]"
            Dim npipe As Integer = MyReader.ReadLine()
            'npipe = npipe - 1
            For i As Integer = 0 To npipe - 1
                grdLink.Rows.Add()
                currentRow = MyReader.ReadFields()
                j = -1
                For Each currentField In currentRow
                    j = j + 1
                    grdLink.Item(j, i).Value = currentField
                Next
            Next
        End Using
    End Sub


    Private Sub cmdSaveUnLinkFile_Click(sender As System.Object, e As System.EventArgs) Handles cmdSaveFile.Click
        Dim SFD As New SaveFileDialog
        SFD.Filter = "in. Files (*.in)|*.in"

        Dim infile As String
        If SFD.ShowDialog() = Windows.Forms.DialogResult.OK Then
            infile = SFD.FileName
            save_inputfile(infile)
            Try
                Create_Pipe_GIS(infolder)
                Create_node_GIS(infolder)
            Catch ex As Exception
                MsgBox("Error in converting input into a GIS file")
            End Try

        End If


    End Sub

    Sub Create_Pipe_GIS(infolder As String)

        Dim NetworkShp As New MapWinGIS.Shapefile
        NetworkShp = g_MW.Layers(getLayerHandle(cmbNetwork.Text)).GetObject
        Dim shp_pipe As MapWinGIS.ShapefileClass
        shp_pipe = New MapWinGIS.ShapefileClass

        Dim Pshape_name As String = infolder & "\Pipes.shp"

        NetworkShp.SaveAs(Pshape_name)
        'open the shapefile
        shp_pipe.Open(Pshape_name)
        shp_pipe.StartEditingTable()

        Dim n_field As Integer = NetworkShp.NumFields


        While n_field <> 1
            shp_pipe.EditDeleteField(1)
            n_field = shp_pipe.NumFields
        End While
        Dim fldIdx As Long = shp_pipe.NumFields
        For i As Integer = 1 To 7
            Dim newFld As New MapWinGIS.Field
            Select Case i
                Case 1
                    newFld.Name = "ID"
                    newFld.Width = 5
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 2
                    newFld.Name = "IN"
                    newFld.Width = 5
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 3
                    newFld.Name = "FN"
                    newFld.Width = 5
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 4
                    newFld.Name = "Length"
                    newFld.Width = 50
                    newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                Case 5
                    newFld.Name = "QCl"
                    newFld.Width = 50
                    newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                Case 6
                    newFld.Name = "Diam"
                    newFld.Width = 50
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 7
                    newFld.Name = "Rough"
                    newFld.Width = 50
                    newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
            End Select
            fldIdx = shp_pipe.NumFields
            shp_pipe.EditInsertField(newFld, fldIdx) 'insert column in dbf
        Next

        shp_pipe.EditDeleteField(0)

        For i As Integer = 0 To shp_pipe.NumFields - 1
            For j As Integer = 0 To shp_pipe.NumShapes - 1
                shp_pipe.EditCellValue(i, j, grdLink.Item(i, j).Value)
            Next
        Next

        shp_pipe.StopEditingTable(True)

    End Sub
    Sub Create_node_GIS(infolder As String)

        Dim Nshape_name As String = infolder & "\Nodes.shp"
        Dim FieldIndices As New Hashtable


            Dim sftype As MapWinGIS.ShpfileType
            sftype = MapWinGIS.ShpfileType.SHP_POINT


            Dim newSF As New MapWinGIS.Shapefile
        newSF.CreateNew(Nshape_name, sftype)
            newSF.StartEditingShapes(True)

            Dim i, j, jj, f As Integer
            'Add the fields:

        Dim idFld As New MapWinGIS.Field
        Dim fldIdx As Long = newSF.NumFields

            'Create Fields
        For j = 1 To 7
            Dim newFld As New MapWinGIS.Field

            Select Case j
                Case 1
                    newFld.Name = "ID"
                    newFld.Width = 5
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 2
                    newFld.Name = "X"
                    newFld.Width = 10
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 3
                    newFld.Name = "Y"
                    newFld.Width = 10
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 4
                    newFld.Name = "Type"
                    newFld.Width = 50
                    newFld.Type = MapWinGIS.FieldType.STRING_FIELD
                Case 5
                    newFld.Name = "Q"
                    newFld.Width = 10
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 6
                    newFld.Name = "Elev"
                    newFld.Width = 10
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 7
                    newFld.Name = "Area"
                    newFld.Width = 10
                    newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
            End Select

            fldIdx = newSF.NumFields
            newSF.EditInsertField(newFld, fldIdx)
            FieldIndices.Add(newFld.Name, fldIdx)
        Next
            ''
            ''******************************************************************
            '' Add Record to Shape file

            jj = -1
        For i = 1 To grdNode.Rows.Count
         
                Dim newShp As New MapWinGIS.Shape
                Dim newPt As New MapWinGIS.Point

                newShp.Create(sftype)

            newPt.x = grdNode.Item(1, i - 1).Value
            newPt.y = grdNode.Item(2, i - 1).Value
                newShp.InsertPoint(newPt, 0)
            newSF.EditInsertShape(newShp, i - 1)
            For f = 1 To 7
                newSF.EditCellValue(f - 1, i - 1, grdNode.Item(f - 1, i - 1).Value)
            Next
        Next
            newSF.StopEditingShapes(True, True)
            newSF.Close()

            MsgBox("The file has been converted.", MsgBoxStyle.Information, "Complete")


    End Sub

    Private Sub cmdOpenUnLinkFile_Click(sender As System.Object, e As System.EventArgs) Handles cmdOpenFile.Click
        Dim OFD As New OpenFileDialog
        OFD.Filter = "in. Files (*.in)|*.in"

        Dim infile As String
        If OFD.ShowDialog() = Windows.Forms.DialogResult.OK Then
            infile = OFD.FileName
            open_inputfile(infile)
            cmdSaveFile.Visible = True
            cmdcreatenet.Visible = True
            TC1.TabPages(2).Enabled = True
        End If

    End Sub








    Sub open_pipefile(pipefile As String)
        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(pipefile) '(myStream)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(";")
            Dim currentRow As String()
            Dim ndiam As Integer = MyReader.ReadLine()
            Nup_diam.Value = ndiam
            Dim currentField As String
            Dim j As Integer
            DG_diam.Rows.Clear()
            DG_cost.Rows.Clear()
            For i As Integer = 0 To ndiam - 1
                DG_diam.Rows.Add()
                currentRow = MyReader.ReadFields()
                j = -1
                For Each currentField In currentRow
                    j = j + 1
                    DG_diam.Item(j, i).Value = currentField
                    If j = 0 Then
                        DG_cost.Rows.Add()
                        DG_cost.Item(j, i).Value = currentField
                    End If
                Next
            Next
        End Using
    End Sub



    Sub save_pipefile(pipefile As String)
        Using sw As StreamWriter = New StreamWriter(pipefile)
            sw.WriteLine(DG_diam.Rows.Count)
            For i As Integer = 0 To DG_diam.Rows.Count - 1
                For j As Integer = 0 To 2
                    If j = 2 Then
                        sw.WriteLine(DG_diam.Item(j, i).Value)
                    Else
                        sw.Write(DG_diam.Item(j, i).Value & ";")
                    End If
                Next
            Next
        End Using
    End Sub



    Function solve_net(ReservoirID As Integer, ByVal Q() As Single, ByRef diam() As Single, ByRef rough() As Single, ByRef cost() As Single, ByRef diam1() As Single, ByRef rough1() As Single, ByRef cost1() As Single) As Single
        Dim hf As New LIDM
        Dim flowPath As PathProperties
        Dim D, o As Integer
        Main_SP.Main()



        o = ReservoirID
        Dim F, T As Integer
        UpdateProgress.Maximum = nNode
        UpdateProgress.Minimum = 0
        UpdateProgress.Visible = True

        Dim hf_tot As Single
        Dim hf_tot1 As Single
        Dim hfmax As Single = 0
        Dim id_hfmax As Single = 0 ' node number worst placed in the network

        clearGraphic()
        Dim Network As New MapWinGIS.Shapefile
        Network = g_MW.Layers(getLayerHandle(cmbNetwork.Text)).GetObject

        'Dim tStart As String = Now
        For iNode As Integer = 1 To nNode ' identify the worst placed hydrant
            hf_tot = 0
            If Node(iNode).Q <> 0 Then 'hydrant
                D = Node(iNode).id
                If o <> D Then
                    flowPath = createSPaths(FlowPathHyd, PipeNetwork, o, D)
                    For i As Integer = 1 To flowPath.nPt - 1
                        F = flowPath.pNode(i).id
                        T = flowPath.pNode(i + 1).id
                        Dim ShapeID As Integer = getLinkId(F, T)
                        hf_tot += hf.DW_headloss(rough(ShapeID), diam(ShapeID), Q(ShapeID), Link(ShapeID).L)
                    Next
                    hf_tot = hf_tot + (Node(D).z - Node(o).z)
                    If hfmax < hf_tot Then
                        hfmax = hf_tot
                        id_hfmax = D ' node id of the worst placed hydrant
                    End If
                End If
            End If

            UpdateProgress.Value = iNode
        Next

        D = id_hfmax ' 

        flowPath = createSPaths(FlowPathHyd, PipeNetwork, o, D)
        Dim beta(flowPath.nPt) As Single
        Dim beta_min As Single
        Dim beta_id As Integer

        Dim id As Integer
        Dim htot As Single = 0

        For i As Integer = 1 To flowPath.nPt - 1
            F = flowPath.pNode(i).id
            T = flowPath.pNode(i + 1).id
            Dim ShapeID As Integer = getLinkId(F, T)

            hf_tot = Math.Abs(hf.DW_headloss(rough(ShapeID), diam(ShapeID), Q(ShapeID), Link(ShapeID).L))
            htot += hf_tot

            id = increase_diam_id(diam(ShapeID))

            If id = 0 Then
                diam1(beta_id) = diam(beta_id)
                cost1(beta_id) = diam(beta_id)
                rough1(beta_id) = diam(beta_id)
            Else

                diam1(ShapeID) = Pipe_C(id).diam
                cost1(ShapeID) = Pipe_C(id).price
                rough1(ShapeID) = Pipe_C(id).rough

            End If

            If id = 0 Then
                beta(i) = 1000000 ' maximum diam has been reached
            Else
                hf_tot1 = Math.Abs(hf.DW_headloss(rough1(ShapeID), diam1(ShapeID), Q(ShapeID), Link(ShapeID).L))
                beta(i) = (Link(ShapeID).L * (cost1(ShapeID) - cost(ShapeID))) / (hf_tot - hf_tot1) ' calculate beta for all pipes
            End If


            ' determin the smallest beta value
            If i = 1 Then
                beta_min = beta(i)
                beta_id = ShapeID
            Else
                If beta_min >= beta(i) Then
                    beta_min = beta(i)
                    beta_id = ShapeID
                End If
            End If
        Next
        If htot + (Node(D).z - Node(o).z) + CSng(Hmin_txt.Text) < CSng(Hpiez_txt.Text) Then
            For j As Integer = 1 To nLink
                grdLink.Item(5, j - 1).Value = diam(j)
                Link(j).diam = diam(j)
            Next
            ' MessageBox.Show("solution was found")
            GoTo 1
        End If

        diam(beta_id) = diam1(beta_id)
        rough(beta_id) = rough1(beta_id)
        cost(beta_id) = cost1(beta_id)

        id = increase_diam_id(diam(beta_id))
        If id = 0 Then
            diam1(beta_id) = diam(beta_id)
            cost1(beta_id) = diam(beta_id)
            rough1(beta_id) = diam(beta_id)
        Else
            diam1(beta_id) = Pipe_C(id).diam
            cost1(beta_id) = Pipe_C(id).price
            rough1(beta_id) = Pipe_C(id).rough
        End If
        drawShape(Network.Shape(beta_id - 1), Drawing.Color.Red, 2)

1:      Return htot + (Node(D).z - Node(o).z) + CSng(Hmin_txt.Text)

    End Function

    Function increase_diam_id(ByVal diam As Single) As Integer
        Dim ndiam As Integer = DG_diam.Rows.Count
        Dim maxdiam As Single = DG_diam.Item(0, ndiam - 1).Value
        If diam = Pipe_C(ndiam).diam Then
            increase_diam_id = 0
        Else
            For j As Integer = 1 To ndiam - 1
                If Pipe_C(j).diam = diam Then
                    diam = Pipe_C(j + 1).diam
                    increase_diam_id = j + 1 ' rows number
                    Exit For
                End If
            Next
        End If
        Return increase_diam_id
    End Function


    Private Sub CheckBtn_Click(sender As System.Object, e As System.EventArgs)

    End Sub



    Private Sub PlotValue(ByVal old_y As Single, ByVal new_y As Single, ByVal ymax As Single)
        ' See if we're on the worker thread and thus
        ' need to invoke the main UI thread.
        If Me.InvokeRequired Then
            ' Make arguments for the delegate.
            Dim args As Object() = {old_y, new_y}

            ' Make the delegate.
            Dim plot_value_delegate As PlotValueDelegate
            plot_value_delegate = AddressOf PlotValue

            ' Invoke the delegate on the main UI thread.
            Me.Invoke(plot_value_delegate, args)

            ' We're done.
            Exit Sub
        End If

        ' Make the Bitmap and Graphics objects.
        Dim wid As Integer = picGraph.ClientSize.Width
        Dim hgt As Integer = picGraph.ClientSize.Height
        Dim bm As New Bitmap(wid, hgt)
        Dim gr As Graphics = Graphics.FromImage(bm)
        Dim m_Ymid As Integer
        ' Dim GRID_STEP As Integer = 10 ' Convert.ToInt32(txtGridSpacing.Text) 'Assign Grid Spacing
        '  m_Ymid = hgt
        ' Move the old data one pixel to the left.
        Try
            gr.DrawImage(picGraph.Image, -1, 0)
        Catch ex As Exception

        End Try


        ' Erase the right edge and draw guide lines.
        ' gr.DrawLine(Pens.Blue, wid - 10, 0, wid - 10, hgt - 10)
        'For i As Integer = m_Ymid To picGraph.ClientSize.Height Step GRID_STEP
        '    gr.DrawLine(Pens.LightBlue, wid - 2, i, wid - 1, i)
        'Next i
        'For i As Integer = m_Ymid To 0 Step -GRID_STEP
        '    gr.DrawLine(Pens.LightBlue, wid - 2, i, wid - 1, i)
        'Next i

        ' Plot a new pixel
        old_y = hgt - (hgt / ymax) * old_y
        new_y = hgt - (hgt / ymax) * new_y

        Dim p As New Pen(System.Drawing.Color.Blue, 5)

        gr.DrawLine(p, wid - 2, old_y, wid - 1, new_y)


        ' Display the result.
        picGraph.Image = bm
        picGraph.Refresh()

        gr.Dispose()
    End Sub

    Private Delegate Sub PlotValueDelegate(ByVal old_y As Single, ByVal new_y As Single, ByVal ymax As Single)










    Private Sub grdLink_DoubleClick(sender As System.Object, e As System.EventArgs) Handles grdLink.DoubleClick
        Try

            clearGraphic()
            Dim pipeShp As New MapWinGIS.Shapefile
            Dim pipeLyr As Integer = getLayerHandle(cmbNetwork.Text)
            pipeShp = g_MW.Layers(pipeLyr).GetObject

            For i As Integer = 0 To grdLink.Rows.Count - 1
                If grdLink.Rows(i).Selected = True Then
                    Dim id As Integer = grdLink.Item(0, i).Value
                    drawShape(pipeShp.Shape(i), System.Drawing.Color.Yellow, 3)
                    Exit Sub
                End If

            Next

        Catch ex As Exception

        End Try
    End Sub

    Private Sub grdNode_DoubleClick(sender As System.Object, e As System.EventArgs) Handles grdNode.DoubleClick
        Try
            clearGraphic()
            Dim hydShp As New MapWinGIS.Shapefile
            Dim hydLyr As Integer = getLayerHandle(CmbHydrant.Text)
            hydShp = g_MW.Layers(hydLyr).GetObject

            For i As Integer = 0 To grdNode.Rows.Count - 1
                If grdNode.Rows(i).Selected = True Then
                    Dim id As Integer = grdNode.Item(0, i).Value

                    Dim x As Double = grdNode.Item(1, i).Value
                    Dim y As Double = grdNode.Item(2, i).Value
                    g_MW.View.Draw.DrawPoint(x, y, 5, System.Drawing.Color.Yellow)
                    Exit Sub
                End If

            Next

        Catch ex As Exception

        End Try
    End Sub




    Private Sub cmbNetwork_SelectedValueChanged_1(sender As System.Object, e As System.EventArgs) Handles cmbNetwork.SelectedValueChanged
        CmbDiam.Items.Clear()
        CMBrough.Items.Clear()
        Dim pipeShp As New MapWinGIS.Shapefile
        Dim pipeLyr As Integer = getLayerHandle(cmbNetwork.Text)
        pipeShp = g_MW.Layers(pipeLyr).GetObject
        For i As Integer = 0 To pipeShp.NumFields - 1
            Dim fieldname As String = pipeShp.Field(i).Name
            CmbDiam.Items.Add(fieldname)
            CMBrough.Items.Add(fieldname)
        Next
    End Sub

    Private Sub CmbHydrant_SelectedValueChanged_1(sender As System.Object, e As System.EventArgs) Handles CmbHydrant.SelectedValueChanged
        CMBArea.Items.Clear()
        CMBDisch.Items.Clear()
        Dim hydShp As New MapWinGIS.Shapefile
        Dim hydLyr As Integer = getLayerHandle(CmbHydrant.Text)
        hydShp = g_MW.Layers(hydLyr).GetObject
        For i As Integer = 0 To hydShp.NumFields - 1
            Dim fieldname As String = hydShp.Field(i).Name
            CMBArea.Items.Add(fieldname)
            CMBDisch.Items.Add(fieldname)
        Next
    End Sub

    



    Private Sub SetSize()
        Dim loc As New Point(10, 10)
        zg1.Location = loc
        Zg2.Location = loc
        zg3.Location = loc
        ' Leave a small margin around the outside of the control
        Dim size As New Size(Me.ClientRectangle.Width - 20, Me.ClientRectangle.Height - 20)
        zg1.Size = size
        Zg2.Size = size
        zg3.Size = size

    End Sub









    Private Sub btnpath_Click(sender As System.Object, e As System.EventArgs) Handles btnpath.Click
        Dim Network As New MapWinGIS.Shapefile
        Network = g_MW.Layers(getLayerHandle(cmbNetwork.Text)).GetObject

        Dim TestPath As PathProperties
        Dim O, D As Integer
        Main_SP.Main()
        O = CInt(txtOrigin.Text)
        D = CInt(txtDestination.Text)
        Dim F, T As Integer
        clearGraphic()


        TestPath = createSPaths(FlowPathHyd, PipeNetwork, O, D)
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox7.Items.Clear()
        For i As Integer = TestPath.nPt To 1 Step -1
            ListBox1.Items.Add(TestPath.pNode(i).id)
            F = TestPath.pNode(i - 1).id
            T = TestPath.pNode(i).id

            If i > 1 Then
                Dim ff As Integer = HashLinkID(T & "-" & F)
                ListBox2.Items.Add(FormatNumber(Link(ff).diam, 0))
                ListBox7.Items.Add(FormatNumber(Link(ff).Discharge, 0))
            End If

            Dim ShapeID As Integer = getLinkId(F, T)
            If Not (Network.Shape(ShapeID - 1) Is Nothing) Then
                drawShape(Network.Shape(ShapeID - 1), Drawing.Color.Red)
            End If
        Next
    End Sub




    Private Sub DG_stat_SelectionChanged(sender As System.Object, e As System.EventArgs)
        DG_stat.ClearSelection()
    End Sub




    Sub check_pipes()  'Check that there are no pipes downstream larger than upstreams
        Dim Network As New MapWinGIS.Shapefile
        Network = g_MW.Layers(getLayerHandle(cmbNetwork.Text)).GetObject
        clearGraphic()

        Dim TestPath As PathProperties
        Dim O, D As Integer
        D = 1
        For iNode As Integer = 1 To nNode
            O = Node(iNode).id
            TestPath = createSPaths(FlowPathHyd, PipeNetwork, O, D)
            Dim chkdiam As Single = 0
            For ii As Integer = TestPath.nPt To 1 Step -1
                If ii > 1 Then
                    Dim ff As Integer = HashLinkID(TestPath.pNode(ii).id & "-" & TestPath.pNode(ii - 1).id)
                    If Link(ff).diam >= chkdiam Then
                        chkdiam = Link(ff).diam
                    Else
                        Link(ff).diam = chkdiam ' increase pipe diameter of the downstream pipe

                        drawShape(Network.Shape(ff - 1), Drawing.Color.Yellow)
                    End If
                End If
            Next
        Next

        For ilink As Integer = 0 To nLink - 1
            grdLink.Item(5, ilink).Value = Link(ilink + 1).diam
        Next
    End Sub





    Private Sub DG_StatVel_SelectionChanged(sender As System.Object, e As System.EventArgs)
        DG_StatVel.ClearSelection()
    End Sub

    Private Sub Button3_Click_1(sender As System.Object, e As System.EventArgs) Handles btn_checkpipe.Click
        check_pipes()
        MessageBox.Show("pipe diameters were checked")
    End Sub









    Private Sub Button6_Click_1(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        Dim folderbrowserdialog1 As New FolderBrowserDialog
        If (folderbrowserdialog1.ShowDialog() = DialogResult.OK) Then
            infolder = folderbrowserdialog1.SelectedPath
            TextBox1.Text = infolder
            cmdAutoGen.Visible = True
            cmdOpenFile.Visible = True
            GroupBox2.Visible = True
            GroupBox4.Visible = True
        End If
    End Sub

    Private Sub BtnAKLA_Click(sender As System.Object, e As System.EventArgs) Handles BtnAKLA.Click
        Dim num_nodes, i As Integer
        Dim num_links As Integer
        Dim t As Integer
        Dim tleft As Integer
        Dim value As Single
        Dim id As String
        Dim infile As String = infolder & "\Input.inp"

        TC5.TabPages.Remove(TPLoop)


        UpdateProgress.Minimum = 0
        UpdateProgress.Maximum = CInt(It_nb_txt.Text)
        UpdateProgress.Visible = True

        Dim f As Integer = 0
        Dim itnb As Integer = CInt(It_nb_txt.Text)
        For i = 0 To grdNode.Rows.Count - 1
            If grdNode.Item(3, i).Value = "Hydrant" Then
                f = f + 1
            End If
        Next



        
      
        Dim H(grdNode.Rows.Count, itnb) As Single 'pressure for each hydrant and each iteration
        Dim vel(nLink + npump, itnb) As Single 'Velocity for each section and each iteration
        Dim Hyd_nb As Integer = f

        Dim conffile As String = infolder & "\configuration.txt" ' Save configurationfile
        Using sw1 As StreamWriter = New StreamWriter(conffile)
            Dim Hfile As String = infolder & "\Pressure_head.txt" ' Save pressure file
            Dim Vfile As String = infolder & "\Velocity.txt" ' Save velocity file
            Using sw2 As StreamWriter = New StreamWriter(Hfile)
                Using sw3 As StreamWriter = New StreamWriter(Vfile)

                    For j As Integer = 1 To itnb


                        Application.DoEvents()
                        UpdateProgress.Value = j

                        'Create Inputfile
                        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                        'create_inputfile(infile)

                        Dim hydrantShp As New MapWinGIS.Shapefile
                        hydrantShp = g_MW.Layers(getLayerHandle(CmbHydrant.Text)).GetObject
                        clearGraphic()

                        Using sw As StreamWriter = New StreamWriter(infile)

                            ' Dim i As Integer
                            sw.WriteLine("[TITLE]")
                            sw.WriteLine("Title")
                            sw.WriteLine("")
                            sw.WriteLine("[JUNCTIONS]")
                            sw.WriteLine(";ID                  Elev            Demand          Pattern         ")

                            Dim hyd(nNode) As Integer
                            Dim hyd_id(nNode) As Integer
                            Dim hyd_open(nNode) As Integer ' 1 = open; 0 = closed
                            f = 0
                            Dim arrTest As New List(Of Integer)


                            For i = 1 To nNode
                                hyd_open(i) = 0

                                If Node(i).Type = "Hydrant" Then
                                    f = f + 1
                                    hyd(f) = Node(i).id ' hydrant id
                                    arrTest.Add(f)
                                    If j = 1 Then
                                        sw2.Write(hyd(f) & ";") 'write hydrant number on first row
                                    End If
                                End If
                            Next
                            If j = 1 Then
                                sw2.WriteLine()
                            End If

                            ReDim Preserve hyd(f)



                            Dim Qtot As Integer = 0

                            Dim Qcl As Single = CSng(txtQup.Text) ' upstream discharge



                            Do Until Qtot >= Qcl Or f = 0
                                Dim randomvalue As Integer
                                '  Dim rn As New Random
                                '   randomvalue = rn.Next(0, f - 1) ' random number between 0 and f-1
                                randomvalue = CInt(Math.Ceiling(Rnd() * f - 2)) + 1
                                Dim idh As Integer = arrTest(randomvalue)

                               Qtot += Node(hyd(idh)).Q
                                hyd_open(hyd(idh)) = 1
                                g_MW.View.Draw.DrawPoint(Node(hyd(idh)).x, Node(hyd(idh)).y, 4, System.Drawing.Color.Red)

                                arrTest.RemoveAt(randomvalue)
                                f -= 1
                            Loop


                            sw1.WriteLine()
                          

                            onode = 0 'open node

                            For i = 1 To nNode
                                If Node(i).Type = "Hydrant" Or Node(i).Type = "Node" Then
                                    If hyd_open(Node(i).id) = 1 Then 'open
                                        If chkb_emitter.Checked = True Then
                                            sw.WriteLine(Node(i).id & "   " & Node(i).z & "    " & Node(i).Q & "      ;")
                                        Else
                                            If Chkb_FR.Checked = False Then
                                                sw.WriteLine(Node(i).id & "   " & Node(i).z & "    " & Node(i).Q & "      ;")
                                            Else
                                                sw.WriteLine(Node(i).id & "   " & Node(i).z & "    " & 0 & "      ;") ' in case of hydrant CC
                                            End If
                                            onode += 1 'open hydrants
                                            ReDim Preserve Newnode(onode) ' z_NN(ONode)
                                            Newnode(onode).z = Node(i).z
                                            Newnode(onode).id = Node(i).id
                                            Newnode(onode).q = Node(i).Q
                                        End If
                                    Else 'closed
                                        sw.WriteLine(Node(i).id & "   " & Node(i).z & "    " & 0 & "      ;")
                                    End If
                                End If
                            Next

                            ' Hydrant CC
                            If onode <> 0 And Chkb_FR.Checked = True Then
                                Dim fff As Integer = nNode
                                For i = 1 To onode
                                    fff += 1
                                    Newnode(i).Fnode = fff
                                    sw.WriteLine(Newnode(i).Fnode & "   " & Newnode(i).z & "    " & 0 & "      ;")
                                    fff += 1
                                    Newnode(i).Snode = fff
                                    sw.WriteLine(Newnode(i).Snode & "   " & Newnode(i).z & "    " & 0 & "      ;")
                                Next
                            End If

                            sw.WriteLine("")

                            sw.WriteLine("[RESERVOIRS]")
                            sw.WriteLine(";ID                  Head            Pattern         ")
                            f = 0
                            For i = 1 To nNode
                                If Node(i).Type = "Reservoir" Then
    Dim Hi As Single = Node(i).z + CSng(txtHup.Text)
                                    sw.WriteLine(Node(i).id & "    " & Hi & "     ;")
                                End If
                            Next
                            If f = 0 Then
                                sw.WriteLine("")
                            End If

                            sw.WriteLine("[TANKS]")
                            sw.WriteLine(";ID                  Elevation       InitLevel       MinLevel        MaxLevel        Diameter        MinVol          VolCurve "" ")
                            sw.WriteLine("")

    'Random routine to extract the number of hydrants simultanuously opened
                            sw.WriteLine("")
                            sw.WriteLine("[PIPES]")
                            sw.WriteLine(";ID                 Node1               Node2               Length          Diameter        Roughness       MinorLoss       Status "" ")
                            Dim b1, b2, b3, b4 As Integer
                            Dim b5, b6, b7 As Single

                            For i = 1 To nLink
                                b1 = Link(i).id 'ID
                                b2 = Link(i).NodeF 'Node 1
                                b3 = Link(i).NodeT 'Node 2
                                b4 = Link(i).L 'length
                                b5 = Link(i).diam 'diameter
                                b6 = Link(i).rough 'Roughness
                                b7 = 0
                                sw.WriteLine(b1 & "  " & b2 & "  " & b3 & "     " & b4 & "  " & b5 & "  " & b6 & "  " & b7 & "   " & "  Open    ;")
                                If j = 1 Then ' write heading on velocity file
                                    If i = nLink Then
                                        sw3.WriteLine(FormatNumber(b1, 0))
                                    Else
                                        sw3.Write(FormatNumber(b1, 0) & ";")
                                    End If
                                End If
                            Next
                            ' Hydrant CC
                            If onode <> 0 And Chkb_FR.Checked = True Then
                                For i = 1 To onode
                                    b1 += 1
                                    sw.WriteLine(b1 & "  " & Newnode(i).id & "  " & Newnode(i).Fnode & "     " & 0.001 & "  " & b5 & "  " & b6 & "  " & b7 & "   " & "  CV    ;")
                                Next
                            End If

                          
                            sw.WriteLine()

                            sw.WriteLine("  ")
                            sw.WriteLine("[PUMPS]")
                            sw.WriteLine(";ID                  Node1               Node2               Parameters")
                            Try
                                If DG_pump.Rows.Count = 0 Then
                                    sw.WriteLine("  ")
                                Else
                                    Dim id1 As Integer = nNode
                                    For i = 1 To DG_pump.Rows.Count
                                        sw.WriteLine(";PUMP: PUMP:")
                                        id1 = nNode + 1
                                        Dim linkID As Integer = CInt(DG_pump.Item(1, i - 1).Value)
                                        Dim Node1 As Integer = Link(linkID).NodeF
                                        Dim Node2 As Integer = Link(linkID).NodeT
                                        Dim param As String = "HEAD " & DG_pump.Item(0, i - 1).Value
                                        sw.WriteLine(id1 & " " & Node1 & " " & Node2 & " " & param & ";")
                                    Next
                                End If

                            Catch ex As Exception
                                sw.WriteLine("  ")
                            End Try


                            sw.WriteLine("[VALVES]")
                            sw.WriteLine(";ID                  Node1               Node2               Diameter        Type    Setting         MinorLoss   ")
                            ' Hydrant CC
                            If onode <> 0 And Chkb_FR.Checked = True Then
                                For i = 1 To onode
                                    b1 += 1
                                    sw.WriteLine(b1 & "  " & Newnode(i).Fnode & "  " & Newnode(i).Snode & "     " & b5 & "  FCV  " & Newnode(onode).q & "  0    ;")
                                Next
                            End If

                            sw.WriteLine("  ")
                            sw.WriteLine("[TAGS]")
                            sw.WriteLine("  ")
                            sw.WriteLine("[DEMANDS]")
                            sw.WriteLine(";Junction            Demand          Pattern             Category")
                            sw.WriteLine("  ")
                            sw.WriteLine("[STATUS]")
                            sw.WriteLine(";ID                  Status/Setting")
                            sw.WriteLine("  ")
                            sw.WriteLine("[PATTERNS]")
                            sw.WriteLine(";ID                  Multipliers")
                            sw.WriteLine("  ")
                            sw.WriteLine("[CURVES]")
                            sw.WriteLine(";ID                  X-Value         Y-Value")
                            sw.WriteLine(";PUMP: ")
                            Try
                                If DG_pump.Rows.Count = 0 Then
                                    sw.WriteLine("  ")
                                Else
                                    For i = 1 To pumpcurve.GetUpperBound(0)
                                        sw.WriteLine(pumpcurve(i))
                                    Next
                                End If

                            Catch ex As Exception
                                sw.WriteLine("  ")
                            End Try

                            sw.WriteLine("[CONTROLS]")
                            sw.WriteLine("  ")
                            sw.WriteLine("[RULES]")
                            sw.WriteLine("  ")
                            sw.WriteLine("[ENERGY]")
                            sw.WriteLine(" Global Efficiency   75")
                            sw.WriteLine(" Global Price        0")
                            sw.WriteLine(" Demand Charge       0")
                            sw.WriteLine("  ")
                            sw.WriteLine("[EMITTERS]")
                            sw.WriteLine(";Junction            Coefficient")

                            If chkb_emitter.Checked = False Then 'emitter
                                Dim k, hdes, h0 As Single
                                Try
                                    hdes = CSng(Hdes_txt.Text)
                                    h0 = CSng(H0_txt.Text)
                                Catch ex As Exception
                                    MessageBox.Show("Missing values for Hdes and Hmin")
                                    Exit Sub
                                End Try


                                If onode <> 0 And Chkb_FR.Checked = True Then 'Flowe limiter
                                    For i = 1 To onode

                                        k = (Newnode(i).q) / ((hdes - h0) ^ 0.5)
                                        sw.WriteLine(Newnode(i).Snode & "     " & k)
                                    Next
                                End If

                                If onode <> 0 And Chkb_FR.Checked = False Then
                                    For i = 1 To onode
                                        k = (Newnode(i).q) / ((hdes - h0) ^ 0.5)
                                        sw.WriteLine(Newnode(i).id & "     " & k)
                                    Next
                                End If

                            End If
                            sw.WriteLine("  ")

                            sw.WriteLine("[OPTIONS]")
                            sw.WriteLine(" Units               LPS")
                            sw.WriteLine(" Headloss            D-W")
                            sw.WriteLine(" Specific Gravity    1")
                            sw.WriteLine(" Viscosity           1")
                            sw.WriteLine(" Trials              500")
                            sw.WriteLine(" Accuracy            0.001")
                            sw.WriteLine(" Unbalanced          Continue 10")
                            sw.WriteLine(" Pattern             1")

                            sw.WriteLine(" Demand Multiplier   ")
                            sw.WriteLine(" Emitter Exponent     ")
                          


                            sw.WriteLine("  ")
                            sw.WriteLine("[COORDINATES]")
                            sw.WriteLine(";Node                X-Coord             Y-Coord")
                            For i = 0 To grdNode.Rows.Count - 1
                                sw.WriteLine(grdNode.Item(0, i).Value & "   " & grdNode.Item(1, i).Value & "     " & grdNode.Item(2, i).Value)
                            Next
                            sw.WriteLine("  ")
                            sw.WriteLine("[VERTICES]")
                            sw.WriteLine(";Link                X-Coord             Y-Coord")
                            sw.WriteLine("  ")
                            sw.WriteLine("[LABELS]")
                            sw.WriteLine(";X-Coord           Y-Coord          Label & Anchor Node")
                            sw.WriteLine("  ")
                            sw.WriteLine("[BACKDROP]")
                            sw.WriteLine(" DIMENSIONS      0.00                0.00                10000.00            10000.00        ")
                            sw.WriteLine(" UNITS           None")
                            sw.WriteLine(" FILE            ")
                            sw.WriteLine(" OFFSET          0.00                0.00            ")
                            sw.WriteLine("")
                            sw.WriteLine("[END]")
                            sw.WriteLine("")
                        End Using

    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                        Call ENopen(infile, "", "")
                        Call ENgetcount(EN_NODECOUNT, num_nodes)
                        Call ENgetcount(EN_LINKCOUNT, num_links)
                        ENsolveH()
                        ENopenQ()
                        ENinitQ(0)

                        '      ReDim H1(num_nodes), Q1(num_nodes), V1(num_links), Qpipe(num_links)
                        If Chkb_FR.Checked = True Then
                            num_nodes = num_nodes - (4 * onode)
                            num_links = num_links - (4 * onode)
                        End If


                        Do
                            ENrunQ(t)
                            For i = 0 To num_nodes - 1 Step 1
                                ENgetnodevalue(i, EN_QUALITY, value)
                                id = "".PadRight(255, Chr(0))
                                ENgetnodeid(i, id)
                                id = id.Trim(Chr(0))
                                '  Console.WriteLine("Node " + id + " Quality " + Str(value))
                                ENgetnodevalue(i, EN_PRESSURE, value)


                                If grdNode.Item(3, i).Value = "Hydrant" Then
                                    sw2.Write(FormatNumber(Str(value), 0) & ";")
                                    H(i + 1, j) = FormatNumber(Str(value), 0)
                                End If
                            

                                ' ListBox3.Items.Add("Node " + id + " Pressure " + Str(value))
                            Next
                            sw2.WriteLine()
                            For i = 0 To num_links - 1 Step 1
                                ENgetlinkvalue(i + 1, EN_FLOW, value)

                                id = "".PadRight(255, Chr(0))
                                ENgetlinkid(i + 1, id)
                                id = id.Trim(Chr(0))
                                ' Console.WriteLine("Link " + id + " Flow " + Str(value))
                                '     ListBox4.Items.Add(i + 1 & ": " & Str(value))

                                ENgetlinkvalue(i + 1, EN_HEADLOSS, value)
                                '   ListBox5.Items.Add(Str(value))

                                ENgetlinkvalue(i + 1, EN_VELOCITY, value)
                                '   ListBox6.Items.Add(Link(i + 1).id & ": " & i & ": " & Str(value))

                                sw3.Write(FormatNumber(Str(value), 2) & ";") 'velocity
                                vel(i + 1, j) = Str(value)
                              

                            Next
                            sw3.WriteLine()
                            ENstepQ(tleft)
                            ' Stop
                        Loop Until (tleft = 0)
                        ENcloseQ()
                        ENsolveQ()
                        ENreport()

                        ENclose()
                    Next
                    UpdateProgress.Value = 0
                End Using 'Pressure head
            End Using 'Velocity

        End Using ' Configuration
        UpdateProgress.Visible = False

    '&&&&&&&&&&&&&&&&&&&&Plot&&&&&&&&&&&&&&&&&&&&&&& 

    Dim myPane As GraphPane = zg1.GraphPane
        myPane.Title.Text = ""
        myPane.XAxis.Title.Text = "Hydrant number"
        myPane.YAxis.Title.Text = "HPD (m)"
        myPane.CurveList.Clear()
    Dim list As New PointPairList
    Dim x As Double, y As Double
    Dim ii As Integer
    Dim hmin As Single = CSng(TxtHmin.Text)

    'Reliability Graph
    Dim myPane2 As GraphPane = Zg2.GraphPane
        myPane2.Title.Text = ""
        myPane2.XAxis.Title.Text = "Hydrant number"
        myPane2.YAxis.Title.Text = "Reliability %"
        myPane2.CurveList.Clear()
    Dim list2 As New PointPairList
    Dim x2 As Double, y2 As Double


    Dim ff As Integer = -1
        DGHPD.Rows.Clear()
        For ii = 0 To num_nodes - 1
    Dim hpd(itnb - 1) As Single
    Dim hyd As Boolean = False

    Dim sum_rel As Integer = 0
    Dim alpha As Single = 0
            If grdNode.Item(3, ii).Value = "Hydrant" Then
                For j As Integer = 1 To itnb
                    hyd = True
                    x = CInt(ii + 1)
                    y = (H(ii + 1, j) - hmin) / hmin
                    If H(ii + 1, j) >= hmin Then
                        sum_rel += 1
                    End If
                    hpd(j - 1) = y
                    list.Add(x, y)
                Next

                alpha = (sum_rel / itnb) * 100 'reliability
                x2 = CInt(ii + 1)
                y2 = alpha
                list2.Add(x2, y2)
            End If
            If hyd = True Then ' hydrant
                Array.Sort(hpd)
                DGHPD.Rows.Add()
                ff += 1
                DGHPD.Item(0, ff).Value = ii + 1 ' hydrant ID
                DGHPD.Item(1, ff).Value = FormatNumber(hpd(CInt(0.9 * (itnb - 1))), 2) ' 10% to be exceeded
                DGHPD.Item(2, ff).Value = FormatNumber(hpd(CInt(0.8 * (itnb - 1))), 2) ' 20% to be exceeded
                DGHPD.Item(3, ff).Value = FormatNumber(hpd(CInt(0.7 * (itnb - 1))), 2) ' 30% to be exceeded
                DGHPD.Item(4, ff).Value = FormatNumber(hpd(CInt(0.6 * (itnb - 1))), 2) ' 40% to be exceeded
                DGHPD.Item(5, ff).Value = FormatNumber(hpd(CInt(0.5 * (itnb - 1))), 2) '50% to be exceeded
                DGHPD.Item(6, ff).Value = FormatNumber(hpd(CInt(0.4 * (itnb - 1))), 2) ' 60% to be exceeded
                DGHPD.Item(7, ff).Value = FormatNumber(hpd(CInt(0.3 * (itnb - 1))), 2) ' 70% to be exceeded
                DGHPD.Item(8, ff).Value = FormatNumber(hpd(CInt(0.2 * (itnb - 1))), 2) ' 80% to be exceeded
                DGHPD.Item(9, ff).Value = FormatNumber(hpd(CInt(0.1 * (itnb - 1))), 2) ' 90% to be exceeded
                DGHPD.Item(10, ff).Value = FormatNumber(alpha, 0)
            End If
        Next
        DGVelocity.Rows.Clear()

        For ii = 0 To grdLink.Rows.Count - 1
    Dim v(itnb - 1) As Single
            ff = -1
            For j As Integer = 1 To itnb
                If FormatNumber(vel(ii + 1, j), 3) <> 0 Then
                    ff += 1
                    v(ff) = FormatNumber(vel(ii + 1, j), 3)
                End If
            Next
            ReDim Preserve v(ff)
            Array.Sort(v)
            DGVelocity.Rows.Add()
            DGVelocity.Item(0, ii).Value = Link(ii + 1).id ' hydrant ID

            If ff <> -1 Then

                DGVelocity.Item(1, ii).Value = FormatNumber(v(CInt(0.9 * (ff))), 2) ' 10% to be exceeded
                DGVelocity.Item(2, ii).Value = FormatNumber(v(CInt(0.8 * (ff))), 2) ' 20% to be exceeded
                DGVelocity.Item(3, ii).Value = FormatNumber(v(CInt(0.7 * (ff))), 2) ' 30% to be exceeded
                DGVelocity.Item(4, ii).Value = FormatNumber(v(CInt(0.6 * (ff))), 2) ' 40% to be exceeded
                DGVelocity.Item(5, ii).Value = FormatNumber(v(CInt(0.5 * (ff))), 2) '50% to be exceeded
                DGVelocity.Item(6, ii).Value = FormatNumber(v(CInt(0.4 * (ff))), 2) ' 60% to be exceeded
                DGVelocity.Item(7, ii).Value = FormatNumber(v(CInt(0.3 * (ff))), 2) ' 70% to be exceeded
                DGVelocity.Item(8, ii).Value = FormatNumber(v(CInt(0.2 * (ff))), 2) ' 80% to be exceeded
                DGVelocity.Item(9, ii).Value = FormatNumber(v(CInt(0.1 * (ff))), 2) ' 90% to be exceeded
                DGVelocity.Item(10, ii).Value = FormatNumber(StandardDeviation(v), 2)
            End If
        Next

    Dim myCurve As LineItem
        myCurve = myPane.AddCurve("", list, System.Drawing.Color.Blue, SymbolType.Diamond)
        myCurve.Symbol.Fill = New Fill(System.Drawing.Color.White)
        myCurve.Line.IsVisible = False
        myPane.XAxis.MajorGrid.IsVisible = True
        myPane.YAxis.MajorGrid.IsVisible = True
        myPane.XAxis.Color = System.Drawing.Color.Red
        myPane.YAxis.Scale.FontSpec.FontColor = System.Drawing.Color.Black
        myPane.YAxis.Title.FontSpec.FontColor = System.Drawing.Color.Black
        myPane.XAxis.Scale.Min = 0
        myPane.XAxis.Scale.Max = num_nodes + 1
        myPane.YAxis.Scale.Min = -1
        myPane.YAxis.Scale.Max = 4
        myPane.Chart.Fill = New Fill(System.Drawing.Color.White, System.Drawing.Color.LightGray, 45.0F)
        myPane.Legend.IsVisible = True
        myPane.Title.IsVisible = False
        zg1.IsShowContextMenu = True
        zg1.IsEnableHEdit = False
        zg1.IsEnableZoom = True
        zg1.IsShowHScrollBar = True
        zg1.IsShowVScrollBar = False
        zg1.IsAutoScrollRange = True
        zg1.IsEnableVZoom = False
        zg1.IsShowPointValues = True
        zg1.IsShowPointValues = True

    Dim myCurve2 As LineItem
        myCurve2 = myPane2.AddCurve("", list2, System.Drawing.Color.Red, SymbolType.Circle)
        myCurve2.Symbol.Fill = New Fill(System.Drawing.Color.White)
        myCurve2.Line.IsVisible = False
        myPane2.XAxis.MinorGrid.IsVisible = True
        myPane2.YAxis.MajorGrid.IsVisible = True
        myPane2.XAxis.Color = System.Drawing.Color.Red
        myPane2.YAxis.Scale.FontSpec.FontColor = System.Drawing.Color.Black
        myPane2.YAxis.Title.FontSpec.FontColor = System.Drawing.Color.Black
        myPane2.XAxis.Scale.Min = 0
        myPane2.XAxis.Scale.Max = num_nodes + 1
        myPane2.YAxis.Scale.Min = 0
        myPane2.YAxis.Scale.Max = 110

        myPane2.Chart.Fill = New Fill(System.Drawing.Color.White, System.Drawing.Color.LightGray, 45.0F)
        myPane2.Legend.IsVisible = True
        myPane2.Title.IsVisible = False
        Zg2.IsShowContextMenu = True
        Zg2.IsEnableHEdit = False
        Zg2.IsEnableZoom = True
        Zg2.IsShowHScrollBar = True
        Zg2.IsShowVScrollBar = False
        Zg2.IsAutoScrollRange = True
        Zg2.IsEnableVZoom = False
        Zg2.IsShowPointValues = True
        Zg2.IsShowPointValues = True

        SetSize()

        zg1.AxisChange()
        Zg2.AxisChange()

        TC5.TabPages.Add(TPLoop)

        MessageBox.Show("Done")

    'Catch ex As Exception
    '    MessageBox.Show("Error")
    'End Try
    End Sub

    Sub write_epanet_file(infile As String, Qcl As Single, Hup As Single)
        Dim hydrantShp As New MapWinGIS.Shapefile
        hydrantShp = g_MW.Layers(getLayerHandle(CmbHydrant.Text)).GetObject
        clearGraphic()
        Using sw As StreamWriter = New StreamWriter(infile)
            sw.WriteLine("[TITLE]")
            sw.WriteLine("Title")
            sw.WriteLine("")
            sw.WriteLine("[JUNCTIONS]")
            sw.WriteLine(";ID                  Elev            Demand          Pattern         ")

            Dim hyd(nNode) As Integer
            Dim hyd_id(nNode) As Integer
            Dim hyd_open(nNode) As Integer ' 1 = open; 0 = closed
            Dim f As Integer = 0
            Dim arrTest As New List(Of Integer)
            Dim i As Integer
            For i = 1 To nNode
                hyd_open(i) = 0

                If Node(i).Type = "Hydrant" Then
                    f = f + 1
                    hyd(f) = Node(i).id ' hydrant id
                    arrTest.Add(f)
                End If
            Next

            ReDim Preserve hyd(f)

            'generate random hydrants
            Dim Qtot As Integer = 0

            Do Until Qtot >= Qcl Or f = 0
                Dim randomvalue As Integer

                randomvalue = CInt(Math.Ceiling(Rnd() * f - 2)) + 1
                Dim idh As Integer = arrTest(randomvalue)

                Qtot += Node(hyd(idh)).Q
                hyd_open(hyd(idh)) = 1
                g_MW.View.Draw.DrawPoint(Node(hyd(idh)).x, Node(hyd(idh)).y, 4, System.Drawing.Color.Red)

                arrTest.RemoveAt(randomvalue)
                f -= 1
            Loop

            ONode = 0 'open node
            
            For i = 1 To nNode
                If Node(i).Type = "Hydrant" Or Node(i).Type = "Node" Then
                    If hyd_open(Node(i).id) = 1 Then 'open
                        If chkb_emitter.Checked = True Then
                            sw.WriteLine(Node(i).id & "   " & Node(i).z & "    " & Node(i).Q & "      ;")
                        Else
                            If Chkb_FR.Checked = False Then
                                sw.WriteLine(Node(i).id & "   " & Node(i).z & "    " & Node(i).Q & "      ;")
                            Else
                                sw.WriteLine(Node(i).id & "   " & Node(i).z & "    " & 0 & "      ;") ' in case of hydrant CC
                            End If
                            onode += 1 'open hydrants
                            ReDim Preserve Newnode(onode) ' z_NN(ONode)
                            Newnode(onode).z = Node(i).z
                            Newnode(onode).id = Node(i).id
                            Newnode(onode).q = Node(i).Q
                        End If
                    Else 'closed
                        sw.WriteLine(Node(i).id & "   " & Node(i).z & "    " & 0 & "      ;")
                    End If
                End If
            Next

            ' Hydrant CC
            If onode <> 0 And Chkb_FR.Checked = True Then
                Dim ff As Integer = nNode
                For i = 1 To onode
                    ff += 1
                    Newnode(i).Fnode = ff
                    sw.WriteLine(Newnode(i).Fnode & "   " & Newnode(i).z & "    " & 0 & "      ;")
                    ff += 1
                    Newnode(i).Snode = ff
                    sw.WriteLine(Newnode(i).Snode & "   " & Newnode(i).z & "    " & 0 & "      ;")
                Next
            End If



            sw.WriteLine("")

            sw.WriteLine("[RESERVOIRS]")
            sw.WriteLine(";ID                  Head            Pattern         ")
            f = 0
            For i = 1 To nNode
                If Node(i).Type = "Reservoir" Then
                    Dim Hi As Single = Node(i).z + Hup
                    sw.WriteLine(Node(i).id & "    " & Hi & "     ;")
                End If
            Next
            If f = 0 Then
                sw.WriteLine("")
            End If

            sw.WriteLine("[TANKS]")
            sw.WriteLine(";ID                  Elevation       InitLevel       MinLevel        MaxLevel        Diameter        MinVol          VolCurve "" ")
            sw.WriteLine("")

            'Random routine to extract the number of hydrants simultanuously opened
            sw.WriteLine("")
            sw.WriteLine("[PIPES]")
            sw.WriteLine(";ID                 Node1               Node2               Length          Diameter        Roughness       MinorLoss       Status "" ")
            Dim b1, b2, b3, b4 As Integer
            Dim b5, b6, b7 As Single
            For i = 1 To nLink
                b1 = Link(i).id 'ID
                b2 = Link(i).NodeF 'Node 1
                b3 = Link(i).NodeT 'Node 2
                b4 = Link(i).L 'length
                b5 = Link(i).diam 'diameter
                b6 = Link(i).rough 'Roughness
                b7 = 0
                sw.WriteLine(b1 & "  " & b2 & "  " & b3 & "     " & b4 & "  " & b5 & "  " & b6 & "  " & b7 & "   " & "  Open    ;")
            Next

            ' Hydrant CC
            If onode <> 0 And Chkb_FR.Checked = True Then
                For i = 1 To onode
                    b1 += 1
                    sw.WriteLine(b1 & "  " & Newnode(i).id & "  " & Newnode(i).Fnode & "     " & 0.001 & "  " & b5 & "  " & b6 & "  " & b7 & "   " & "  CV    ;")
                Next
            End If



            sw.WriteLine("  ")
            sw.WriteLine("[PUMPS]")
            sw.WriteLine(";ID                  Node1               Node2               Parameters")
            Try
                If npump = 0 Then
                    sw.WriteLine("  ")
                Else
                    Dim id As Integer = nNode
                    For i = 1 To DG_pump.Rows.Count
                        sw.WriteLine(";PUMP: PUMP:")
                        id = nNode + 1
                        Dim linkID As Integer = CInt(DG_pump.Item(1, i - 1).Value)
                        Dim Node1 As Integer = Link(linkID).NodeF
                        Dim Node2 As Integer = Link(linkID).NodeT
                        Dim param As String = "HEAD " & DG_pump.Item(0, i - 1).Value
                        sw.WriteLine(id & " " & Node1 & " " & Node2 & " " & param & ";")
                    Next
                End If

            Catch ex As Exception
                sw.WriteLine("  ")
            End Try
            sw.WriteLine("[VALVES]")
            sw.WriteLine(";ID                  Node1               Node2               Diameter        Type    Setting         MinorLoss   ")
            ' Hydrant CC
            If onode <> 0 And Chkb_FR.Checked = True Then
                For i = 1 To onode
                    b1 += 1
                    sw.WriteLine(b1 & "  " & Newnode(i).Fnode & "  " & Newnode(i).Snode & "     " & b5 & "  FCV  " & Newnode(onode).q & "  0    ;")
                Next
            End If

            sw.WriteLine("  ")


            sw.WriteLine("[TAGS]")
            sw.WriteLine("  ")
            sw.WriteLine("[DEMANDS]")
            sw.WriteLine(";Junction            Demand          Pattern             Category")
            sw.WriteLine("  ")
            sw.WriteLine("[STATUS]")
            sw.WriteLine(";ID                  Status/Setting")
            sw.WriteLine("  ")
            sw.WriteLine("[PATTERNS]")
            sw.WriteLine(";ID                  Multipliers")
            sw.WriteLine("  ")

            sw.WriteLine("[CURVES]")
            sw.WriteLine(";ID                  X-Value         Y-Value")
            'sw.WriteLine(";PUMP: ")
            Try
                If DG_pump.Rows.Count = 0 Then
                    sw.WriteLine("  ")
                Else
                    For i = 1 To pumpcurve.GetUpperBound(0)
                        sw.WriteLine(pumpcurve(i))
                    Next
                End If

            Catch ex As Exception
                sw.WriteLine("  ")
            End Try

            sw.WriteLine("[CONTROLS]")
            sw.WriteLine("  ")
            sw.WriteLine("[RULES]")
            sw.WriteLine("  ")
            sw.WriteLine("[ENERGY]")
            sw.WriteLine(" Global Efficiency   75")
            sw.WriteLine(" Global Price        0")
            sw.WriteLine(" Demand Charge       0")
            sw.WriteLine("  ")
            sw.WriteLine("[EMITTERS]")
            sw.WriteLine(";Junction            Coefficient")

            If chkb_emitter.Checked = False Then 'emitter
                Dim k, hdes, h0 As Single
                Try
                    hdes = CSng(Hdes_txt.Text)
                    h0 = CSng(H0_txt.Text)
                Catch ex As Exception
                    MessageBox.Show("Missing values for Hdes and Hmin")
                    Exit Sub
                End Try


                If onode <> 0 And Chkb_FR.Checked = True Then 'Flowe limiter
                    For i = 1 To onode

                        k = (Newnode(i).q) / ((hdes - h0) ^ 0.5)
                        sw.WriteLine(Newnode(i).Snode & "     " & k)
                    Next
                End If

                If onode <> 0 And Chkb_FR.Checked = False Then
                    For i = 1 To onode
                        k = (Newnode(i).q) / ((hdes - h0) ^ 0.5)
                        sw.WriteLine(Newnode(i).id & "     " & k)
                    Next
                End If

            End If
           







            sw.WriteLine("  ")

            sw.WriteLine("[OPTIONS]")
            sw.WriteLine(" Units               LPS")
            sw.WriteLine(" Headloss            D-W")
            sw.WriteLine(" Specific Gravity    1")
            sw.WriteLine(" Viscosity           1")
            sw.WriteLine(" Trials              1000")
            sw.WriteLine(" Accuracy            0.001")
            sw.WriteLine(" Unbalanced          Continue 10")
            sw.WriteLine(" Pattern             ")

            sw.WriteLine(" Demand Multiplier   ")
            sw.WriteLine(" Emitter Exponent     ")

            sw.WriteLine("  ")
            sw.WriteLine("[COORDINATES]")
            sw.WriteLine(";Node                X-Coord             Y-Coord")
            For i = 1 To nNode
                sw.WriteLine(Node(i).id & "   " & Node(i).x & "     " & Node(i).y)
            Next

            sw.WriteLine("  ")
            sw.WriteLine("[VERTICES]")
            sw.WriteLine(";Link                X-Coord             Y-Coord")
            sw.WriteLine("  ")
            sw.WriteLine("[LABELS]")
            sw.WriteLine(";X-Coord           Y-Coord          Label & Anchor Node")
            sw.WriteLine("  ")
            sw.WriteLine("[BACKDROP]")
            sw.WriteLine(" DIMENSIONS      0.00                0.00                10000.00            10000.00        ")
            sw.WriteLine(" UNITS           None")
            sw.WriteLine(" FILE            ")
            sw.WriteLine(" OFFSET          0.00                0.00            ")
            sw.WriteLine("")
            sw.WriteLine("[END]")
            sw.WriteLine("")
            ' Stop
        End Using
    End Sub ' Generate random hydrants to satisfy Qcl discharge

    Sub run_epanet(infile As String, ByRef H1() As Single, ByRef Q1() As Single, ByRef Qpipe() As Single, ByRef V1() As Single)

        Dim num_nodes, i As Integer
        Dim num_links As Integer
        Dim t As Integer
        Dim tleft As Integer
        Dim value As Single
        Dim id As String

        Call ENopen(infile, "", "")
        Call ENgetcount(EN_NODECOUNT, num_nodes)
        Call ENgetcount(EN_LINKCOUNT, num_links)
        ENsolveH()
        ENopenQ()
        ENinitQ(0)

        ReDim H1(num_nodes), Q1(num_nodes), V1(num_links), Qpipe(num_links)
        Do
            ENrunQ(t)
            For i = 0 To num_nodes - 1 Step 1
                ENgetnodevalue(i, EN_QUALITY, value)
                id = "".PadRight(255, Chr(0))
                ENgetnodeid(i, id)
                id = id.Trim(Chr(0))
                '  Console.WriteLine("Node " + id + " Quality " + Str(value))
                ENgetnodevalue(i, EN_PRESSURE, value)

                ' If grdNode.Item(3, i).Value = "Hydrant" Then
                '  sw2.Write(FormatNumber(Str(value), 0) & ";") ' pressure
                H1(i + 1) = FormatNumber(Str(value), 0)
                'test
                ENgetnodevalue(i, EN_DEMAND, value)
                Q1(i + 1) = Str(value)
                '  End If
                ' ListBox3.Items.Add("Node " + id + " Pressure " + Str(value))
            Next
            '
            For i = 0 To num_links - 1 Step 1
                ENgetlinkvalue(i + 1, EN_FLOW, value) ' Discharge

                id = "".PadRight(255, Chr(0))
                ENgetlinkid(i + 1, id)
                id = id.Trim(Chr(0))
                ' Console.WriteLine("Link " + id + " Flow " + Str(value))
                '    ListBox4.Items.Add(i + 1 & ": " & Str(value))
                Qpipe(i) = Str(value)

                ENgetlinkvalue(i + 1, EN_HEADLOSS, value) ' friction losses
                'ListBox5.Items.Add(Str(value))

                ENgetlinkvalue(i + 1, EN_VELOCITY, value) 'Velocity
                ' ListBox6.Items.Add(Link(i + 1).id & ": " & i & ": " & Str(value))
                '  sw3.Write(FormatNumber(Str(Value), 2) & ";") 'velocity
                V1(i + 1) = Str(value)
            Next
            '
            ENstepQ(tleft)
            ' Stop
        Loop Until (tleft = 0)
        Application.DoEvents()
        ENcloseQ()
        Application.DoEvents()
        ENsolveQ()
        ENreport()



        ENclose()
    End Sub

    Public Function StandardDeviation(NumericArray As Object) As Double

        Dim dblSum As Double, dblSumSqdDevs As Double, dblMean As Double
        Dim lngCount As Long, dblAnswer As Double
        Dim vElement As Object
        Dim lngStartPoint As Long, lngEndPoint As Long, lngCtr As Long


        lngCount = UBound(NumericArray)

        On Error Resume Next
        lngCount = 0

        'the check below will allow
        'for 0 or 1 based arrays.

        vElement = NumericArray(0)

        lngStartPoint = IIf(Err.Number = 0, 0, 1)
        lngEndPoint = UBound(NumericArray)

        'get sum and sample size
        For lngCtr = lngStartPoint To lngEndPoint
            vElement = NumericArray(lngCtr)
            If IsNumeric(vElement) Then
                lngCount = lngCount + 1
                dblSum = dblSum + CDbl(vElement)
            End If
        Next

        'get mean
        If lngCount > 1 Then
            dblMean = dblSum / lngCount

            'get sum of squared deviations
            For lngCtr = lngStartPoint To lngEndPoint
                vElement = NumericArray(lngCtr)

                If IsNumeric(vElement) Then
                    dblSumSqdDevs = dblSumSqdDevs + _
                    ((vElement - dblMean) ^ 2)
                End If
            Next

            'divide result by sample size - 1 and get square root. 
            'this function calculates standard deviation of a sample.  
            'If your  set of values represents the population, use sample
            'size not sample size - 1

            If lngCount > 1 Then
                lngCount = lngCount - 1 'eliminate for population values
                dblAnswer = Sqrt(dblSumSqdDevs / lngCount)
            End If

        End If

        StandardDeviation = dblAnswer

        Exit Function
    End Function


    Private Sub BtnPlotAkla_Click(sender As System.Object, e As System.EventArgs) Handles BtnPlotAkla.Click
        clearGraphic()
        Dim hydShp As New MapWinGIS.Shapefile
        Dim hydLyr As Integer = getLayerHandle(CmbHydrant.Text)
        hydShp = g_MW.Layers(hydLyr).GetObject

        Dim f As Integer = 0
        Select Case CBAkla.Text
            Case "HPD 10%"
                f = 1
            Case "HPD 20%"
                f = 2
            Case "HPD 30%"
                f = 3
            Case "HPD 40%"
                f = 4
            Case "HPD 50%"
                f = 5
            Case "HPD 60%"
                f = 6
            Case "HPD 70%"
                f = 7
            Case "HPD 80%"
                f = 8
            Case "HPD 90%"
                f = 9
            Case "Reliability (%)"
                f = 10
        End Select

        Dim j As Integer = 0
        Dim count(6) As Integer
        Array.Clear(count, 0, 6)
        If Chkb_FR.Checked = True Then
            nNode = nNode - (4 * onode)

        End If

        For i As Integer = 0 To nNode - 1
            If grdNode.Item(3, i).Value = "Hydrant" Then
                Dim id As Integer = grdNode.Item(0, i).Value
                Dim x As Double = grdNode.Item(1, i).Value
                Dim y As Double = grdNode.Item(2, i).Value
                Dim col As New System.Drawing.Color
                j += 1
                Dim HPDvalue As Single = DGHPD.Item(f, j - 1).Value

                If f <> 10 Then 'HPD
                    Select Case HPDvalue
                        Case Is <= -0.75
                            col = Red
                            count(1) += 1
                        Case Is <= -0.5
                            col = Yellow
                            count(2) += 1
                        Case Is <= -0.1
                            col = LightGreen
                            count(3) += 1
                        Case Is <= 0.1
                            col = Green
                            count(4) += 1
                        Case Is <= 0.5
                            col = LightBlue
                            count(5) += 1
                        Case Else
                            col = Blue
                            count(6) += 1
                    End Select
                Else
                    Select Case HPDvalue 'Reliability
                        Case Is <= 25
                            col = Red
                            count(1) += 1
                        Case Is <= 50
                            col = Yellow
                            count(2) += 1
                        Case Is <= 75
                            col = Green
                            count(3) += 1
                        Case Is <= 90
                            col = LightBlue
                            count(4) += 1
                        Case Else
                            col = Blue
                            count(5) += 1
                    End Select
                End If
                g_MW.View.Draw.DrawPoint(x, y, 5, col)
            End If
        Next
        Dim counttot As Integer = count(1) + count(2) + count(3) + count(4) + count(5) + count(6)

        DG_stat.Rows.Clear()
        If f = 10 Then
            For jj As Integer = 1 To 5
                DG_stat.Rows.Add()
                Dim dgvRow As DataGridViewRow = DG_stat.Rows(jj - 1)
                Select Case jj
                    Case 1
                        DG_stat.Item(0, jj - 1).Value = "0< <=25%"
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(1) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.Red
                    Case (2)
                        DG_stat.Item(0, jj - 1).Value = "25< <=50%"
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(2) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.Yellow
                    Case 3

                        DG_stat.Item(0, jj - 1).Value = "50< <=75%"
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(3) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.Green
                    Case 4
                        DG_stat.Item(0, jj - 1).Value = "75< <=90%"
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(4) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.LightBlue

                    Case 5
                        DG_stat.Item(0, jj - 1).Value = "90< <=100%"
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(5) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.Blue
                        dgvRow.DefaultCellStyle.ForeColor = Color.White
                End Select
            Next
        Else
            For jj As Integer = 1 To 6
                DG_stat.Rows.Add()
                Dim dgvRow As DataGridViewRow = DG_stat.Rows(jj - 1)
                Select Case jj
                    Case 1
                        DG_stat.Item(0, jj - 1).Value = "<= -0.75"
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(1) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.Red
                    Case 2
                        DG_stat.Item(0, jj - 1).Value = "-0.75< <=-0.5"
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(2) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.Yellow
                    Case 3

                        DG_stat.Item(0, jj - 1).Value = "-0.5< <=-0.1"
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(3) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.LightGreen
                    Case 4
                        DG_stat.Item(0, jj - 1).Value = "-0.1< <=0.1"
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(4) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.Green

                    Case 5
                        DG_stat.Item(0, jj - 1).Value = "0.1< <=0.5"
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(5) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.LightBlue
                    Case 6
                        DG_stat.Item(0, jj - 1).Value = "0.5 < "
                        DG_stat.Item(1, jj - 1).Value = FormatNumber(count(6) * 100 / counttot, 2)
                        dgvRow.DefaultCellStyle.BackColor = Color.Blue
                        dgvRow.DefaultCellStyle.ForeColor = Color.White
                End Select
            Next
        End If

    End Sub


    Private Sub Btn_vel_Click(sender As System.Object, e As System.EventArgs) Handles Btn_vel.Click
        clearGraphic()
        Dim pipeShp As New MapWinGIS.Shapefile
        Dim pipeLyr As Integer = getLayerHandle(cmbNetwork.Text)
        pipeShp = g_MW.Layers(pipeLyr).GetObject


        Dim f As Integer = 0

        Select Case CB_vel.Text
            Case "10%"
                f = 1
            Case "20%"
                f = 2
            Case "30%"
                f = 3
            Case "40%"
                f = 4
            Case "50%"
                f = 5
            Case "60%"
                f = 6
            Case "70%"
                f = 7
            Case "80%"
                f = 8
            Case "90%"
                f = 9
            Case Else
                f = 10
        End Select


        Dim j As Integer = 0
        Dim count(5) As Single

        Array.Clear(count, 0, 5)

        For i As Integer = 0 To grdLink.Rows.Count - 1

            Dim col As New System.Drawing.Color
            j += 1
            Dim Velvalue As Single
            Try
                Velvalue = DGVelocity.Item(f, j - 1).Value
            Catch ex As Exception
                GoTo 1
            End Try

            Select Case Velvalue
                Case Is <= 0.5
                    col = Blue
                    count(1) += Link(i + 1).L
                Case Is <= 1
                    col = Green
                    count(2) += Link(i + 1).L
                Case Is <= 1.5
                    col = Yellow
                    count(3) += Link(i + 1).L
                Case Is <= 2
                    col = Orange
                    count(4) += Link(i + 1).L
                Case Else
                    col = Red
                    count(5) += Link(i + 1).L
            End Select
            drawShape(pipeShp.Shape(i), col, 3)
1:      Next
        Dim counttot As Integer = count(1) + count(2) + count(3) + count(4) + count(5)

        DG_StatVel.Rows.Clear()

        For jj As Integer = 1 To 5
            DG_StatVel.Rows.Add()
            Dim dgvRow As DataGridViewRow = DG_StatVel.Rows(jj - 1)
            Select Case jj
                Case 1
                    DG_StatVel.Item(0, jj - 1).Value = "<= 0.5 m/s"
                    DG_StatVel.Item(1, jj - 1).Value = FormatNumber(count(1), 0)
                    DG_StatVel.Item(2, jj - 1).Value = FormatNumber(count(1) * 100 / counttot, 2)
                    dgvRow.DefaultCellStyle.BackColor = Color.Blue
                    dgvRow.DefaultCellStyle.ForeColor = Color.White
                Case 2
                    DG_StatVel.Item(0, jj - 1).Value = "0.5 < <=1 m/s"
                    DG_StatVel.Item(1, jj - 1).Value = FormatNumber(count(2), 0)
                    DG_StatVel.Item(2, jj - 1).Value = FormatNumber(count(2) * 100 / counttot, 2)
                    dgvRow.DefaultCellStyle.BackColor = Color.Green
                Case 3

                    DG_StatVel.Item(0, jj - 1).Value = "1< <=1.5 m/s"
                    DG_StatVel.Item(1, jj - 1).Value = FormatNumber(count(3), 0)
                    DG_StatVel.Item(2, jj - 1).Value = FormatNumber(count(3) * 100 / counttot, 2)
                    dgvRow.DefaultCellStyle.BackColor = Color.Yellow
                Case 4
                    DG_StatVel.Item(0, jj - 1).Value = "1.5< <=2 m/s"
                    DG_StatVel.Item(1, jj - 1).Value = FormatNumber(count(4), 0)
                    DG_StatVel.Item(2, jj - 1).Value = FormatNumber(count(4) * 100 / counttot, 2)
                    dgvRow.DefaultCellStyle.BackColor = Color.Orange

                Case 5
                    DG_StatVel.Item(0, jj - 1).Value = "> 2 m/s"
                    DG_StatVel.Item(1, jj - 1).Value = FormatNumber(count(5), 0)
                    DG_StatVel.Item(2, jj - 1).Value = FormatNumber(count(5) * 100 / counttot, 2)
                    dgvRow.DefaultCellStyle.BackColor = Color.Red
            End Select
        Next

    End Sub



    Function near_diam_row(diam As Single) As Integer
        Dim id As Integer = 0
        For i As Integer = 0 To DG_diam.Rows.Count - 1
            If diam >= DG_diam.Item(0, i).Value Then
                id = i + 1
            End If
        Next
        Return id
    End Function

    Sub pipe_cost()
        Dim ndiam As Integer = DG_cost.Rows.Count
        Dim Diam_L(ndiam), diam_cost(ndiam) As Single 'cost and diameter for each diameter
        Array.Clear(Diam_L, 0, ndiam)
        Array.Clear(diam_cost, 0, ndiam)
        DG_error.Visible = False
        Label4.Visible = False
        DG_error.Rows.Clear()
        Dim f As Integer = -1

        Dim cost_tot, length_tot As Single
        cost_tot = 0
        length_tot = 0

        For j As Integer = 1 To nLink ' Calculate pipe cost
            length_tot += Link(j).L
            Dim chk As Boolean = False
            For i As Integer = 1 To ndiam
                If DG_cost.Item(0, i - 1).Value = Link(j).diam Then
                    Diam_L(i) += Link(j).L
                    diam_cost(i) += Link(j).L * DG_diam.Item(1, i - 1).Value
                    chk = True
                    Exit For
                End If
                If i = ndiam And chk = False Then
                    DG_error.Visible = True
                    Label4.Visible = True
                    'Can't find diameter
                    DG_error.Rows.Add()
                    f += 1
                    DG_error.Item(0, f).Value = Link(j).id
                    DG_error.Item(1, f).Value = Link(j).diam
                End If
            Next
        Next


        For i As Integer = 1 To ndiam
            DG_cost.Item(1, i - 1).Value = FormatNumber(Diam_L(i), 0, TriState.False)
            DG_cost.Item(2, i - 1).Value = FormatNumber(diam_cost(i), 0, TriState.False)
            cost_tot += diam_cost(i)
        Next
        txtLength.Text = FormatNumber(length_tot, 0, TriState.False)
        txtcost.Text = FormatNumber(cost_tot, 0, TriState.False)

    End Sub
    Private Sub Btn_Open_diam_Click(sender As System.Object, e As System.EventArgs) Handles Btn_Open_diam.Click
        Dim OFD As New OpenFileDialog
        OFD.Filter = "iPipe. Files (*.Pipe)|*.Pipe"

        Dim pipefile As String
        If OFD.ShowDialog() = Windows.Forms.DialogResult.OK Then
            pipefile = OFD.FileName
            open_pipefile(pipefile)
            pipe_cost()
        End If
    End Sub

    Private Sub Btn_save_diam_Click(sender As System.Object, e As System.EventArgs) Handles Btn_save_diam.Click
        Dim SFD As New SaveFileDialog
        SFD.Filter = "Pipe. Files (*.Pipe)|*.Pipe"

        Dim pipefile As String
        If SFD.ShowDialog() = Windows.Forms.DialogResult.OK Then
            pipefile = SFD.FileName
            save_pipefile(pipefile)
        End If
    End Sub




    Private Sub NUD_Q_ValueChanged(sender As System.Object, e As System.EventArgs) Handles NUD_Q.ValueChanged
        Dim f As Integer = NUD_Q.Value
        If f > DG_multiQ.Rows.Count Then
            DG_multiQ.Rows.Add()
            DG_percentile.Rows.Add()
            Exit Sub
        End If
        If f < DG_multiQ.Rows.Count Then
            DG_multiQ.Rows.RemoveAt(DG_multiQ.Rows.Count - 1)
            DG_percentile.Rows.RemoveAt(DG_multiQ.Rows.Count - 1)
            Exit Sub
        End If

    End Sub


    Private Sub Btn_CC_Click(sender As System.Object, e As System.EventArgs) Handles Btn_CC.Click
        Dim infile As String = infolder & "\Input.inp"
        Dim Hup As Single = CSng(txt_Hup.Text)
        Dim itnb As Integer = CInt(txt_itnb.Text)
        Dim hmin As Single = CSng(txthmin1.Text)
        Dim disch_nb As Integer = DG_multiQ.Rows.Count
        UpdateProgress.Minimum = 0
        UpdateProgress.Maximum = itnb * disch_nb
        UpdateProgress.Visible = True
        Dim f As Integer = 0
        Dim hup1(itnb - 1) As Single
        Try

       
        TC5.TabPages.Remove(TPLoop)
        Dim myPane As GraphPane = ZedGraphControl1.GraphPane
        myPane.CurveList.Clear()

        Dim CCfile As String = infolder & "\CC.txt" ' Save pressure file
        Using sw As StreamWriter = New StreamWriter(CCfile)


            For i As Integer = 0 To DG_multiQ.Rows.Count - 1
                Dim Qup As Single = CSng(DG_multiQ.Item(0, i).Value)
                sw.WriteLine(Qup)
                For j As Integer = 1 To itnb
                    f += 1
                    UpdateProgress.Value = f
                    Application.DoEvents()
                    Dim h(nNode), Q(nLink + npump) As Single
                    Dim V(nLink + npump), Qp(nLink + npump) As Single
                    Array.Clear(h, 0, nNode)
                    write_epanet_file(infile, Qup, Hup)
                    run_epanet(infile, h, Q, Qp, V)
                    Dim hyd_min As Single = 9999

                    If onode <> 0 And Chkb_FR.Checked = True Then
                        For ii As Integer = 1 To onode
                            Q(Newnode(ii).id) = Q(Newnode(ii).Snode)
                        Next
                    End If

                    For jj As Integer = 1 To nNode

                        If h(jj) <> 0 And h(jj) < hyd_min Then
                            hyd_min = h(jj)
                        End If
                    Next
                    Select Case hyd_min
                        Case 9999
                            GoTo 1
                            'MsgBox("error")
                            'Exit Sub
                        Case Is > hmin
                            hup1(j - 1) = Hup - (hyd_min - hmin)
                        Case Is < hmin
                            hup1(j - 1) = Hup + (hmin - hyd_min)
                    End Select
                    If j = itnb Then
                        sw.WriteLine(hup1(j - 1))
                    Else
                        sw.Write(hup1(j - 1) & ";")
                    End If
1:              Next

                Array.Sort(hup1)

                Dim list As New PointPairList()
                DG_percentile.Item(0, i).Value = Qup
                For ii As Integer = 0 To 10
                    Dim y As Double = CDbl(Qup)
                    Dim x As Double = hup1((itnb - 1) * ii / 10)
                    list.Add(x, y)
                    DG_percentile.Item(ii + 1, i).Value = x
                Next
                Dim myCurve As LineItem = myPane.AddCurve("", list, Color.Blue, SymbolType.Diamond)
                myCurve.Line.IsVisible = False
                ' Horizontal box and whisker chart
                ' yval is the vertical position of the box & whisker
                Dim yval As Double = CDbl(Qup)
                ' pct5 = 5th percentile value
                Dim pct5 As Double = hup1(CInt((itnb - 1) * 0.05))
                ' pct25 = 25th percentile value
                Dim pct25 As Double = hup1(CInt((itnb - 1) * 0.25))
                ' median = median value
                Dim median As Double = hup1(CInt((itnb - 1) * 0.5))
                ' pct75 = 75th percentile value
                Dim pct75 As Double = hup1(CInt((itnb - 1) * 0.75))
                ' pct95 = 95th percentile value
                Dim pct95 As Double = hup1(CInt((itnb - 1) * 0.95))
                ' Draw the box
                Dim list2 As New PointPairList()
                list2.Add(pct25, yval, median)
                list2.Add(median, yval, pct75)
                Dim myBar As HiLowBarItem = myPane.AddHiLowBar("", list2, Color.Black)
                ' set the size of the box (in points, scaled to graph size)
                myBar.Bar.Size = 20
                myBar.Bar.Fill.IsVisible = False
                myPane.BarSettings.Base = BarBase.Y

                ' Draw the whiskers
                Dim xwhisk As Double() = {pct5, pct25, PointPair.Missing, pct75, pct95}
                Dim ywhisk As Double() = {yval, yval, yval, yval, yval}
                Dim list3 As New PointPairList()
                list3.Add(xwhisk, ywhisk)
                Dim mywhisk As LineItem = myPane.AddCurve("", list3, Color.Black, SymbolType.None)
                myPane.XAxis.Title.Text = "Pressure head (m)"
                myPane.YAxis.Title.Text = "Discharge (l/s)"
                myPane.Title.IsVisible = False

                ZedGraphControl1.AxisChange()

                ZedGraphControl1.RestoreScale(myPane)

            Next
        End Using
            UpdateProgress.Visible = False
        Catch ex As Exception

        End Try


    End Sub








    Private Sub txtOrigin_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtOrigin.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""

    End Sub


    Private Sub txtDestination_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtDestination.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = "" 'allow only integers

    End Sub

    Private Sub qs_txt_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs)
        Dim int As New Interfaces
        int.double_keypress(txtHup.Text, sender, e)
    End Sub


    Private Sub T_txt_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs)
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = "" 'allow only integers
    End Sub

    Private Sub T_txt_Validated(sender As System.Object, e As System.EventArgs)
        If T_txt.TextLength > 0 Then     'Validate that the textbox is between 0 and 24
            Dim number = Double.Parse(T_txt.Text)
            If number < 0.0 Then
                number = 0.0
            ElseIf number > 24 Then
                number = 24
            End If
            T_txt.Text = number.ToString()
        End If
    End Sub

    Private Sub MinHyd_txt_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs)
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = "" 'allow only integers
    End Sub





    Private Sub DG_diam_EditingControlShowing(sender As System.Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DG_diam.EditingControlShowing
        '---restrict inputs on the fourth field---
        If DG_diam.CurrentCell.ColumnIndex <= 2 And Not e.Control Is Nothing Then
            Dim tb As TextBox = CType(e.Control, TextBox)

            '---add an event handler to the TextBox control---
            Dim int As New Interfaces
            AddHandler tb.KeyPress, AddressOf int.TextBox_KeyPress
        End If


    End Sub



    Private Sub Nup_diam_ValueChanged(sender As System.Object, e As System.EventArgs) Handles Nup_diam.ValueChanged
        Dim i As Integer = Nup_diam.Value
        For j As Integer = 0 To i
            Select Case Nup_diam.Value
                Case Is > DG_diam.Rows.Count
                    DG_diam.Rows.Add()
                Case Is < DG_diam.Rows.Count
                    DG_diam.Rows.RemoveAt(DG_diam.Rows.GetLastRow(DataGridViewElementStates.None))
                Case Is = DG_diam.Rows.Count
                    Exit Sub
            End Select
        Next
    End Sub

    Private Sub It_nb_txt_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles It_nb_txt.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = "" 'allow only integers
    End Sub

    Private Sub txtHup_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtHup.KeyPress
        Dim int As New Interfaces
        int.double_keypress(txtHup.Text, sender, e)
    End Sub


    Private Sub txtQup_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtQup.KeyPress
        Dim int As New Interfaces
        int.double_keypress(txtHup.Text, sender, e)
    End Sub

    Private Sub TxtHmin_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtHmin.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = "" 'allow only integers
    End Sub

    Private Sub txt_itnb_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_itnb.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = "" 'allow only integers
    End Sub

    Private Sub txt_Hup_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Hup.KeyPress
        Dim int As New Interfaces
        int.double_keypress(txtHup.Text, sender, e)
    End Sub

    Private Sub txthmin1_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txthmin1.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = "" 'allow only integers
    End Sub


    Private Sub GroupBox4_Enter(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub CB_disch_SelectedIndexChanged(sender As System.Object, e As System.EventArgs)
        If CB_disch.Text = "Mixed Reservoirs" Then
            CB_Cl_Res.Visible = False
            lbl_res.Visible = False
        Else
            CB_Cl_Res.Visible = True
            lbl_res.Visible = True
        End If
    End Sub





    Private Sub CB_design_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CB_design.SelectedIndexChanged
        Select Case CB_design.Text
            Case "Pareto Front"
                GroupBox5.Visible = True
                TabControl3.TabPages.Insert(2, TabPage2)
                BtnDesign_Cl.Visible = False
                Btn_design_pareto.Visible = True
                Btn_SFR.Visible = False
                GroupBox6.Visible = False
            Case "Q Clement"
                GroupBox5.Visible = False
                TabControl3.TabPages.Remove(TabPage2)
                BtnDesign_Cl.Visible = True
                Btn_design_pareto.Visible = False
                GroupBox6.Visible = False
                Btn_SFR.Visible = False
            Case "Several Flow Regime"
                GroupBox5.Visible = False
                TabControl3.TabPages.Remove(TabPage2)
                BtnDesign_Cl.Visible = False
                Btn_design_pareto.Visible = False
                GroupBox6.Visible = True
                Btn_SFR.Visible = True
        End Select
    End Sub

    Private Sub BtnDesign_Click(sender As System.Object, e As System.EventArgs) Handles BtnDesign_Cl.Click
        Try
            Label8.Visible = True
            Label9.Visible = True
            Label10.Visible = True
            picGraph.Visible = True
            Dim hup As Single = CSng(Hpiez_txt.Text)
            Dim Hmin As Single = CSng(Hmin_txt.Text)
            Dim ndiam As Integer = DG_diam.Rows.Count
            ReDim Pipe_C(ndiam) ' read pipe size/cost grid
            Dim maxdiam As Single

            Dim RiD As Integer

            RiD = CInt(CB_res_design.Text)



            If nres = 0 Then
                MsgBox("Error! there is no reservoir in your network")
                Exit Sub
            End If
            For i As Integer = 1 To ndiam ' assign initial diameters

                Pipe_C(i).diam = DG_diam.Item(0, i - 1).Value
                Pipe_C(i).price = DG_diam.Item(1, i - 1).Value
                Pipe_C(i).rough = DG_diam.Item(2, i - 1).Value
                If i = ndiam Then
                    maxdiam = DG_diam.Item(0, i - 1).Value
                End If
            Next

            Dim Beta(nLink) As Single
            Dim HF As New LIDM
            Dim rough(nLink), diam(nLink), diam1(nLink), rough1(nLink), cost(nLink), cost1(nLink) As Single
            Dim Qcl(nLink) As Single

            For j As Integer = 1 To nLink
                Dim id As Integer = near_diam_row(Math.Sqrt((4 * Link(j).Discharge) * 1000 / (Math.PI * 2.5)))
                If id = ndiam Then
                    id = ndiam - 1
                End If
                rough(j) = Pipe_C(id + 1).rough
                diam(j) = Pipe_C(id + 1).diam 'smallest diameters
                cost(j) = Pipe_C(id + 1).price
                Qcl(j) = grdLink.Item(4, j - 1).Value
            Next
            Dim h As Single
            Dim hi As Integer = 0
            Dim ymax As Single


            For i As Integer = 1 To 10000 ' iteration
                Dim Hpiez As Single
                Application.DoEvents()
                Hpiez = solve_net(RiD, Qcl, diam, rough, cost, diam1, rough1, cost1)
                If i = 1 Or i = 235 Or i = 470 Or i = 705 Or i = 940 Then
                    ymax = Hpiez
                    Label9.Text = FormatNumber(Hpiez, 0) & " m"
                    If Not (picGraph.Image Is Nothing) Then
                        picGraph.Image.Dispose()
                    End If
                End If

                If h = Hpiez Then
                    hi += 1
                    If hi = 2 Then
                        MessageBox.Show("You have to increase the upstream pressure to find solution")
                        Exit Sub
                    End If
                Else
                    h = Hpiez
                End If
                Label8.Text = FormatNumber(Hpiez, 0) & " m"
                Label10.Text = " Iteration: " & i
                If Hpiez < CSng(Hpiez_txt.Text) Then
                    'Check that there are no pipes downstream larger than upstreams
                    check_pipes()
                    pipe_cost()
                    MessageBox.Show("solution was found")
                    Exit For
                End If
                PlotValue(h, Hpiez, ymax)

            Next

            UpdateProgress.Value = 0
            UpdateProgress.Visible = False

            Label9.Visible = False
            Label10.Visible = False

            picGraph.Visible = False
        Catch ex As Exception
            MessageBox.Show("Error, check that your project is loaded")
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Btn_design_pareto.Click
        ' Try


        Dim design(CInt(txt_Dnb.Text) - 1, nLink) As Integer
        Dim cost_D(CInt(txt_Dnb.Text) - 1) As Integer
        Dim Rel_D(CInt(txt_Dnb.Text) - 1) As Single
        Dim I_D(CInt(txt_Dnb.Text) - 1) As Single
        ' Dim rel_levl As Single = CInt(CBRel.Text) / 100
        Label8.Visible = True
        Label10.Visible = True
        Label9.Visible = True
        Dim myPane As GraphPane = zg3.GraphPane
        myPane.Title.Text = ""
        myPane.XAxis.Title.Text = "Cost"
        myPane.YAxis.Title.Text = "Reliability"
        myPane.CurveList.Clear()
        Dim list As New PointPairList
        Dim list1 As New PointPairList
        Dim myCurve As LineItem
        Dim myCurve1 As LineItem

        Dim Des_file As String = infolder & "\Pareto_front.des"


        Using sw As StreamWriter = New StreamWriter(Des_file)

            sw.WriteLine("cost ; Reliability; Beta_Parameter")


            For jj As Integer = 1 To CInt(txt_Dnb.Text)

                Dim hup As Single = CSng(Hpiez_txt.Text)
                Dim Hmin As Single = CSng(Hmin_txt.Text)
                Dim Qup As Single = CSng(txt_Qup.Text)

                Dim ndiam As Integer = DG_diam.Rows.Count
                ReDim Pipe_C(ndiam) ' read pipe size/cost grid
                Dim maxdiam As Single

                Dim RiD As Integer = CInt(CB_res_design.Text)

                If nres = 0 Then
                    MsgBox("Error! there is no reservoir in your network")
                    Exit Sub
                End If
                For i As Integer = 1 To ndiam
                    Pipe_C(i).diam = DG_diam.Item(0, i - 1).Value
                    Pipe_C(i).price = DG_diam.Item(1, i - 1).Value
                    Pipe_C(i).rough = DG_diam.Item(2, i - 1).Value
                    If i = ndiam Then
                        maxdiam = DG_diam.Item(0, i - 1).Value
                    End If
                Next

                Dim Beta(nLink) As Single
                Dim HF As New LIDM
                Dim rough(nLink), diam(nLink), diam1(nLink), rough1(nLink), cost(nLink), cost1(nLink) As Single




                Dim h As Single
                Dim hi As Integer = 0
                Dim ymax As Single
                Dim infile As String = infolder & "\Input.inp"

                Dim hs(nNode), Qs(nNode) As Single
                Dim V(nLink + npump), Qcli(nLink + npump) As Single
                Array.Clear(hs, 0, nNode)

                write_epanet_file(infile, Qup, hup)
                run_epanet(infile, hs, Qs, Qcli, V)

                For j As Integer = 1 To nLink
                    Dim id As Integer = near_diam_row(Math.Sqrt((4 * Qcli(j) * 1000) / (Math.PI * 2.5)))
                    rough(j) = Pipe_C(id + 1).rough
                    diam(j) = Pipe_C(id + 1).diam 'smallest diameters
                    cost(j) = Pipe_C(id + 1).price
                Next


                ReDim Preserve Qcli(nLink)
                For i As Integer = 1 To 10000 ' iteration
                    Dim Hpiez As Single
                    Application.DoEvents()
                    Hpiez = solve_net(RiD, Qcli, diam, rough, cost, diam1, rough1, cost1)

                    If i = 1 Or i = 235 Or i = 470 Or i = 705 Or i = 940 Then
                        ymax = Hpiez
                        Label9.Text = FormatNumber(Hpiez, 0) & " m"
                        If Not (picGraph.Image Is Nothing) Then
                            picGraph.Image.Dispose()
                        End If
                    End If

                    If h = Hpiez Then
                        hi += 1
                        If hi > 85 Then
                            MessageBox.Show("You have to increase the upstream pressure to find solution")
                            Exit Sub
                        End If
                    Else
                        h = Hpiez
                    End If
                    Label8.Text = FormatNumber(Hpiez, 0) & " m"
                    Label10.Text = "Design :" & jj & vbNewLine & "Iteration: " & i
                    If Hpiez < CSng(Hpiez_txt.Text) Then
                        'Check that there are no pipes downstream larger than upstreams
                        check_pipes()
                        pipe_cost()
                        For jjj As Integer = 1 To nLink
                            design(jj - 1, jjj) = grdLink.Item(5, jjj - 1).Value
                        Next
                        cost_D(jj - 1) = CInt(txtcost.Text)

                        Dim rel_tot As Single = 0
                        ' MessageBox.Show("solution was found")
                        Dim rel_y(CInt(txt_itD.Text) - 1) As Single
                        For j As Integer = 1 To CInt(txt_itD.Text)
                            Application.DoEvents()
                            Label9.Text = "Configuration " & j
                            Dim h1(nNode), Q1(nLink + npump) As Single
                            Dim V1(nLink + npump), QP(nLink + npump) As Single
                            Array.Clear(h1, 0, nNode)
                            write_epanet_file(infile, Qup, hup)
                            '  Application.DoEvents()
                            run_epanet(infile, h1, Q1, QP, V1)
                            Dim sum_rel As Integer = 0
                            Dim f As Integer = 0

                            For jjj As Integer = 1 To nNode
                                If Node(jjj).Type = "Hydrant" Then
                                    f += 1
                                    If h1(jjj) >= Hmin Then
                                        sum_rel += 1
                                    End If
                                End If
                            Next
                            rel_y(j - 1) = sum_rel * 100 / f
                            ' list.Add(cost_D(j - 1), rel_y(j - 1))
                            rel_tot += rel_y(j - 1)
                        Next
                        rel_tot = rel_tot / CInt(txt_itD.Text) ' average reliability of all configurations
                        '     Array.Sort(rel_y, 0, CInt(txt_itD.Text))

                        Rel_D(jj - 1) = rel_tot ' rel_y(CInt(txt_itD.Text) - 1)
                        list.Add(CInt(cost_D(jj - 1)), CInt(Rel_D(jj - 1)))

                        I_D(jj - 1) = Rel_D(jj - 1) / cost_D(jj - 1)
                        sw.WriteLine(CInt(cost_D(jj - 1)) & ";" & CInt(Rel_D(jj - 1)) & ";" & I_D(jj - 1))
                        Exit For
                    End If
                Next
            Next
            Dim idi As Integer = 0
            For jj As Integer = 0 To CInt(txt_Dnb.Text) - 1
                If jj = 0 Then
                    idi = jj
                Else
                    If I_D(idi) < I_D(jj) Then
                        idi = jj
                    End If
                End If
            Next

            list1.Add(CInt(cost_D(idi)), CInt(Rel_D(idi)))
            sw.WriteLine()
            sw.WriteLine("====== ; =============; ==================")
            sw.WriteLine(CInt(cost_D(idi)) & ";" & CInt(Rel_D(idi)) & ";" & I_D(idi))

            For jj As Integer = 1 To grdLink.Rows.Count
                grdLink.Item(5, jj - 1).Value = design(idi, jj)
                Link(jj).diam = design(idi, jj)
            Next
            pipe_cost()

            myCurve1 = myPane.AddCurve("", list1, System.Drawing.Color.Red, SymbolType.Diamond)
            myCurve1.Symbol.Fill = New Fill(System.Drawing.Color.Red)
            myCurve = myPane.AddCurve("", list, System.Drawing.Color.Blue, SymbolType.Diamond)
            myCurve.Symbol.Fill = New Fill(System.Drawing.Color.Blue)


            myCurve.Line.IsVisible = False
            myCurve1.Line.IsVisible = False

            myPane.XAxis.MajorGrid.IsVisible = True
            myPane.YAxis.MajorGrid.IsVisible = True
            myPane.XAxis.Color = System.Drawing.Color.Red
            myPane.YAxis.Scale.FontSpec.FontColor = System.Drawing.Color.Black
            myPane.YAxis.Title.FontSpec.FontColor = System.Drawing.Color.Black
            Array.Sort(Rel_D)
            myPane.YAxis.Scale.Min = Rel_D(0) - 10
            myPane.YAxis.Scale.Max = 110
            myPane.Chart.Fill = New Fill(System.Drawing.Color.White, System.Drawing.Color.LightGray, 45.0F)
            myPane.Legend.IsVisible = True
            myPane.Title.IsVisible = False

            zg3.IsShowContextMenu = True
            zg3.IsEnableHEdit = False
            zg3.IsEnableZoom = True
            zg3.IsShowHScrollBar = True
            zg3.IsShowVScrollBar = False
            zg3.IsAutoScrollRange = True
            zg3.IsEnableVZoom = False
            zg3.IsShowPointValues = True
           
            SetSize()

            zg3.AxisChange()
         
            UpdateProgress.Value = 0
            UpdateProgress.Visible = False
            Label8.Visible = False
            Label10.Visible = False
            Label9.Visible = False
        End Using
        MsgBox("Done!")
        'Catch ex As Exception

        'End Try
    End Sub



    Private Sub Hpiez_txt_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Hpiez_txt.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
    End Sub


    Private Sub Hmin_txt_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Hmin_txt.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
    End Sub

    Private Sub txt_Dnb_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Dnb.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
    End Sub




    Private Sub txt_itD_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_itD.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
    End Sub

    Private Sub txt_Qup_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Qup.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then e.KeyChar = ""
    End Sub



    Private Sub CBRel_KeyPress_1(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs)
        e.KeyChar = ""
    End Sub


    Private Sub btn_GIS_HPD_Click(sender As System.Object, e As System.EventArgs) Handles btn_GIS_HPD.Click
        Try


            Dim hydrantShp As New MapWinGIS.Shapefile
            hydrantShp = g_MW.Layers(getLayerHandle(CmbHydrant.Text)).GetObject

            Dim shp_node As MapWinGIS.ShapefileClass
            shp_node = New MapWinGIS.ShapefileClass

            Dim Nshape_name As String = infolder & "\HPD.shp"
            hydrantShp.SaveAs(Nshape_name)

            'open the shapefile
            shp_node.Open(Nshape_name)
            shp_node.StartEditingTable()

            Dim n_field As Integer = hydrantShp.NumFields

            While n_field <> 1
                shp_node.EditDeleteField(1)
                n_field = shp_node.NumFields
            End While
            Dim fldIdx As Long = shp_node.NumFields
            For i As Integer = 1 To 11
                Dim newFld As New MapWinGIS.Field
                Select Case i
                    Case 1
                        newFld.Name = "ID"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                    Case 2
                        newFld.Name = "HPD_10"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 3
                        newFld.Name = "HPD_20"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 4
                        newFld.Name = "HPD_30"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 5
                        newFld.Name = "HPD_40"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 6
                        newFld.Name = "HPD_50"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 7
                        newFld.Name = "HPD_60"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 8
                        newFld.Name = "HPD_70"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 9
                        newFld.Name = "HPD_80"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 10
                        newFld.Name = "HPD_90"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 11
                        newFld.Name = "Reliability"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                End Select
                fldIdx = shp_node.NumFields
                shp_node.EditInsertField(newFld, fldIdx) 'insert column in dbf
            Next

            shp_node.EditDeleteField(0)
            'Try

            Dim f As Integer = -1

            For j As Integer = 0 To shp_node.NumShapes - 1 'row
                If grdNode.Item(3, j).Value = "Hydrant" Then
                    f += 1
                    For i As Integer = 0 To shp_node.NumFields - 1 'column
                        shp_node.EditCellValue(i, j, DGHPD.Item(i, f).Value)
                    Next
                End If
            Next

            shp_node.StopEditingTable(True)
            MsgBox("HPD results are converted to GIS")
        Catch ex As Exception
            MsgBox("Error")
        End Try
    End Sub

    Private Sub Btn_GIS_Vel_Click(sender As System.Object, e As System.EventArgs) Handles Btn_GIS_Vel.Click
        Try
            Dim NetworkShp As New MapWinGIS.Shapefile
            NetworkShp = g_MW.Layers(getLayerHandle(cmbNetwork.Text)).GetObject
            Dim shp_pipe As MapWinGIS.ShapefileClass
            shp_pipe = New MapWinGIS.ShapefileClass

            Dim Pshape_name As String = infolder & "\Velocity.shp"

            NetworkShp.SaveAs(Pshape_name)
            'open the shapefile
            shp_pipe.Open(Pshape_name)
            shp_pipe.StartEditingTable()

            Dim n_field As Integer = NetworkShp.NumFields


            While n_field <> 1
                shp_pipe.EditDeleteField(1)
                n_field = shp_pipe.NumFields
            End While
            Dim fldIdx As Long = shp_pipe.NumFields
            For i As Integer = 1 To 7
                Dim newFld As New MapWinGIS.Field
                Select Case i
                    Case 1
                        newFld.Name = "ID"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                    Case 2
                        newFld.Name = "V_10"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 3
                        newFld.Name = "V_20"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 4
                        newFld.Name = "V_30"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 5
                        newFld.Name = "V_40"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 6
                        newFld.Name = "V_50"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 7
                        newFld.Name = "V_60"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 8
                        newFld.Name = "V_70"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 9
                        newFld.Name = "V_80"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 10
                        newFld.Name = "V_90"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                    Case 11
                        newFld.Name = "Stdv"
                        newFld.Width = 5
                        newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                End Select
                fldIdx = shp_pipe.NumFields
                shp_pipe.EditInsertField(newFld, fldIdx) 'insert column in dbf
            Next

            shp_pipe.EditDeleteField(0)

            For i As Integer = 0 To shp_pipe.NumFields - 1
                For j As Integer = 0 To shp_pipe.NumShapes - 1
                    shp_pipe.EditCellValue(i, j, DGVelocity.Item(i, j).Value)
                Next
            Next
            shp_pipe.StopEditingTable(True)
            MsgBox("Velocity results are converted to GIS")
        Catch ex As Exception
            MsgBox("Error")
        End Try
    End Sub


    Private Sub Form1_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Dim response As MsgBoxResult
        response = MsgBox("Are you sure you want to exit", CType(MsgBoxStyle.Question + MsgBoxStyle.YesNo, MsgBoxStyle), "Confirm")
        If response = MsgBoxResult.No Then
            e.Cancel = True
        End If
    End Sub

    Private Sub Btn_SFR_Click(sender As System.Object, e As System.EventArgs) Handles Btn_SFR.Click
        Dim infile As String = infolder & "\Input.inp"
        Dim Qup As Single = CSng(Qup_SFR.Text)
        Dim hup As Single = CSng(Hpiez_txt.Text)


        Dim hs(nNode), qs(nNode) As Single
        Dim V(nLink), Qcli(nLink) As Single
        Array.Clear(hs, 0, nNode)

        For i As Integer = 1 To CInt(Conf_nb_SFR.Text)
            Application.DoEvents()
            write_epanet_file(infile, Qup, hup)
            run_epanet(infile, hs, qs, Qcli, V)
            For j As Integer = 1 To nLink
                If i = 1 Then
                    Link(j).Discharge = Qcli(j)
                ElseIf Qcli(j) > Link(j).Discharge Then
                    Link(j).Discharge = Qcli(j)
                End If
            Next
        Next
        For j As Integer = 1 To nLink
            grdLink.Item(4, j - 1).Value = FormatNumber(Link(j).Discharge, 0)
        Next

        BtnDesign_Click(sender, e)
        MsgBox("Done")

    End Sub





    Private Sub chkb_emitter_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkb_emitter.CheckedChanged
        If chkb_emitter.Checked = True Then
            GB_emitter.Enabled = False
            Chkb_FR.Checked = False
        Else
            GB_emitter.Enabled = True

        End If
    End Sub








    Private Sub Btnedit_Click(sender As System.Object, e As System.EventArgs) Handles Btnedit.Click

        If CBpump.Text = "" Then
            MessageBox.Show("Select pump first!", "", MessageBoxButtons.OK)
            Exit Sub
        Else
            Try
                Dim result As Integer = MessageBox.Show("Are you sure you want to change " & CBpump.Text & " ?", "caption", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    Exit Sub
                ElseIf result = DialogResult.Yes Then

                    Dim con As New OleDb.OleDbConnection

                    Dim ds As New System.Data.DataSet
                    Dim da As OleDb.OleDbDataAdapter
                    Dim sql As String
                    Dim dbProvider As String
                    Dim dbSource As String

                    dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
                    dbSource = "Data Source=" & myStream

                    con.ConnectionString = dbProvider & dbSource
                    con.Open()

                    sql = "select distinct * FROM Pumptbl"
                    da = New OleDb.OleDbDataAdapter(sql, con)
                    da.Fill(ds, "Pump table") ' You can call it "Bacon Sandwish" and will still work
                    con.Close() ' ds is now the table with the infor


                    ' remove all database
                    Dim Nrows As Integer = ds.Tables("pump table").Rows.Count - 1
                    For i As Integer = 0 To Nrows
                        If ds.Tables("pump table").Rows(i).Item(1) = CBpump.Text Then
                            ds.Tables("pump table").Rows(i).Delete()
                            Nrows = Nrows - 1
                        End If
                    Next

                    Try


                        'add new rows
                        For i As Integer = 0 To DG_pumpChar.Rows.Count - 1
                            Dim dsnewrow As DataRow
                            dsnewrow = ds.Tables("pump table").NewRow()
                            dsnewrow.Item(1) = CBpump.Text 'Pump name
                            dsnewrow.Item(2) = DG_pumpChar.Item(0, i).Value 'Q
                            dsnewrow.Item(3) = DG_pumpChar.Item(1, i).Value 'H
                            '  dsnewrow.Item(4) = DG_pumpChar.Item(2, i).Value 'Efficiency
                            ds.Tables("pump table").Rows.Add(dsnewrow)
                        Next
                    Catch ex As Exception
                        MessageBox.Show("Missing pump data, check your input")
                        Exit Sub
                    End Try


                    Dim cb As New OleDb.OleDbCommandBuilder(da)
                    da.Update(ds, "pump table")
                    MsgBox(" Record was successfully updated")


                End If
            Catch ex As Exception
                MessageBox.Show("Unexpected error has occured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try



        End If

    End Sub

    Private Sub CBpump_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CBpump.SelectedIndexChanged
        Dim con As New OleDb.OleDbConnection

        Dim ds As New System.Data.DataSet
        Dim da As OleDb.OleDbDataAdapter
        Dim sql As String
        Dim dbProvider As String
        Dim dbSource As String

        dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        dbSource = "Data Source=" & System.Windows.Forms.Application.StartupPath & "\Plugins\COPAM\pump.mdb"

        con.ConnectionString = dbProvider & dbSource
        con.Open()

        sql = "select * FROM Pumptbl"
        da = New OleDb.OleDbDataAdapter(sql, con)
        da.Fill(ds, "pump table") ' You can call it "Bacon Sandwish" and will still work
        con.Close() ' ds is now the table with the infor

        Dim f As Integer = -1
        DG_pumpChar.Rows.Clear()
        For i As Integer = 0 To ds.Tables("pump table").Rows.Count - 1
            If ds.Tables("pump table").Rows(i).Item(1) = CBpump.Text Then
                DG_pumpChar.Rows.Add()
                f += 1
                DG_pumpChar.Item(0, f).Value = ds.Tables("pump table").Rows(i).Item(2)
                DG_pumpChar.Item(1, f).Value = ds.Tables("pump table").Rows(i).Item(3)
                '   DG_pumpChar.Item(2, f).Value = ds.Tables("pump table").Rows(i).Item(4)
            End If
        Next

        Nup.Value = f + 1

    End Sub

    Private Sub Btn_OFD_pump_Click(sender As System.Object, e As System.EventArgs) Handles Btn_OFD_pump.Click

        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath & "\Plugins\COPAM\"
        openFileDialog1.Filter = "mdb files (*.mdb)|*.mdb"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog1.FileName
                If (myStream IsNot Nothing) Then
                    CBpump.Items.Clear()
                    Dim con As New OleDb.OleDbConnection
                    Dim ds As New System.Data.DataSet
                    Dim da As OleDb.OleDbDataAdapter
                    Dim sql As String
                    Dim dbProvider As String
                    Dim dbSource As String

                    dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
                    dbSource = "Data Source=" & myStream

                    con.ConnectionString = dbProvider & dbSource
                    con.Open()

                    sql = "select * FROM Pumptbl"
                    da = New OleDb.OleDbDataAdapter(sql, con)
                    da.Fill(ds, "pump table") ' You can call it "Bacon Sandwish" and will still work
                    con.Close() ' ds is now the table with the infor

                    Dim f As Integer = -1
                    CBpump.Items.Clear()
                    For i As Integer = 1 To ds.Tables("pump table").Rows.Count
                        Dim str As String = ds.Tables("pump table").Rows(i - 1).Item(1)
                        If CBpump.Items.Count = 0 Then
                            CBpump.Items.Add(str)
                            GoTo 1
                        Else
                            For j As Integer = 1 To CBpump.Items.Count
                                If str = CBpump.Items(j - 1) Then
                                    GoTo 1
                                End If
                            Next
                        End If
                        CBpump.Items.Add(str)

1:
                    Next
                    Btnaddnew.Enabled = True
                    Btnedit.Enabled = True
                    BtnDelete.Enabled = True
                    GBpump.Enabled = True

                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
            End Try
        End If
    End Sub

    Private Sub Nup_ValueChanged_1(sender As System.Object, e As System.EventArgs) Handles Nup.ValueChanged
        Dim i As Integer = Nup.Value
        For j As Integer = 0 To i
            Select Case Nup.Value
                Case Is > DG_pumpChar.Rows.Count
                    DG_pumpChar.Rows.Add()
                Case Is < DG_pumpChar.Rows.Count
                    DG_pumpChar.Rows.RemoveAt(DG_pumpChar.Rows.GetLastRow(DataGridViewElementStates.None))
                Case Is = DG_pumpChar.Rows.Count
                    Exit Sub
            End Select
        Next
    End Sub


    Private Sub Btnaddnew_Click(sender As System.Object, e As System.EventArgs) Handles Btnaddnew.Click
        If pumpid_txt.Text = "" Then
            MessageBox.Show("Select a name for the pump first!", "", MessageBoxButtons.OK)
            Exit Sub
        Else
            'Try


            Dim con As New OleDb.OleDbConnection

            Dim ds As New System.Data.DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim sql As String
            Dim dbProvider As String
            Dim dbSource As String

            dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
            dbSource = "Data Source=" & myStream

            con.ConnectionString = dbProvider & dbSource
            con.Open()

            sql = "select distinct * FROM Pumptbl"
            da = New OleDb.OleDbDataAdapter(sql, con)
            da.Fill(ds, "Pump table") ' You can call it "Bacon Sandwish" and will still work
            con.Close() ' ds is now the table with the infor

            ' Check for an existing record
            Dim Nrows As Integer = ds.Tables("pump table").Rows.Count - 1
            For i As Integer = 0 To Nrows
                If ds.Tables("pump table").Rows(i).Item(1) = pumpid_txt.Text Then
                    MessageBox.Show("Record already exist")
                    Exit Sub
                End If
            Next

            Try
                'add new rows
                For i As Integer = 0 To DG_pumpChar.Rows.Count - 1
                    Dim dsnewrow As DataRow
                    dsnewrow = ds.Tables("pump table").NewRow()
                    dsnewrow.Item(1) = pumpid_txt.Text 'Pump name
                    dsnewrow.Item(2) = DG_pumpChar.Item(0, i).Value 'Q
                    dsnewrow.Item(3) = DG_pumpChar.Item(1, i).Value 'H
                    'dsnewrow.Item(4) = DG_pumpChar.Item(2, i).Value 'Efficiency
                    ds.Tables("pump table").Rows.Add(dsnewrow)
                Next
                Dim cb As New OleDb.OleDbCommandBuilder(da)
                da.Update(ds, "pump table")
                CBpump.Items.Add(pumpid_txt.Text)
                pumpid_txt.Text = ""
                MsgBox(" Record was successfully added")
            Catch ex As Exception
                MessageBox.Show("Missing pump data, check your input")
                Exit Sub
            End Try






            'Catch ex As Exception
            '    MessageBox.Show("Unexpected error has occured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'End Try



        End If


    End Sub

    Private Sub BtnDelete_Click(sender As System.Object, e As System.EventArgs) Handles BtnDelete.Click
        If CBpump.Text = "" Then
            MessageBox.Show("Select pump first!", "", MessageBoxButtons.OK)
            Exit Sub
        Else
            Try
                Dim result As Integer = MessageBox.Show("Are you sure you want to delete " & CBpump.Text & " ?", "caption", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    Exit Sub
                ElseIf result = DialogResult.Yes Then

                    Dim con As New OleDb.OleDbConnection

                    Dim ds As New System.Data.DataSet
                    Dim da As OleDb.OleDbDataAdapter
                    Dim sql As String
                    Dim dbProvider As String
                    Dim dbSource As String

                    dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
                    dbSource = "Data Source=" & myStream

                    con.ConnectionString = dbProvider & dbSource
                    con.Open()

                    sql = "select distinct * FROM Pumptbl"
                    da = New OleDb.OleDbDataAdapter(sql, con)
                    da.Fill(ds, "Pump table") ' You can call it "Bacon Sandwish" and will still work
                    con.Close() ' ds is now the table with the infor


                    ' remove all database
                    Dim Nrows As Integer = ds.Tables("pump table").Rows.Count - 1
                    For i As Integer = 0 To Nrows
                        If ds.Tables("pump table").Rows(i).Item(1) = CBpump.Text Then
                            ds.Tables("pump table").Rows(i).Delete()
                            Nrows = Nrows - 1
                        End If
                    Next
                    DG_pumpChar.Rows.Clear()
                    Nup.Value = 0
                    CBpump.Items.Remove(CBpump.Text)
                    CBpump.Text = CBpump.Items(0)

                    Dim cb As New OleDb.OleDbCommandBuilder(da)
                    da.Update(ds, "pump table")

                    MsgBox(" Record was successfully removed")


                End If
            Catch ex As Exception
                MessageBox.Show("Unexpected error has occured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try



        End If
    End Sub


    Private Sub NUP_pump_ValueChanged(sender As System.Object, e As System.EventArgs) Handles NUP_pump.ValueChanged
        Dim i As Integer = NUP_pump.Value

        CType(DG_pump.Columns(0), DataGridViewComboBoxColumn).Items.Clear()


        CType(DG_pump.Columns(1), DataGridViewComboBoxColumn).Items.Clear()
        For ii As Integer = 1 To CBpump.Items.Count
            CType(DG_pump.Columns(0), DataGridViewComboBoxColumn).Items.Add(CBpump.Items(ii - 1))
        Next
        For ii As Integer = 1 To grdLink.Rows.Count
            CType(DG_pump.Columns(1), DataGridViewComboBoxColumn).Items.Add(CStr(grdLink.Item(0, ii - 1).Value))
        Next



        For j As Integer = 0 To i
            Select Case NUP_pump.Value
                Case Is > DG_pump.Rows.Count
                    DG_pump.Rows.Add()

                Case Is < DG_pump.Rows.Count
                    DG_pump.Rows.RemoveAt(DG_pump.Rows.GetLastRow(DataGridViewElementStates.None))
                Case Is = DG_pump.Rows.Count
                    Exit Sub
            End Select
        Next


    End Sub

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        clearGraphic()

        UpdateProgress.Maximum = DGHPD.Rows.Count
        UpdateProgress.Minimum = 0
        UpdateProgress.Visible = True


        Dim O, D As Integer

        Dim hmin As Single = CSng(TxtHmin.Text)

        DGloop.Rows.Clear()
        Dim H_id As Integer
        Select Case CBloop.Text 'HPD %
            Case "10%"
                H_id = 1
            Case "20%"
                H_id = 2
            Case "30%"
                H_id = 3
            Case "40%"
                H_id = 4
            Case "50%"
                H_id = 5
            Case "60%"
                H_id = 6
            Case "70%"
                H_id = 7
            Case "80%"
                H_id = 8
            Case "90%"
                H_id = 9
        End Select
        ReDim loop_C(DGHPD.Rows.Count)
        Dim f As Integer = -1
        For i As Integer = 1 To DGHPD.Rows.Count
            Application.DoEvents()
            UpdateProgress.Value = i
            Dim idf As Integer = DGHPD.Item(0, i - 1).Value
            O = Node(idf).id
            Dim x1 As Single = Node(idf).x
            Dim y1 As Single = Node(idf).y

            Dim H1 As Single = (DGHPD.Item(H_id, i - 1).Value * hmin) + hmin
            Dim beta As Single = 10000000

            If H1 < hmin Then ' only hydrants with low pressure
                f = f + 1
                For j As Integer = 1 To DGHPD.Rows.Count
                    Dim idi As Integer = DGHPD.Item(0, j - 1).Value
                    D = Node(idi).id
                    Dim H2 As Single = (DGHPD.Item(H_id, j - 1).Value * hmin) + hmin


                    Dim testpath As PathProperties
                    Main_SP.Main()
                    testpath = createSPaths(FlowPathHyd, PipeNetwork, O, D)
                    If O <> D And H2 > hmin Then ' no link , different points and low pressure
                        testpath = createSPaths(FlowPathHyd, PipeNetwork, O, D)
                        If testpath.nPt > 1 Then


                            Dim x2 As Single = Node(idi).x
                            Dim y2 As Single = Node(idi).y
                            Dim L As Single = Math.Sqrt(((x1 - x2) ^ 2) + ((y1 - y2) ^ 2))
                            Dim deltaH As Single = Abs(H1 - H2)
                            If beta > L / deltaH Then
                                beta = L / deltaH
                                loop_C(f).id = f + 1
                                loop_C(f).NodeF = O
                                loop_C(f).NodeT = D
                                loop_C(f).L = FormatNumber(L, 0)
                                loop_C(f).DH = FormatNumber(deltaH, 2)
                                loop_C(f).beta = FormatNumber(beta, 2)
                            End If
                        End If
                    End If
                Next
                DGloop.Rows.Add()
                DGloop.Item(0, f).Value = loop_C(f).id
                DGloop.Item(1, f).Value = loop_C(f).NodeF
                DGloop.Item(2, f).Value = loop_C(f).NodeT
                DGloop.Item(3, f).Value = loop_C(f).L
                DGloop.Item(4, f).Value = loop_C(f).DH
                DGloop.Item(5, f).Value = loop_C(f).beta
            Else
            End If ' low pressure

        Next



        UpdateProgress.Visible = True

        MsgBox("finished")


    End Sub

    Private Sub DGloop_DoubleClick(sender As System.Object, e As System.EventArgs) Handles DGloop.DoubleClick
        Try

            clearGraphic()


            For i As Integer = 0 To DGloop.Rows.Count - 1
                If DGloop.Rows(i).Selected = True Then
                    Dim NIN, NFI As Integer
                    NIN = DGloop.Item(1, i).Value
                    NFI = DGloop.Item(2, i).Value

                    LineGIS(Node(NIN).x, Node(NIN).y, Node(NFI).x, Node(NFI).y, 3, Drawing.Color.Red)

                    '  Exit Sub
                End If

            Next

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CBloop_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBloop.SelectedIndexChanged

    End Sub

    Private Sub Label43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label43.Click

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Try
        Dim sfd As New SaveFileDialog
        sfd.Filter = "Shapefiles (*.shp)|*.shp"
        If sfd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If IO.File.Exists(sfd.FileName) Then
                IO.File.Delete(sfd.FileName)
                IO.File.Delete(IO.Path.ChangeExtension(sfd.FileName, ".shx"))
                IO.File.Delete(IO.Path.ChangeExtension(sfd.FileName, ".dbf"))
            End If

            Create_Pipe_GIS_loop(sfd.FileName)
        End If

    End Sub

    Sub Create_Pipe_GIS_loop(ByVal path As String)

        Dim NetworkShp As New MapWinGIS.Shapefile
        NetworkShp = g_MW.Layers(getLayerHandle(cmbNetwork.Text)).GetObject
        Dim shp_pipe As MapWinGIS.ShapefileClass
        shp_pipe = New MapWinGIS.ShapefileClass

        Dim Pshape_name As String = path

        NetworkShp.SaveAs(Pshape_name)
        'open the shapefile
        shp_pipe.Open(Pshape_name)
        shp_pipe.StartEditingTable()

        Dim n_field As Integer = NetworkShp.NumFields


        While n_field <> 1
            shp_pipe.EditDeleteField(1)
            n_field = shp_pipe.NumFields
        End While
        Dim fldIdx As Long = shp_pipe.NumFields
        For i As Integer = 1 To 7
            Dim newFld As New MapWinGIS.Field
            Select Case i
                Case 1
                    newFld.Name = "ID"
                    newFld.Width = 5
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 2
                    newFld.Name = "IN"
                    newFld.Width = 5
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 3
                    newFld.Name = "FN"
                    newFld.Width = 5
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 4
                    newFld.Name = "Length"
                    newFld.Width = 50
                    newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                Case 5
                    newFld.Name = "QCl"
                    newFld.Width = 50
                    newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
                Case 6
                    newFld.Name = "Diam"
                    newFld.Width = 50
                    newFld.Type = MapWinGIS.FieldType.INTEGER_FIELD
                Case 7
                    newFld.Name = "Rough"
                    newFld.Width = 50
                    newFld.Type = MapWinGIS.FieldType.DOUBLE_FIELD
            End Select
            fldIdx = shp_pipe.NumFields
            shp_pipe.EditInsertField(newFld, fldIdx) 'insert column in dbf
        Next

        shp_pipe.EditDeleteField(0)

        For i As Integer = 0 To shp_pipe.NumFields - 1
            For j As Integer = 0 To shp_pipe.NumShapes - 1
                shp_pipe.EditCellValue(i, j, grdLink.Item(i, j).Value)
            Next
        Next

        shp_pipe.StopEditingTable(True)


        Dim newshapeindex = shp_pipe.NumShapes
        '3)points to be added
        Dim xypoints As MapWinGIS.Point
        Dim ff As Integer = nLink

        For i As Integer = 1 To DGloop.Rows.Count
            If DGloop.Rows(i - 1).Selected = True Then
                ff += 1
                Dim shape As New MapWinGIS.Shape
                shp_pipe.StartEditingShapes(True)
                shape.Create(MapWinGIS.ShpfileType.SHP_POLYLINE)
                xypoints = New MapWinGIS.Point()
                xypoints.x = Node(DGloop.Item(1, i - 1).Value).x
                xypoints.y = Node(DGloop.Item(1, i - 1).Value).y
                shape.InsertPoint(xypoints, 0)

                xypoints = New MapWinGIS.Point()
                xypoints.x = Node(DGloop.Item(2, i - 1).Value).x
                xypoints.y = Node(DGloop.Item(2, i - 1).Value).y
                shape.InsertPoint(xypoints, 0)
                shp_pipe.EditInsertShape(shape, newshapeindex) 'insert shape "row" inthe dbf

                For f As Integer = 0 To 6
                    If f = 0 Then
                        shp_pipe.EditCellValue(f, newshapeindex, ff)
                    End If
                    If f = 1 Then
                        shp_pipe.EditCellValue(f, newshapeindex, DGloop.Item(1, i - 1).Value)
                    End If
                    If f = 2 Then
                        shp_pipe.EditCellValue(f, newshapeindex, DGloop.Item(2, i - 1).Value)
                    End If
                    If f = 3 Then
                        shp_pipe.EditCellValue(f, newshapeindex, DGloop.Item(3, i - 1).Value)
                    End If
                    If f = 4 Then
                        shp_pipe.EditCellValue(f, newshapeindex, 0)
                    End If
                    If f = 5 Then
                        shp_pipe.EditCellValue(f, newshapeindex, 0)
                    End If
                    If f = 6 Then
                        shp_pipe.EditCellValue(f, newshapeindex, 0)
                    End If
                Next

            End If

        Next

        shp_pipe.StopEditingShapes(True, True)

    End Sub

    Private Sub zg3_Load(sender As System.Object, e As System.EventArgs) Handles zg3.Load

    End Sub

    Private Sub ZedGraphControl1_Load(sender As System.Object, e As System.EventArgs) Handles ZedGraphControl1.Load

    End Sub

    Private Sub BtnClement_Click_1(sender As System.Object, e As System.EventArgs) Handles BtnClement.Click

        clearGraphic()
        Dim Network As New MapWinGIS.Shapefile
        Network = g_MW.Layers(getLayerHandle(cmbNetwork.Text)).GetObject
        Dim flowPath As PathProperties


        Main_SP.Main()

        Dim F, T As Integer
        Dim col As UInteger = System.Convert.ToUInt32(RGB(255, 255, 255))
        UpdateProgress.Maximum = nNode
        UpdateProgress.Minimum = 0
        UpdateProgress.Visible = True
        ReDim Q(nLink)
        Dim A(nLink) As Single  'Sum of area
        Dim Nhyd(nLink) As Integer 'Nb of downstream hydrants


        Dim O, D As Integer
        Array.Clear(Q, 0, nLink)
        Array.Clear(Nhyd, 0, nLink)

        Dim res_i As Integer
        If nres = 1 Or CB_disch.Text = "Selected Reservoir" Then 'calculation from unique reservoir
            res_i = 1
        Else
            res_i = nres
        End If



        For Ff As Integer = 1 To res_i ' Clement per reservoirs
            clearGraphic()
            If nres > 1 And CB_disch.Text = "Selected Reservoir" Then
                O = CInt(CB_Cl_Res.Text)
            Else
                O = ResId(Ff)
            End If

            '  O = 1
            For iNode As Integer = 1 To nNode

                D = Node(iNode).id
                If O <> D Then
                    flowPath = createSPaths(FlowPathHyd, PipeNetwork, O, D)
                    Application.DoEvents()

                    If flowPath.nPt = 1 Then
                        'Show node didn't connect to origin node
                        ' g_MW.View.Draw.DrawPoint(Node(iNode).x, Node(iNode).y, 10, Color.Magenta)
                        PointGIS(Node(iNode).x, Node(iNode).y, 20, System.Drawing.Color.Blue)
                    Else
                        If Node(iNode).Q <> 0 Then
                            For i As Integer = 1 To flowPath.nPt - 1
                                F = flowPath.pNode(i).id
                                T = flowPath.pNode(i + 1).id
                                Dim ShapeID As Integer = getLinkId(F, T)
                                Q(ShapeID) += Node(iNode).Q ' disch
                                Nhyd(ShapeID) += 1 'Number of downstream hydrants
                                A(ShapeID) += Node(iNode).A
                                drawShape(Network.Shape(ShapeID - 1), Drawing.Color.OrangeRed, 2)
                            Next
                        End If
                    End If
                End If
                UpdateProgress.Value = iNode
            Next

            For i As Integer = 1 To nLink
                Dim Area As Single = A(i)  '  ha total irrigated area
                Dim ds As Single = CSng(Q(i) / Nhyd(i)) ' nominal discharge 
                Dim r As Single = CInt(T_txt.Text) / 24 'utilization
                Dim qs As Single = CSng(qs_txt.Text)  ' specific continuous flow l/s/ha multiplied by total area
                Dim pq As Single = CSng(CB_pq.Text)
                Dim Qcl As Single = 0
                Dim Qmin As Single = 0
                If Nhyd(i) = 0 Then
                    Qmin = 0
                    Qcl = 0
                Else
                    If Nhyd(i) <= CInt(MinHyd_txt.Text) Then
                        Qmin = Nhyd(i) * ds
                    Else
                        Qmin = CInt(MinHyd_txt.Text) * ds
                    End If

                    Qcl = Clement(Nhyd(i), Area, ds, r, qs, pq)
                End If



                If Ff > 1 And Link(i).Discharge > Math.Max(Qmin, Qcl) Then
                Else
                    Link(i).Discharge = Math.Max(Qmin, Qcl)
                    grdLink.Item(4, i - 1).Value = FormatNumber(Math.Max(Qmin, Qcl), 2)
                End If
            Next

        Next
        MsgBox("done")
        UpdateProgress.Visible = False
    End Sub

 

End Class