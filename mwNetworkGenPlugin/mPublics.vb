Imports MapWinUtility

Module mPublics
    '    

    Public infolder As String 'folder for storing results
    Public g_MW As MapWindow.Interfaces.IMapWin
    Public g_MapWindowForm As System.Windows.Forms.Form

    Public nres As Integer 'nb of reservoirs
    Public npump As Integer

    Public ResId() As Integer ' reservoir ID

    Public myStream As String ' path for pump.mdb

    Public pumpcurve() As String
    Public onode As Integer

    ' Friend WithEvents G_form1 As Form1
    '    Public g_X As Double
    '    Public g_Y As Double
    '    Public iRow As Integer
    '    Public Xorigin, Yorigin As Double
    '    Public MoveSpk As Boolean = False
    Public m_sf As MapWinGIS.Shapefile

    '    Public FirstPointClick As Boolean = False
    '    Public SecondPointClick As Boolean = False
    '    '
    Public g_MapWin As MapWindow.Interfaces.IMapWin

    Public CenPTA As POINTAPI ' Centroid of Planted Area
    Public CenGRD As POINTAPI ' Centroid of Grid
    'Friend WithEvents pwaGIS01Editor As frmMain
    Friend WithEvents frmNetworkGen As Form1 'frmNetGen
    '------------------------------------------
    'pwaGIS01 vaiable
    Public n As Integer = 0
    Public AddNode As Boolean = False
    Public moveNode As Boolean = False
    Public radius As Double = 500
    Public xO, yO
    Public xx(), yy(), Opp()
    Public LastNode As Integer
    Public LastX, LastY
    Public NodeOpt As String
    Public SelEdgeRow As Integer = -1
    '
    'Active DMA Command
    Public Enum DMACommand
        none
        select1stDMA
        select2ndDMA
        drawSpiltLine
        SelectSpiltLine
        SelectDMA2split
        MeterType
        PipeSplittingType
        SelectNode
    End Enum
    'Active Pipe Command
    Public Enum PipeEditCommand
        none
        select1stPipe
        select2ndPipe
        SelectExtedSide
        'drawSpiltLine
        'MeterType
        'PipeSplittingType
    End Enum
    '----------------------------------------------------
    'Tempolary Object
    '----------------------------------------------------
    Public ActiveDMACommand As DMACommand
    Public ActivePipeEditCommand As PipeEditCommand
    Public tmpLine As New MapWinGIS.Shape
    Public tmpPt As New MapWinGIS.Point
    Public tmpPt2 As New MapWinGIS.Point

    Public tmp1stDMA As New MapWinGIS.Shape
    Public tmp2ndDMA As New MapWinGIS.Shape
    Public tmp1stPipe As New MapWinGIS.Shape
    Public tmp2ndPipe As New MapWinGIS.Shape

    Public tmpDMA4SplitID As Integer = -1
    Public tmp1stDMAID As Integer = -1
    Public tmp2ndDMAID As Integer = -1
    '
    Public tmp1stPipeID As Integer = -1
    Public tmp2ndPipeID As Integer = -1
    'Tempolary Coordinate
    Public Xorigin, Yorigin As Double

    'MODEL
    Public Structure SegmentProperties
        Dim PipeID As String
        Dim x1 As Double
        Dim y1 As Double
        Dim x2 As Double
        Dim y2 As String
    End Structure

    Public Structure checkvalve
        Dim F As Integer
        Dim T As Integer
    End Structure

   
    '-------------------------------------------
    'Public NetFile As String
    Public DecreaseFac As Double = 1
    '------------------------------------------

    'DMA Layer
    Public DMA_Field(5) As MapWinGIS.Field
    Public nDMA_Field As Integer = 5
    'DMA_Edge Layer
    Public DMAEDGE_Field(3) As MapWinGIS.Field
    Public nDMAEDGE_Field As Integer = 3
    'Node Layer
    Public Node_Field(4) As MapWinGIS.Field
    Public nNode_Field As Integer = 4
    'Meter Layer
    Public Meter_Field(8) As MapWinGIS.Field
    Public nMeter_Field As Integer = 8
    'Pipe Layer
    Public Pipe_Field(12) As MapWinGIS.Field
    Public nPipe_Field As Integer = 12
    '    
    Public CHKALL As Boolean = False
    '
    Structure IdDMA
        Dim DMAZONE As Integer
        Dim ZoneName As Integer
        Dim numMeter As Integer
        Dim Area As Integer
        Dim numBldg As Integer
    End Structure

    Structure IdDMAEDGE
        Dim MWShapeID As Integer
        Dim Type As Integer
        Dim PIPE_ID As Integer
    End Structure

    Structure IdNode
        Dim NodeID As Integer
        Dim Type As Integer
        Dim Z As Integer
        Dim Demand As Integer
    End Structure

    Structure IdMeter
        Dim BLDG_ID As Integer
        Dim PIPE_ID As Integer
        Dim CUSTCODE As Integer
        Dim xLINK As Integer
        Dim yLINK As Integer
        Dim LinkID As Integer
    End Structure

    Structure IdPipe
        Dim ELEVATION As Integer
        Dim PIPE_ID As Integer
        Dim PIPE_TYPE As Integer
        Dim PIPE_SIZE As Integer
        Dim PipeCLASS As Integer
        Dim DEPTH As Integer
        Dim PipeLength As Integer
        Dim F As Integer
        Dim T As Integer
        Dim LinkID As Integer
        Dim Discharge As Integer
        Dim Direction As Integer
    End Structure

    

    Public FieldDMA As IdDMA
    Public FieldDMAEdge As IdDMAEDGE
    Public FieldNode As IdNode
    Public FieldMeter As IdMeter
    Public FieldNetwork As IdPipe

    Public Structure pipeDBProperties
        Dim PIPE_SIZE As Double
        Dim PIPE_TYPE As String
        Dim PIPE_CLASS As String
    End Structure

    Public pipeDB(19) As pipeDBProperties
    Public BillingDBFile As String = ""

    Public Function getColNum(ByVal FieldName As String, ByVal DatabaseName As String) As Integer
        Dim Col As Integer = -1
        Dim TestTable As New MapWinGIS.Table()
        Dim success As Boolean
        Dim Fname As String = DatabaseName

        success = TestTable.Open(Fname)
        If success = True Then
            For i As Integer = 0 To TestTable.NumFields - 1
                If UCase(FieldName) = UCase(TestTable.Field(i).Name) Then
                    Col = i
                End If
            Next
            TestTable.Close()
        Else
            Col = -1
        End If

        Return Col
    End Function
    '
    ' Get Column number
    '
    Public Function getColNum(ByVal FieldName As String, ByVal Shapefile As MapWinGIS.Shapefile) As Integer
        Dim Col As Integer = -1
        Try
            For i As Integer = 0 To Shapefile.NumFields - 1
                If UCase(FieldName) = UCase(Shapefile.Field(i).Name) Then
                    Col = i
                End If
            Next
        Catch ex As Exception
            Return -1
        End Try
        Return Col
    End Function
    '
    ' Get Column number
    '
    Public Function getColNum(ByVal FieldName As String, ByVal grdView As System.Windows.Forms.DataGridView) As Integer
        Dim Col As Integer = -1
        Try
            For i As Integer = 0 To grdView.ColumnCount - 1
                If UCase(FieldName) = UCase(grdView.Columns(i).HeaderText) Then
                    Col = i
                End If
            Next
        Catch ex As Exception
            Return -1
        End Try
        Return Col
    End Function

    Public Function getLayerHandle(ByVal layerName As String) As Integer
        For i As Integer = 0 To g_MW.Layers.NumLayers - 1
            If UCase(g_MW.Layers(g_MW.Layers.GetHandle(i)).Name) = UCase(layerName) Then
                Return g_MW.Layers.GetHandle(i)
            End If
        Next
        Return -1
    End Function

    '    Public Function getPotID(ByVal id As Integer, ByVal layerName As String, ByRef pt As MapWinGIS.Point, ByRef area As Double) As Integer

    '        Dim iLayer As Integer = getLayerHandle(layerName)
    '        Dim potShape As New MapWinGIS.Shapefile
    '        Dim success As Boolean = potShape.Open(g_MW.Layers(iLayer).FileName)
    '        area = 0
    '        'For i As Integer = 0 To potShape.NumShapes - 1
    '        '    If CInt(potShape.CellValue(1, i)) = id Then
    '        Try
    '            pt = potShape.Shape(id).Centroid
    '            area = potShape.Shape(id).Area
    '            Dim BlockField As Integer = getColNum("block", potShape)
    '            BlockID = potShape.CellValue(BlockField, id) ' field block
    '        Catch ex As Exception

    '        End Try
    '        'Exit Function
    '        '    End If
    '        'Next

    '        Return -1
    '    End Function
    '#End Region
    Public Function ShpID(ByVal shp As MapWinGIS.Shapefile, ByVal pt As MapWinGIS.Point) As Integer
        Dim id As Integer
        Dim L As Double
        MapWinGeoProc.SpatialOperations.GetShapeNearestToPoint(shp, pt, id, L)
        'g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
        'g_MW.View.Draw.DrawCircle(pt.x, pt.y, 20, Drawing.Color.Black, False)
        Return id
    End Function


    Public Function intesecLC(ByVal cx As _
    Single, ByVal cy As Single, ByVal radius As Single, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As _
    Single, ByVal y2 As Single, ByRef ix1 As Single, ByRef _
    iy1 As Single, ByRef ix2 As Single, ByRef iy2 As _
    Single) As Integer
        Dim dx As Single
        Dim dy As Single
        Dim a As Single
        Dim B As Single
        Dim C As Single
        Dim det As Single
        Dim t As Single

        dx = x2 - x1
        dy = y2 - y1

        a = dx * dx + dy * dy
        B = 2 * (dx * (x1 - cx) + dy * (y1 - cy))
        C = (x1 - cx) * (x1 - cx) + (y1 - cy) * (y1 - cy) - _
            radius * radius

        det = B * B - 4 * a * C
        If (a <= 0.0000001) Or (det < 0) Then
            ' No real solutions.
            intesecLC = 0
        ElseIf det = 0 Then
            ' One solution.
            t = -B / (2 * a)
            ix1 = x1 + t * dx
            iy1 = y1 + t * dy
            intesecLC = 1
        Else
            ' Two solutions.
            t = (-B + Math.Sqrt(det)) / (2 * a)
            ix1 = x1 + t * dx
            iy1 = y1 + t * dy
            t = (-B - Math.Sqrt(det)) / (2 * a)
            ix2 = x1 + t * dx
            iy2 = y1 + t * dy
            intesecLC = 2
        End If
    End Function

    Public Function NodeCol(ByVal Typ As String) As System.Drawing.Color
        NodeCol = Drawing.Color.White
        If Typ = "Outlet" Then NodeCol = Drawing.Color.Cyan
        If Typ = "Junction" Then NodeCol = Drawing.Color.Yellow
    End Function

    Public Sub CreateSp(ByVal x As Double, ByVal y As Double, ByVal NodeTYP As String)
        On Error GoTo Error_handler

        Dim success As Boolean
        Dim fieldindex As Long
        Dim shapeIndex As Long
        Dim data As Integer
        'objects
        '
        'Todo: Change the active layer from current layer to hydrant Layer
        '
        'Dim sf As MapWinGIS.Shapefile = g_MW.Layers(getLayerHandle(LayoutEditor.cmbHydrantFile.Text)).GetObject()
        Dim HydLayerID As Integer ' = getLayerHandle(pwaGIS01Editor.cmbMeter.Text)
        Dim sf As MapWinGIS.Shapefile = g_MW.Layers(HydLayerID).GetObject()
        '
        ' Existing hydrant check
        Dim ExistPump As Boolean = False
        If NodeTYP = "PUMP" Then
            For i As Integer = 1 To sf.NumShapes - 1
                If sf.CellValue(2, i) = NodeTYP Then
                    ExistPump = True
                    PointGIS(sf.Shape(i).Point(0).x, sf.Shape(i).Point(0).y, 10, Drawing.Color.Red)
                End If
            Next
        End If
        If ExistPump = True Then
            MsgBox("Delete the existing pump before add new pump:")
            Exit Sub
        End If
        Dim field As MapWinGIS.Field
        Dim shape As MapWinGIS.Shape
        Dim point As MapWinGIS.Point

        'Create and add shape
        shape = New MapWinGIS.Shape
        sf.StartEditingShapes()

        'Create a new Point shape
        With shape
            success = .Create(MapWinGIS.ShpfileType.SHP_POINT)
            If Not success Then
                MsgBox("Error in creating shape: " & .ErrorMsg(.LastErrorCode))
                GoTo Error_handler
            End If
        End With ' shape

        point = New MapWinGIS.Point

        'Set the values for the point to be inserted
        point.x = x
        point.y = y
        'g_MW.Layers(g_MW.Layers.CurrentLayer).ClearLabels()
        'Dim xx, yy As Double
        'g_MW.View.ProjToPixel(x, y, xx, yy)
        'g_MW.Layers(g_MW.Layers.CurrentLayer).AddLabelEx("KASEM", Drawing.Color.Azure, x, y, MapWinGIS.tkHJustification.hjCenter, 0)
        g_MW.Layers(HydLayerID).LabelsVisible = True
        'g_MW.Layers(g_MW.Layers.CurrentLayer).LabelsVisible = True
        'Insert the point into the shape
        With shape
            success = .InsertPoint(point, 0) ' shape.numPoints)
            If Not success Then
                MsgBox("Error in adding point: " & .ErrorMsg(.LastErrorCode))
                GoTo Error_handler
            End If
        End With ' shape

        'Insert the shape into the shapefile
        shapeIndex = sf.NumShapes
        With sf
            success = .EditInsertShape(shape, shapeIndex)
            If Not success Then
                MsgBox("Error in adding field: " & .ErrorMsg(.LastErrorCode))
                GoTo Error_handler
            End If
        End With ' sf

        'Add value to at least one attribute
        'Use shapeindex as dummy value:
        If NodeTYP = "JUNCTION" Then
        End If
        data = shapeIndex
        With sf

            success = .EditCellValue(2, shapeIndex, NodeTYP)
            If Not success Then
                MsgBox("Error in adding field: " & .ErrorMsg(.LastErrorCode))
                GoTo Error_handler
            End If
        End With ' sf

        'Stop editing shapes in the shapefile, saving changes to shapes,
        'also stopping editing of the attribute table
        With sf
            success = sf.StopEditingShapes(True, True)
            If Not success Then
                MsgBox("Error in adding field: " & .ErrorMsg(.LastErrorCode))
                GoTo Error_handler
            End If

        End With ' sf

        g_MW.Refresh()
        'newShapefile = True

Cleanup:
        On Error Resume Next
        field = Nothing
        shape = Nothing
        point = Nothing
        sf = Nothing

        Exit Sub

Error_handler:
        'newShapefile = False
        GoTo Cleanup
    End Sub
    '
    '
    '
    Public Function getLinkId(ByVal n1 As Integer, ByVal n2 As Integer) As Integer
        Dim CHK As String = HashLinkID(n1 & "-" & n2)
        If CHK <> "" Then
            Return CInt(CHK)
        Else
            Return CInt(-1)
        End If
    End Function
    '
    '
    '
    'Public Function getNodeId(ByVal n1 As Integer) As Integer
    '   Dim CHK As String = HashNodeID(n1)
    '   If CHK <> "" Then
    '      Return CInt(CHK)
    '   Else
    '      Return CInt(-1)
    '   End If
    'End Function

    Function DrawValve(ByVal pt As MapWinGIS.Point, ByVal scale As Integer)
        Dim d As New MapWinGIS.Shapefile
        Dim i As Integer = d.ShapeCategory(0)
    End Function

    Function getLastPoint(ByVal shape As MapWinGIS.Shape) As MapWinGIS.Point
        getLastPoint = shape.Point(shape.numPoints - 1)
    End Function

    Function Distance2Point(ByVal pt1 As MapWinGIS.Point, ByVal pt2 As MapWinGIS.Point) As Double
        Distance2Point = ((pt1.x - pt2.x) ^ 2 + (pt1.y - pt2.y) ^ 2) ^ 0.5
    End Function

    Function SelNearEndLinePoint(ByVal L1 As MapWinGIS.Shape, ByVal L2 As MapWinGIS.Shape, ByRef p1 As MapWinGIS.Point, ByRef p2 As MapWinGIS.Point) As Boolean
        Dim pt1 As MapWinGIS.Point = L1.Point(0)
        Dim pt3 As MapWinGIS.Point = L2.Point(0)
        Dim pt2 As MapWinGIS.Point = getLastPoint(L1)
        Dim pt4 As MapWinGIS.Point = getLastPoint(L2)
        Dim R(3) As Double
        SelNearEndLinePoint = False
        R(0) = Distance2Point(pt1, pt3)
        R(1) = Distance2Point(pt1, pt4)
        R(2) = Distance2Point(pt2, pt3)
        R(3) = Distance2Point(pt2, pt4)
        Dim rMin As Double = 100000000000
        Dim iCase As Integer = -1
        For i As Integer = 0 To 3
            If rMin > R(i) Then
                rMin = R(i)
                iCase = i
                SelNearEndLinePoint = True
            End If
        Next
        If iCase = 0 Then
            'pt1-pt3
            p1 = pt1 : p2 = pt3
        End If
        If iCase = 1 Then
            'pt1-pt4
            p1 = pt1 : p2 = pt4
        End If
        If iCase = 2 Then
            'pt2-pt3
            p1 = pt2 : p2 = pt3
        End If
        If iCase = 3 Then
            'pt2-pt4
            p1 = pt2 : p2 = pt4
        End If
    End Function

    Public Sub getPipeDB()
        pipeDB(1).PIPE_SIZE = 20 : pipeDB(1).PIPE_TYPE = "PBP" : pipeDB(1).PIPE_CLASS = "13.5"
        pipeDB(2).PIPE_SIZE = 25 : pipeDB(2).PIPE_TYPE = "PBP" : pipeDB(2).PIPE_CLASS = "13.5"
        pipeDB(3).PIPE_SIZE = 40 : pipeDB(3).PIPE_TYPE = "PBP" : pipeDB(3).PIPE_CLASS = "13.5"
        pipeDB(4).PIPE_SIZE = 50 : pipeDB(4).PIPE_TYPE = "PBP" : pipeDB(4).PIPE_CLASS = "13.5"
        pipeDB(5).PIPE_SIZE = 100 : pipeDB(5).PIPE_TYPE = "GS" : pipeDB(5).PIPE_CLASS = "MED"
        pipeDB(6).PIPE_SIZE = 110 : pipeDB(6).PIPE_TYPE = "HDPE" : pipeDB(6).PIPE_CLASS = "PN6.3"
        pipeDB(7).PIPE_SIZE = 150 : pipeDB(7).PIPE_TYPE = "GS" : pipeDB(7).PIPE_CLASS = "MED"
        pipeDB(8).PIPE_SIZE = 160 : pipeDB(8).PIPE_TYPE = "HDPE" : pipeDB(8).PIPE_CLASS = "PN6.3"
        pipeDB(9).PIPE_SIZE = 200 : pipeDB(9).PIPE_TYPE = "STP" : pipeDB(9).PIPE_CLASS = "UNG"
        pipeDB(10).PIPE_SIZE = 225 : pipeDB(10).PIPE_TYPE = "HDPE" : pipeDB(10).PIPE_CLASS = "PN6.3"
        pipeDB(11).PIPE_SIZE = 250 : pipeDB(11).PIPE_TYPE = "HDPE" : pipeDB(11).PIPE_CLASS = "PN6.3"
        pipeDB(12).PIPE_SIZE = 280 : pipeDB(12).PIPE_TYPE = "HDPE" : pipeDB(12).PIPE_CLASS = "PN6.3"
        pipeDB(13).PIPE_SIZE = 300 : pipeDB(13).PIPE_TYPE = "STP" : pipeDB(13).PIPE_CLASS = "UNG"
        pipeDB(14).PIPE_SIZE = 315 : pipeDB(14).PIPE_TYPE = "HDPE" : pipeDB(14).PIPE_CLASS = "PN6.3"
        pipeDB(15).PIPE_SIZE = 400 : pipeDB(15).PIPE_TYPE = "AC" : pipeDB(15).PIPE_CLASS = "15"
        pipeDB(16).PIPE_SIZE = 500 : pipeDB(16).PIPE_TYPE = "STP" : pipeDB(16).PIPE_CLASS = "UNG"
        pipeDB(17).PIPE_SIZE = 560 : pipeDB(17).PIPE_TYPE = "HDPE" : pipeDB(17).PIPE_CLASS = "PN6.3"
        pipeDB(18).PIPE_SIZE = 600 : pipeDB(18).PIPE_TYPE = "STP" : pipeDB(18).PIPE_CLASS = "UNG"
        pipeDB(19).PIPE_SIZE = 630 : pipeDB(19).PIPE_TYPE = "HDPE" : pipeDB(19).PIPE_CLASS = "PN6.3"
    End Sub

    Public Function getComPipeSize(ByVal d As Double) As pipeDBProperties
        'Maximun pipe diameter is 630 mm.
        If d < pipeDB(1).PIPE_SIZE Then Return pipeDB(1)
        For i As Integer = 1 To 19 - 1
            If pipeDB(i).PIPE_SIZE <= d And d < pipeDB(i + 1).PIPE_SIZE Then
                Return pipeDB(i)
            End If
        Next
        If d > pipeDB(19).PIPE_SIZE Then Return pipeDB(19)
    End Function



    Function Clement(ByVal nj As Integer, ByVal A As Single, ByVal d As Single, ByVal r As Single, ByVal qs As Single, ByVal pq As Single) As Single

        Dim u As Single
        Dim Qclement As Single
        Select Case pq
            Case 0.9
                u = 1.285
            Case 0.91
                u = 1.345
            Case 0.92
                u = 1.405
            Case 0.93
                u = 1.475
            Case 0.94
                u = 1.555
            Case 0.95
                u = 1.645
            Case 0.96
                u = 1.755
            Case 0.97
                u = 1.885
            Case 0.98
                u = 2.055
            Case 0.99
                u = 2.324
        End Select
        Dim sumrpd, sum2 As Single
        Dim p As Single
        sumrpd = 0
        sum2 = 0


        p = (qs * A) / (r * nj * d)
        sumrpd = nj * p * d
        sum2 = nj * (p * (1 - p) * d * d)

        Qclement = sumrpd + (u * Math.Sqrt(sum2))
        Return Qclement

    End Function


End Module
