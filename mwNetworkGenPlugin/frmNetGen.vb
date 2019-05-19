Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmNetGen
  Private blInit As Boolean '=True if a graph has been initialized
  Private blCoord As Boolean '=True if node coordinates exist
  Dim LinkUnconnect() As Integer

#Region "FormControl"

  Private Sub frmDmaDesign_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    ClearComboLayer()
    loadcomboLayer()
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

  Function CloneShp(ByVal sourceShp As MapWinGIS.Shapefile) As MapWinGIS.Shapefile
    CloneShp = sourceShp.Clone
    CloneShp.StartEditingShapes()
    For i As Integer = 0 To sourceShp.NumShapes - 1
      CloneShp.EditInsertShape(sourceShp.Shape(i), i)
      For j As Integer = 0 To sourceShp.NumFields - 1
        CloneShp.EditCellValue(j, i, sourceShp.CellValue(j, i))
      Next
    Next
    CloneShp.StopEditingShapes()
  End Function

  Function SplitPipebyLine(ByVal Line2Split As MapWinGIS.Shape, ByVal Line4Split As MapWinGIS.Shape, ByRef Pipe1 As MapWinGIS.Shape, ByRef Pipe2 As MapWinGIS.Shape, ByVal Highlight As Boolean) As Boolean
    'Line2Split คือเส้นที่จะถูกตัด
    'Line4Split คือเส้นที่ใช้ตัด
    Pipe1.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
    Pipe2.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
    Dim iVertex As Integer = 0
    Dim PipePt As New MapWinGIS.Point
    Dim FirstPipe As Boolean = True
    Dim L1 As Geo_Function.LineType
    Dim L2 As Geo_Function.LineType
    Dim intersectFound As Boolean = False
    Dim haveIntersec As Boolean = False
    intersectFound = False

    For i As Integer = 0 To Line2Split.numPoints - 2
      L1.Pt1.X = Line2Split.Point(i).x
      L1.Pt1.Y = Line2Split.Point(i).y
      L1.Pt2.X = Line2Split.Point(i + 1).x
      L1.Pt2.Y = Line2Split.Point(i + 1).y
      If intersectFound = False Then
        For j As Integer = 0 To Line4Split.numPoints - 2
          L2.Pt1.X = Line4Split.Point(j).x
          L2.Pt1.Y = Line4Split.Point(j).y
          L2.Pt2.X = Line4Split.Point(j + 1).x
          L2.Pt2.Y = Line4Split.Point(j + 1).y
          Dim pt As Geo_Function.POINTAPI
          Dim x, y As Double

          If Geo_Function.lineIntersection(L1, L2, pt) = True And Geo_Function.Line_Line(L1, L2, x, y) = True Then
            'pt = จุดที่ใช้ตัดเส้น
            If Highlight = True Then
              g_MW.View.Draw.DrawPoint(pt.X, pt.Y, 12, Drawing.Color.Black)
              g_MW.View.Draw.DrawPoint(pt.X, pt.Y, 10, Drawing.Color.Magenta)
            End If
            'If Pipe1.numPoints = 0 Then
            PipePt = New MapWinGIS.Point
            PipePt.x = L1.Pt1.X
            PipePt.y = L1.Pt1.Y
            Pipe1.InsertPoint(PipePt, iVertex)
            iVertex += 1
            'End If
            PipePt = New MapWinGIS.Point
            PipePt.x = pt.X
            PipePt.y = pt.Y
            Pipe1.InsertPoint(PipePt, iVertex)
            iVertex += 1
            FirstPipe = False
            iVertex = 0


            Pipe2.InsertPoint(PipePt, iVertex)
            iVertex += 1
            PipePt = New MapWinGIS.Point
            PipePt.x = L1.Pt2.X
            PipePt.y = L1.Pt2.Y
            Pipe2.InsertPoint(PipePt, iVertex)
            iVertex += 1
            intersectFound = True
            haveIntersec = True
            Exit For 'พบจุดที่เส้นที่ตัดกันแล้ว หมายเหตุเอาเฉพาะจุดที่ตัดกันจุดแรกเท่านั้น
          End If
        Next
        'ถ้ายังไม่พบจุดตัด เพิ่มจุดในเส้น
        If FirstPipe = True Then
          PipePt = New MapWinGIS.Point
          PipePt.x = L1.Pt1.X
          PipePt.y = L1.Pt1.Y
          Pipe1.InsertPoint(PipePt, iVertex)
          iVertex += 1
        End If
      Else 'หลังจากพบจุดตัดแล้ว
        PipePt = New MapWinGIS.Point
        PipePt.x = L1.Pt1.X
        PipePt.y = L1.Pt1.Y
        Pipe2.InsertPoint(PipePt, iVertex)
        iVertex += 1
      End If

    Next
    If haveIntersec = False Then
      PipePt = New MapWinGIS.Point
      PipePt.x = L1.Pt2.X
      PipePt.y = L1.Pt2.Y
      Pipe1.InsertPoint(PipePt, iVertex)
    Else
      PipePt = New MapWinGIS.Point
      PipePt.x = L1.Pt2.X
      PipePt.y = L1.Pt2.Y
      Pipe2.InsertPoint(PipePt, iVertex)
    End If
    If Highlight = True Then
      For i As Integer = 0 To Pipe1.numPoints - 2
        g_MW.View.Draw.DrawLine(Pipe1.Point(i).x, Pipe1.Point(i).y, Pipe1.Point(i + 1).x, Pipe1.Point(i + 1).y, 4, Drawing.Color.Magenta)
      Next
      For i As Integer = 0 To Pipe2.numPoints - 2
        g_MW.View.Draw.DrawLine(Pipe2.Point(i).x, Pipe2.Point(i).y, Pipe2.Point(i + 1).x, Pipe2.Point(i + 1).y, 4, Drawing.Color.Green)
      Next
    End If
    Return haveIntersec
  End Function

  Function SplitPipebyLine(ByVal Network As MapWinGIS.Shapefile, ByVal ActID As Integer, ByVal Line4Split As MapWinGIS.Shape, ByRef Pipe1 As MapWinGIS.Shape, ByRef Pipe2 As MapWinGIS.Shape, ByVal Highlight As Boolean) As Boolean
    Pipe1.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
    Pipe2.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
    Dim iVertex As Integer = 0
    Dim PipePt As New MapWinGIS.Point
    Dim FirstPipe As Boolean = True
    Dim L1 As Geo_Function.LineType
    Dim L2 As Geo_Function.LineType
    Dim intersectFound As Boolean = False
    Dim haveIntersec As Boolean = False

    For i As Integer = 0 To Network.Shape(ActID).numPoints - 2
      L1.Pt1.X = Network.Shape(ActID).Point(i).x
      L1.Pt1.Y = Network.Shape(ActID).Point(i).y
      L1.Pt2.X = Network.Shape(ActID).Point(i + 1).x
      L1.Pt2.Y = Network.Shape(ActID).Point(i + 1).y
      If intersectFound = False Then
        For j As Integer = 0 To Line4Split.numPoints - 2
          L2.Pt1.X = Line4Split.Point(j).x
          L2.Pt1.Y = Line4Split.Point(j).y
          L2.Pt2.X = Line4Split.Point(j + 1).x
          L2.Pt2.Y = Line4Split.Point(j + 1).y
          Dim pt As Geo_Function.POINTAPI
          Dim x, y As Double

          If Geo_Function.lineIntersection(L1, L2, pt) = True And Geo_Function.Line_Line(L1, L2, x, y) = True Then
            'pt = จุดที่ใช้ตัดเส้น
            If Highlight = True Then
              g_MW.View.Draw.DrawPoint(pt.X, pt.Y, 12, Drawing.Color.Black)
              g_MW.View.Draw.DrawPoint(pt.X, pt.Y, 10, Drawing.Color.Magenta)
            End If
            'If Pipe1.numPoints = 0 Then
            PipePt = New MapWinGIS.Point
            PipePt.x = L1.Pt1.X
            PipePt.y = L1.Pt1.Y
            Pipe1.InsertPoint(PipePt, iVertex)
            iVertex += 1
            'End If
            PipePt = New MapWinGIS.Point
            PipePt.x = pt.X
            PipePt.y = pt.Y
            Pipe1.InsertPoint(PipePt, iVertex)
            iVertex += 1
            FirstPipe = False
            iVertex = 0


            Pipe2.InsertPoint(PipePt, iVertex)
            iVertex += 1
            PipePt = New MapWinGIS.Point
            PipePt.x = L1.Pt2.X
            PipePt.y = L1.Pt2.Y
            Pipe2.InsertPoint(PipePt, iVertex)
            iVertex += 1
            intersectFound = True
            haveIntersec = True
            Exit For 'พบจุดที่เส้นที่ตัดกันแล้ว หมายเหตุเอาเฉพาะจุดที่ตัดกันจุดแรกเท่านั้น
          End If
        Next
        'ถ้ายังไม่พบจุดตัด เพิ่มจุดในเส้น
        If FirstPipe = True Then
          PipePt = New MapWinGIS.Point
          PipePt.x = L1.Pt1.X
          PipePt.y = L1.Pt1.Y
          Pipe1.InsertPoint(PipePt, iVertex)
          iVertex += 1
        End If
      Else 'หลังจากพบจุดตัดแล้ว
        PipePt = New MapWinGIS.Point
        PipePt.x = L1.Pt1.X
        PipePt.y = L1.Pt1.Y
        Pipe2.InsertPoint(PipePt, iVertex)
        iVertex += 1
      End If

    Next
    If haveIntersec = False Then
      PipePt = New MapWinGIS.Point
      PipePt.x = L1.Pt2.X
      PipePt.y = L1.Pt2.Y
      Pipe1.InsertPoint(PipePt, iVertex)
    Else
      PipePt = New MapWinGIS.Point
      PipePt.x = L1.Pt2.X
      PipePt.y = L1.Pt2.Y
      Pipe2.InsertPoint(PipePt, iVertex)
    End If
    If Highlight = True Then
      For i As Integer = 0 To Pipe1.numPoints - 2
        g_MW.View.Draw.DrawLine(Pipe1.Point(i).x, Pipe1.Point(i).y, Pipe1.Point(i + 1).x, Pipe1.Point(i + 1).y, 4, Drawing.Color.Magenta)
      Next
      For i As Integer = 0 To Pipe2.numPoints - 2
        g_MW.View.Draw.DrawLine(Pipe2.Point(i).x, Pipe2.Point(i).y, Pipe2.Point(i + 1).x, Pipe2.Point(i + 1).y, 4, Drawing.Color.Green)
      Next
    End If
    Return haveIntersec
  End Function

  Function SplitPipebyPoint(ByVal Network As MapWinGIS.Shapefile, ByVal ActID As Integer, ByVal Point4Split As MapWinGIS.Point, ByRef Pipe1 As MapWinGIS.Shape, ByRef Pipe2 As MapWinGIS.Shape, ByVal Highlight As Boolean) As Boolean
    Pipe1 = New MapWinGIS.Shape
    Pipe2 = New MapWinGIS.Shape
    Pipe1.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
    Pipe2.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
    Dim iVertex As Integer = 0
    Dim PipePt As New MapWinGIS.Point
    Dim FirstPipe As Boolean = True
    Dim L1 As Geo_Function.LineType
    Dim L2 As Geo_Function.LineType
    Dim intersectFound As Boolean = False
    Dim haveIntersec As Boolean = False

    For i As Integer = 0 To Network.Shape(ActID).numPoints - 2
      L1.Pt1.X = Network.Shape(ActID).Point(i).x
      L1.Pt1.Y = Network.Shape(ActID).Point(i).y
      L1.Pt2.X = Network.Shape(ActID).Point(i + 1).x
      L1.Pt2.Y = Network.Shape(ActID).Point(i + 1).y

      'intersectFound = False
      If Geo_Function.Point_Line(Point4Split.x, Point4Split.y, L1) = True Then
        If Highlight = True Then
          'DrawingOptions     
          g_MW.View.Draw.DrawPoint(Point4Split.x, Point4Split.y, 12, Drawing.Color.Black)
          g_MW.View.Draw.DrawPoint(Point4Split.x, Point4Split.y, 10, Drawing.Color.Magenta)
        End If
        'If Pipe1.numPoints = 0 Then
        PipePt = New MapWinGIS.Point
        PipePt.x = L1.Pt1.X
        PipePt.y = L1.Pt1.Y
        Pipe1.InsertPoint(PipePt, iVertex)
        iVertex += 1
        'End If
        PipePt = New MapWinGIS.Point
        PipePt.x = Point4Split.x
        PipePt.y = Point4Split.y
        Pipe1.InsertPoint(PipePt, iVertex)
        iVertex += 1
        FirstPipe = False

        'New pipe

        iVertex = 0
        Pipe2.InsertPoint(PipePt, iVertex)
        iVertex += 1
        PipePt = New MapWinGIS.Point
        PipePt.x = Point4Split.x
        PipePt.y = Point4Split.y
        'PipePt.x = L1.Pt2.X
        'PipePt.y = L1.Pt2.Y
        Pipe2.InsertPoint(PipePt, iVertex)
        iVertex += 1
        intersectFound = True
        haveIntersec = True
      Else
        If FirstPipe = True Then
          PipePt = New MapWinGIS.Point
          PipePt.x = L1.Pt1.X
          PipePt.y = L1.Pt1.Y
          Pipe1.InsertPoint(PipePt, iVertex)
          iVertex += 1
        Else
          PipePt = New MapWinGIS.Point
          PipePt.x = L1.Pt1.X
          PipePt.y = L1.Pt1.Y
          Pipe2.InsertPoint(PipePt, iVertex)
          iVertex += 1
        End If
      End If
    Next

    If haveIntersec = True Then
      If FirstPipe = True Then
        PipePt = New MapWinGIS.Point
        PipePt.x = L1.Pt2.X
        PipePt.y = L1.Pt2.Y
        Pipe1.InsertPoint(PipePt, iVertex)
        iVertex += 1
      Else
        PipePt = New MapWinGIS.Point
        PipePt.x = L1.Pt2.X
        PipePt.y = L1.Pt2.Y
        Pipe2.InsertPoint(PipePt, iVertex)
        iVertex += 1
      End If
    End If
    If Highlight = True Then
      For i As Integer = 0 To Pipe1.numPoints - 2
        g_MW.View.Draw.DrawLine(Pipe1.Point(i).x, Pipe1.Point(i).y, Pipe1.Point(i + 1).x, Pipe1.Point(i + 1).y, 4, Drawing.Color.Magenta)
      Next
      For i As Integer = 0 To Pipe2.numPoints - 2
        g_MW.View.Draw.DrawLine(Pipe2.Point(i).x, Pipe2.Point(i).y, Pipe2.Point(i + 1).x, Pipe2.Point(i + 1).y, 4, Drawing.Color.Green)
      Next
    End If
    Return haveIntersec
  End Function

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

  Function pointCompare(ByVal pt1 As MapWinGIS.Point, ByVal pt2 As MapWinGIS.Point, ByVal AllowError As Double) As Boolean
    pointCompare = False
    If ((pt1.x - pt2.x) ^ 2 + (pt1.y - pt2.y) ^ 2) ^ 0.5 < AllowError Then pointCompare = True
  End Function

  Public Function shapeTouch(ByVal shp1 As MapWinGIS.Shape, ByVal shp2 As MapWinGIS.Shape) As Boolean
    Dim i1, i2, j1, j2 As Integer
    Dim x1, y1 As Double
    shapeTouch = False
    shapeTouch = shp1.Touches(shp2)
    For i As Integer = 0 To shp1.numPoints - 1
      i1 = i
      i2 = i + 1
      If i = shp1.numPoints - 1 Then i2 = 0
      Dim L1 As Geo_Function.LineType
      L1.Pt1.X = shp1.Point(i1).x
      L1.Pt1.Y = shp1.Point(i1).y
      L1.Pt2.X = shp1.Point(i2).x
      L1.Pt2.Y = shp1.Point(i2).y

      For j As Integer = 0 To shp2.numPoints - 1
        j1 = j
        j2 = j + 1
        If j = shp2.numPoints - 1 Then j2 = 0
        Dim L2 As Geo_Function.LineType
        L2.Pt1.X = shp2.Point(j1).x
        L2.Pt1.Y = shp2.Point(j1).y
        L2.Pt2.X = shp2.Point(j2).x
        L2.Pt2.Y = shp2.Point(j2).y

        If Geo_Function.Line_Line(L1, L2, x1, y1) = True Then
          shapeTouch = True
          'LineGIS(L1, 2, Drawing.Color.Aquamarine)
          'LineGIS(L2, 2, Drawing.Color.Aquamarine)
        End If
      Next
    Next
  End Function

 
  Private Sub CreateShape()
    Dim shape As New MapWinGIS.Shape()
    Dim success As Boolean
    'Create a new Point shape
    success = shape.Create(MapWinGIS.ShpfileType.SHP_POINT)
    ' InsertPoint()
    Dim point As New MapWinGIS.Point()
    Dim pointindex As Integer
    'Set the values for the point to be inserted
    point.x = 100
    point.y = 100
    'Set the desired point index for the point to be inserted
    pointindex = 0
    Dim sf As New MapWinGIS.Shapefile()
    'Insert the point into the shape
    success = shape.InsertPoint(point, pointindex)
    success = shape.Create(MapWinGIS.ShpfileType.SHP_POINT)

    Dim shapeindex As Integer
    'Set the shape index
    shapeindex = 0
    'Switch the shapefile into editing mode
    sf.StartEditingShapes()

    ' ShapefileEditTable()
    'Set the shapefile attribute table to be in editing mode
    'success = sf.StartEditingTable(Me)

    'Insert the shape into the shapefile at index 0 if available
    sf.EditInsertShape(shape, shapeindex)

    ' ShapefileStopEditingTable()
    'Stop editing the shapefile attribute table, saving changes to the attribute table
    'success = sf.StopEditingTable(True, Me)
    ' ShapefileStopEditingShapes()

    'Stop editing shapes in the shapefile, saving changes to shapes, also stopping editing of the attribute table
    success = sf.StopEditingShapes(True, True, Me)
    'success = sf.SaveAs("C:\test.shp", Me)
  End Sub

  Function GetShp(ByVal LayerName As String) As MapWinGIS.Shapefile
    Dim iLyr As Integer = getLayerHandle(LayerName)
    If iLyr = -1 Then Return Nothing
    Dim shp As MapWinGIS.Shapefile
    shp = g_MW.Layers(iLyr).GetObject
    Return shp
  End Function

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

  Function Split3Pipe2(ByVal Network As MapWinGIS.Shapefile, ByVal ActID As Integer, ByRef Pipe1 As MapWinGIS.Shape, ByRef Pipe2 As MapWinGIS.Shape, _
                        ByRef Pipe3 As MapWinGIS.Shape, ByVal Highlight As Boolean, ByVal lineindex As Integer, ByVal x As Double, ByVal y As Double) As Boolean

    Dim L1 As Geo_Function.LineType
    Dim C1 As Geo_Function.CircleType
    Dim length As Double = 50000
    Dim haveIntersec As Boolean = False
    Dim pt As Geo_Function.POINTAPI
    Dim interIndex As Integer = lineindex
    Dim firstCircle As Integer = -1
    Dim lastCircle As Integer = -1

    pt.X = x
    pt.Y = y

    ' Result Line
    Dim LN1 As Geo_Function.LineType
    Dim LN2 As Geo_Function.LineType
    Dim LN3 As Geo_Function.LineType
    If interIndex <> -1 Then
      Dim dist As Double = 0
      Dim ptToZeroPoint As Double = length
      Dim ptToLastPoint As Double = length
      Dim i As Integer = interIndex
      Dim iVertex As Integer = 0
      Dim tmpX, tmpY As Double
      'MsgBox("1. ptToZeroPoint = " & ptToZeroPoint & "i = " & i)
      While i > -1 And ptToZeroPoint > 0 ' loop ตั้งแต่ จุดตัด ย้อนไปจนถึง point แรกของเส้น
        If i = interIndex Then
          L1.Pt1.X = pt.X
          L1.Pt1.Y = pt.Y
          L1.Pt2.X = Network.Shape(ActID).Point(i).x
          L1.Pt2.Y = Network.Shape(ActID).Point(i).y
        Else
          L1.Pt1.X = Network.Shape(ActID).Point(i + 1).x
          L1.Pt1.Y = Network.Shape(ActID).Point(i + 1).y
          L1.Pt2.X = Network.Shape(ActID).Point(i).x
          L1.Pt2.Y = Network.Shape(ActID).Point(i).y
        End If
        firstCircle += 1
        ptToZeroPoint -= Geo_Function.Length(L1)
        'MsgBox("2. ptToZeroPoint = " & ptToZeroPoint & "i = " & i)
        i -= 1
        If ptToZeroPoint < 0 Then
          Dim tmpLength As Double = ptToZeroPoint + Geo_Function.Length(L1)
          g_MW.View.Draw.DrawCircle(L1.Pt1.X, L1.Pt1.Y, tmpLength, Drawing.Color.Orange, False)
          C1.pt.X = L1.Pt1.X
          C1.pt.Y = L1.Pt1.Y
          C1.R = tmpLength
          Dim tx1, ty1, tx2, ty2 As Double
          Geo_Function.Line_Circle(L1, C1, tx1, ty1, tx2, ty2)
          If Geo_Function.Point_Line(tx1, ty1, L1) Then
            tmpX = tx1
            tmpY = ty1
          ElseIf Geo_Function.Point_Line(tx2, ty2, L1) Then
            tmpX = tx2
            tmpY = ty2
          End If
          LN1.Pt1.X = Network.Shape(ActID).Point(0).x
          LN1.Pt1.Y = Network.Shape(ActID).Point(0).y
          LN1.Pt2.X = tmpX
          LN1.Pt2.Y = tmpY
          LN2.Pt1.X = tmpX
          LN2.Pt1.Y = tmpY
        End If
      End While

      i = interIndex
      While i < Network.Shape(ActID).numPoints - 1 And ptToLastPoint > 0 'loop ตั้งแต่ จุดตัดไปยัง point สุดท้ายของเส้น
        If i = interIndex Then
          L1.Pt1.X = pt.X
          L1.Pt1.Y = pt.Y
        Else
          L1.Pt1.X = Network.Shape(ActID).Point(i).x
          L1.Pt1.Y = Network.Shape(ActID).Point(i).y
        End If
        L1.Pt2.X = Network.Shape(ActID).Point(i + 1).x
        L1.Pt2.Y = Network.Shape(ActID).Point(i + 1).y
        lastCircle += 1
        ptToLastPoint -= Geo_Function.Length(L1)
        'MsgBox("3. ptToLastPoint = " & ptToLastPoint & "i = " & i)
        i += 1
        If ptToLastPoint < 0 Then
          Dim tmpLength As Double = ptToLastPoint + Geo_Function.Length(L1)
          g_MW.View.Draw.DrawCircle(L1.Pt1.X, L1.Pt1.Y, tmpLength, Drawing.Color.PowderBlue, False)
          C1.pt.X = L1.Pt1.X
          C1.pt.Y = L1.Pt1.Y
          C1.R = tmpLength
          Dim tx1, ty1, tx2, ty2 As Double
          Geo_Function.Line_Circle(L1, C1, tx1, ty1, tx2, ty2)
          If Geo_Function.Point_Line(tx1, ty1, L1) Then
            tmpX = tx1
            tmpY = ty1
          ElseIf Geo_Function.Point_Line(tx2, ty2, L1) Then
            tmpX = tx2
            tmpY = ty2
          End If
          LN3.Pt1.X = tmpX
          LN3.Pt1.Y = tmpY
          LN3.Pt2.X = Network.Shape(ActID).Point(Network.Shape(ActID).numPoints - 1).x
          LN3.Pt2.Y = Network.Shape(ActID).Point(Network.Shape(ActID).numPoints - 1).y
          LN2.Pt2.X = tmpX
          LN2.Pt2.Y = tmpY
        End If
      End While

      Pipe1.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
      Pipe2.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
      Pipe3.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE

      Dim pIndex1 As Integer = 0
      Dim PipePt As New MapWinGIS.Point

      ' pipe1 
      i = 0
      While i <= interIndex 'pt first 
        PipePt = New MapWinGIS.Point
        PipePt.x = Network.Shape(ActID).Point(pIndex1).x
        PipePt.y = Network.Shape(ActID).Point(pIndex1).y
        Pipe1.InsertPoint(PipePt, pIndex1)
        pIndex1 += 1
        i += 1
      End While
      If firstCircle < 1 Then
        PipePt = New MapWinGIS.Point
      End If
      PipePt.x = LN1.Pt2.X
      PipePt.y = LN1.Pt2.Y
      Pipe1.InsertPoint(PipePt, pIndex1) 'byRef

      'pipe3
      PipePt = New MapWinGIS.Point
      PipePt.x = LN3.Pt1.X
      PipePt.y = LN3.Pt1.Y
      Pipe3.InsertPoint(PipePt, 0) 'byRef
      i = interIndex + 1
      pIndex1 = 1
      If lastCircle > 0 Then
        i += 1
      End If
      While i <= Network.Shape(ActID).numPoints - 1 'pt first 
        PipePt = New MapWinGIS.Point
        PipePt.x = Network.Shape(ActID).Point(i).x
        PipePt.y = Network.Shape(ActID).Point(i).y
        Pipe3.InsertPoint(PipePt, pIndex1)
        pIndex1 += 1
        i += 1
      End While

      'pipe2
      i = 0
      pIndex1 = 0
      While i < (firstCircle + lastCircle + 2)
        If i = 0 Then
          PipePt = New MapWinGIS.Point
          PipePt.x = LN2.Pt1.X
          PipePt.y = LN2.Pt1.Y
          Pipe2.InsertPoint(PipePt, pIndex1)
        ElseIf i = firstCircle + lastCircle + 1 Then
          PipePt = New MapWinGIS.Point
          PipePt.x = LN2.Pt2.X
          PipePt.y = LN2.Pt2.Y
          Pipe2.InsertPoint(PipePt, pIndex1)
        Else
          PipePt = New MapWinGIS.Point
          If firstCircle < lastCircle Then
            PipePt.x = Network.Shape(ActID).Point(interIndex + i).x
            PipePt.y = Network.Shape(ActID).Point(interIndex + i).y
          Else
            PipePt.x = Network.Shape(ActID).Point(interIndex + i - 1).x
            PipePt.y = Network.Shape(ActID).Point(interIndex + i - 1).y
          End If
          Pipe2.InsertPoint(PipePt, pIndex1)
        End If
        pIndex1 += 1
        i += 1
      End While

      If ptToZeroPoint > 0 Or ptToLastPoint > 0 Then
        MsgBox("ไม่สามารถสร้างท่อใหม่ได้ เนื่องจากท่อที่สรา้งใหม่ มีความยาวเกินท่อเดิม หรือ ท่อมีจุดต่อเกิน 1 จุด")
      Else
        If Highlight = True Then
          For i = 0 To Pipe1.numPoints - 2
            g_MW.View.Draw.DrawLine(Pipe1.Point(i).x, Pipe1.Point(i).y, Pipe1.Point(i + 1).x, Pipe1.Point(i + 1).y, 4, Drawing.Color.Magenta)
          Next

          For i = 0 To Pipe3.numPoints - 2
            g_MW.View.Draw.DrawLine(Pipe3.Point(i).x, Pipe3.Point(i).y, Pipe3.Point(i + 1).x, Pipe3.Point(i + 1).y, 4, Drawing.Color.OrangeRed)
          Next

          For i = 0 To Pipe2.numPoints - 2
            g_MW.View.Draw.DrawLine(Pipe2.Point(i).x, Pipe2.Point(i).y, Pipe2.Point(i + 1).x, Pipe2.Point(i + 1).y, 4, Drawing.Color.Green)
          Next
        End If
      End If
    Else
      haveIntersec = True
    End If
    Return haveIntersec
  End Function

  Sub addShapeAttribute(ByRef Shp As MapWinGIS.Shapefile, ByVal SourceID As Integer, ByVal TargetID As Integer)
    Shp.StartEditingTable()
    For i As Integer = 0 To Shp.NumFields - 1
      Shp.EditCellValue(i, TargetID, Shp.CellValue(i, SourceID))
    Next
    Shp.StopEditingTable()

    Shp.Save()
  End Sub

  Function ReDirection(ByVal PipeShp As MapWinGIS.Shape) As MapWinGIS.Shape
    Dim tmpReDirection As New MapWinGIS.Shape
    tmpReDirection.ShapeType = MapWinGIS.ShpfileType.SHP_POLYLINE
    For i As Integer = 0 To PipeShp.numPoints - 1
      Dim PipePt = New MapWinGIS.Point
      PipePt.x = PipeShp.Point(PipeShp.numPoints - 1 - i).x
      PipePt.y = PipeShp.Point(PipeShp.numPoints - 1 - i).y
      tmpReDirection.InsertPoint(PipePt, i)
    Next
    Return tmpReDirection
  End Function

#End Region

#Region "2. Network generation TAB"

  Private Sub cmdAutoGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAutoGen.Click
    grdNode.Rows.Clear()
    grdLink.Rows.Clear()
    '--------------------------------------------------------------------------------
    grdNode.GridColor = Color.Blue
    grdNode.CellBorderStyle = DataGridViewCellBorderStyle.None
    grdNode.BackgroundColor = Color.LightGray

    grdNode.DefaultCellStyle.SelectionBackColor = Color.Aquamarine
    grdNode.DefaultCellStyle.SelectionForeColor = Color.Blue

    grdNode.DefaultCellStyle.WrapMode = DataGridViewTriState.[True]

    grdNode.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    grdNode.AllowUserToResizeColumns = False

    grdNode.RowsDefaultCellStyle.BackColor = Color.AliceBlue
    grdNode.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
    '--------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------
    grdLink.GridColor = Color.Blue
    grdLink.CellBorderStyle = DataGridViewCellBorderStyle.None
    grdLink.BackgroundColor = Color.LightGray

    grdLink.DefaultCellStyle.SelectionBackColor = Color.Aquamarine
    grdLink.DefaultCellStyle.SelectionForeColor = Color.Blue

    grdLink.DefaultCellStyle.WrapMode = DataGridViewTriState.[True]

    grdLink.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    grdLink.AllowUserToResizeColumns = False

    grdLink.RowsDefaultCellStyle.BackColor = Color.AliceBlue
    grdLink.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
    '--------------------------------------------------------------------------------

    Dim HashNode As New Hashtable
    Dim NODEid As Integer = 1 ' Start Node at node no. 1
    Dim NetworkShp As New MapWinGIS.Shapefile
    Dim DEM As New MapWinGIS.Grid
    Dim hashString As String
    Dim NodeF, NodeT As Integer
    NetworkShp = g_MW.Layers(getLayerHandle(cmbNetwork.Text)).GetObject
    DEM.Open(g_MW.Layers(getLayerHandle(cmbDEM.Text)).FileName)
    clearGraphic()

    UpdateProgress.Maximum = NetworkShp.NumShapes - 1
    UpdateProgress.Minimum = 0
    UpdateProgress.Visible = True
    For i As Integer = 0 To NetworkShp.NumShapes - 1
      NodeT = -1
      NodeF = -1
      UpdateProgress.Value = i
      Dim pt1 As New MapWinGIS.Point
      Dim pt2 As New MapWinGIS.Point
      Dim LastPt As Integer = NetworkShp.Shape(i).numPoints - 1
      pt1 = NetworkShp.Shape(i).Point(0)
      pt2 = NetworkShp.Shape(i).Point(LastPt)
      '----------------------------------------------
      ' Round coordinate value
      '
      ''NetworkShp.StartEditingShapes()
      'NetworkShp.Shape(i).Point(0).x = Math.Round(pt1.x, 3)
      'NetworkShp.Shape(i).Point(0).y = Math.Round(pt1.y, 3)
      'NetworkShp.Shape(i).Point(LastPt).x = Math.Round(pt2.x, 3)
      'NetworkShp.Shape(i).Point(LastPt).y = Math.Round(pt2.y, 3)
      ''NetworkShp.StopEditingShapes()
      '----------------------------------------------
      pt1 = NetworkShp.Shape(i).Point(0)
      pt2 = NetworkShp.Shape(i).Point(LastPt)
      'Check the first Node 
      hashString = Math.Round(pt1.x, 2) & "," & Math.Round(pt1.y, 2)
      Try
        HashNode.Add(hashString, NODEid)
        NODEid += 1
        NodeF = HashNode(hashString)
        grdNode.Rows.Add()
        Dim ActRow As Integer = grdNode.RowCount - 2
        grdNode.Item(0, ActRow).Value = NodeF
        grdNode.Item(1, ActRow).Value = pt1.x
        grdNode.Item(2, ActRow).Value = pt1.y
        Dim ColDem, RowDEM As Integer
        DEM.ProjToCell(pt1.x, pt1.y, ColDem, RowDEM)
        grdNode.Item(5, ActRow).Value = DEM.Value(ColDem, RowDEM)

      Catch ex As Exception
        NodeF = HashNode(hashString)
      End Try

      If NodeF = -1 Then
        Debug.Print("Pipe id " & i)
      End If
      'Check the second Node 
      hashString = Math.Round(pt2.x, 2) & "," & Math.Round(pt2.y, 2)
      Try
        HashNode.Add(hashString, NODEid)
        NODEid += 1
        NodeT = HashNode(hashString)
        grdNode.Rows.Add()
        Dim ActRow As Integer = grdNode.RowCount - 2
        grdNode.Item(0, ActRow).Value = NodeT
        grdNode.Item(1, ActRow).Value = pt2.x
        grdNode.Item(2, ActRow).Value = pt2.y
        Dim ColDem, RowDEM As Integer
        DEM.ProjToCell(pt2.x, pt2.y, ColDem, RowDEM)
        grdNode.Item(5, ActRow).Value = DEM.Value(ColDem, RowDEM)

        'g_MW.View.Draw.AddDrawingLabel(0, NodeT, Drawing.Color.Black, pt2.x, pt2.y, MapWinGIS.tkHJustification.hjCenter)
      Catch ex As Exception
        NodeT = HashNode(hashString)
      End Try

      If NodeT = -1 Then
        Debug.Print("Pipe id " & i)
      End If

      grdLink.Rows.Add()
      '//grdLink.Item(0, i).Value = NetworkShp.CellValue(PipeIDField, i) ' Link ID
      grdLink.Item(0, i).Value = NetworkShp.CellValue(FieldNetwork.LinkID, i) ' Link ID
      grdLink.Item(1, i).Value = NodeF
      grdLink.Item(2, i).Value = NodeT
      grdLink.Item(3, i).Value = Format(Math.Round(NetworkShp.Shape(i).Length, 3), "0.000")
      grdLink.Item(4, i).Value = NetworkShp.CellValue(FieldNetwork.PIPE_SIZE, i) ' Link ID
    Next
    UpdateProgress.Visible = False
    MsgBox("complete")
  End Sub


  Private Sub cmdSelectNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectNode.Click
    TabControl1.SelectedTab = TabPage6
    ActiveDMACommand = DMACommand.SelectNode
    g_MW.View.CursorMode = MapWinGIS.tkCursorMode.cmNone
  End Sub  'Click on node grid

#End Region

#Region "3. Network Editor TAB"

#Region "  3.1 Network checking process"

  Private Sub cmdOpenUnLinkFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenUnLinkFile.Click
    Dim fopen As New OpenFileDialog
    fopen.Filter = "Unconneceted node file(*.Unconnode)|*.Unconnode"
    fopen.ShowDialog()
    If fopen.FileName <> "" Then
      Dim fx As Integer = FreeFile()
      FileOpen(fx, fopen.FileName, OpenMode.Input)
      Dim i As Integer = 0
      While Not EOF(fx)
        Dim st As String = LineInput(fx)
        Dim a As Object = Split(st, ",")
        grdUnconnectedNode.Rows.Add()
        grdUnconnectedNode.Item(0, i).Value = a(0)
        grdUnconnectedNode.Item(1, i).Value = a(1)
        i += 1
      End While
      FileClose(fx)
    End If
  End Sub

  Private Sub cmdSaveUnLinkFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveUnLinkFile.Click

  End Sub

    Private Sub cmdClearDrawing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearDrawing1.Click
        clearGraphic()
    End Sub

 
#End Region

#Region "  3.2 Pipe extend"

    Private Sub cmdSelectExtendSide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ActivePipeEditCommand = PipeEditCommand.select1stPipe
        'g_MW.View.ShapeDrawingMethod = MapWinGIS.tkShapeDrawingMethod.dmStandard
        clearGraphic()
        If Not tmp1stPipe Is Nothing Then
            drawShape(tmp1stPipe, Drawing.Color.Red)
            Dim endPt As Integer = tmp1stPipe.numPoints - 1
            tmp1stPipe = ReDirection(tmp1stPipe)
            g_MW.View.Draw.DrawPoint(tmp1stPipe.Point(endPt).x, tmp1stPipe.Point(endPt).y, 10, Color.Red)
        End If
        If Not tmp2ndPipe Is Nothing Then drawShape(tmp2ndPipe, Drawing.Color.Green)

    End Sub

    Private Sub cmdSelectExtendPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim endPt As Integer = tmp1stPipe.numPoints - 1
        Dim minR As Double = 100000000000
        Dim xLink, yLink As Double
        For j As Integer = 0 To tmp2ndPipe.numPoints - 2
            Dim j1 As Integer = j
            Dim j2 As Integer = j + 1
            Dim x1, y1, x2, y2, xx, yy As Double
            x1 = tmp2ndPipe.Point(j1).x
            y1 = tmp2ndPipe.Point(j1).y
            x2 = tmp2ndPipe.Point(j2).x
            y2 = tmp2ndPipe.Point(j2).y
            Dim DistanceFromPt As Double = Geo_Function.DistToSegment(tmp1stPipe.Point(endPt).x, tmp1stPipe.Point(endPt).y, x1, y1, x2, y2, xx, yy)
            If DistanceFromPt < minR Then
                minR = DistanceFromPt
                xLink = xx
                yLink = yy
            End If
        Next
        g_MW.View.Draw.DrawLine(tmp1stPipe.Point(endPt).x, tmp1stPipe.Point(endPt).y, xLink, yLink, 2, Color.Magenta)
        g_MW.View.Draw.DrawPoint(tmp1stPipe.Point(endPt).x, tmp1stPipe.Point(endPt).y, 10, Color.Magenta)
        g_MW.View.Draw.DrawPoint(xLink, yLink, 10, Color.Magenta)

    End Sub

#End Region

#End Region

#Region "Common function"

    Function ChekFieldID(ByVal shp As MapWinGIS.Shapefile, ByVal FieldName As String, ByRef FieldID As Integer) As Boolean
        For i As Integer = 0 To shp.NumFields - 1
            If shp.Field(i).Name.ToUpper = FieldName.ToUpper Then
                FieldID = i
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub cmdPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPath.Click
        Dim Network As New MapWinGIS.Shapefile
        Network = g_MW.Layers(getLayerHandle(cmbNetwork.Text)).GetObject


        Dim TestPath As PathProperties
        Dim O, D As Integer
        Main_SP.Main()
        O = CInt(txtOrigin.Text)
        D = CInt(txtDestination.Text)
        Dim F, T As Integer
        Dim x1, x2, y1, y2 As Double

        'AxMap1.ClearDrawings()
        Dim col As UInteger = System.Convert.ToUInt32(RGB(255, 255, 255))

        TestPath = createSPaths(FlowPathHyd, PipeNetwork, O, D)
        ListBox1.Items.Clear()
        For i As Integer = TestPath.nPt To 1 Step -1
            ListBox1.Items.Add(TestPath.pNode(i).id)
        Next
        ListBox1.Refresh()
        clearGraphic()
        For i As Integer = 1 To TestPath.nPt - 1
            F = TestPath.pNode(i).id
            T = TestPath.pNode(i + 1).id
            Dim pt1, pt2 As POINTAPI
            pt1 = getNodeCoord(F)
            pt2 = getNodeCoord(T)
            Dim ShapeID As Integer = getLinkId(F, T)
            If Not (Network.Shape(ShapeID) Is Nothing) Then
                drawShape(Network.Shape(ShapeID), Drawing.Color.Red)
            Else
                Dim xxx
                xxx = 1
            End If
            'x1 = pt1.X
            'y1 = pt1.Y
            'x2 = pt2.X
            'y2 = pt2.Y
            'g_MW.View.Draw.DrawLine(x1, y1, x2, y2, 3, Drawing.Color.AliceBlue)
        Next
    End Sub

#End Region

    Private Sub grdLink_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdLink.CellContentClick
        Try
            Dim RowId As Integer = e.RowIndex
            clearGraphic()
            Dim pipeShp As New MapWinGIS.Shapefile
            Dim pipeLyr As Integer = getLayerHandle(cmbNetwork.Text)
            pipeShp = g_MW.Layers(pipeLyr).GetObject
            Dim pipeIDField As Integer = getColNum("LinkID", pipeShp)

            For i As Integer = 0 To pipeShp.NumShapes - 1
                If sender.Item(0, RowId).Value = pipeShp.CellValue(pipeIDField, i) Then
                    drawShape(pipeShp.Shape(i), Color.Red, 3)
                    g_MW.View.Extents = pipeShp.Shape(i).Extents
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub

  Private Sub cmdInitialNET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInitialNET.Click
    Dim x1, y1, x2, y2, x3, y3 As Double
    Dim NodeF, NodeT As Integer
    clearGraphic()
    nNode = grdNode.RowCount - 1
    HashNodeID.Clear()
    ReDim Node(grdNode.RowCount - 1)
    UpdateProgress.Value = 0
    UpdateProgress.Minimum = 0
    UpdateProgress.Maximum = grdNode.RowCount - 1
    For i As Integer = 0 To grdNode.RowCount - 1
      UpdateProgress.Value = i
      Node(i + 1).x = grdNode.Item(1, i).Value
      Node(i + 1).y = grdNode.Item(1, i).Value
      Node(i + 1).id = i
      HashNodeID.Add(Node(i + 1).id, i)
      g_MW.View.Draw.DrawPoint(Node(i + 1).x, Node(i + 1).y, 3, Drawing.Color.Blue)
    Next


    nLink = grdLink.RowCount - 1
    ReDim Link(nLink)
    HashLinkID.Clear()
    UpdateProgress.Maximum = nLink
    UpdateProgress.Minimum = 0
    UpdateProgress.Visible = True

    For i As Integer = 0 To grdLink.RowCount - 1
      UpdateProgress.Value = i
      Dim F As Integer = grdLink.Item(1, i).Value
      Dim T As Integer = grdLink.Item(2, i).Value

      x1 = Node(F).x
      y1 = Node(F).y
      x1 = Node(T).x
      y1 = Node(T).y
      Link(i + 1).NodeF = F
      Link(i + 1).NodeT = T
      Link(i + 1).id = i
      Link(i + 1).L = grdLink.Item(2, i).Value
      g_MW.View.Draw.DrawLine(x1, y1, x2, y2, 1, Drawing.Color.Red)
      'drawShape(NetworkShp.Shape(i), Drawing.Color.Red)
      Try
        '--------------------------------------------------------
        ' Create HashTable for Link ID and Node ID
        '
        HashLinkID.Add(F & "-" & T, i)
        HashLinkID.Add(T & "-" & F, i)
      Catch ex As Exception
      End Try
      'Debug.Print("'" & F & " - " & T & " , " & ShpID)
    Next
    'For i As Integer = 0 To NetworkShp.NumShapes - 1
    '   Dim F As Integer = Link(i + 1).NodeF
    '   Dim T As Integer = Link(i + 1).NodeT
    '   grdLink.Item(4, i).Value = Link(i + 1).id 'HashLinkID(T & "-" & F)
    'Next

    UpdateProgress.Visible = False
    MsgBox("Network was created")
    cmdNetworkCheck.Enabled = True
  End Sub
End Class
