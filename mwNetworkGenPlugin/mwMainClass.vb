Imports System.IO

Public Class mwLayoutEditorClass
    Implements MapWindow.Interfaces.IPlugin

    Private Const netGIS01Button1 As String = "COPAM GIS"


    Private m_Button As MapWindow.Interfaces.ToolbarButton
    Private m_MapWindowForm As Windows.Forms.Form

    Public ReadOnly Property Author() As String Implements MapWindow.Interfaces.IPlugin.Author
        Get
            Return "Andre Daccache"
        End Get
    End Property
    Public ReadOnly Property Name() As String Implements MapWindow.Interfaces.IPlugin.Name
        Get
            Return "Copam"
        End Get
    End Property

    Public ReadOnly Property version() As String Implements MapWindow.Interfaces.IPlugin.Version
        Get
            Return "V1.0"
        End Get
    End Property
   
   Public ReadOnly Property BuildDate() As String Implements MapWindow.Interfaces.IPlugin.BuildDate
      Get
         Dim myDate As String
         Try
                myDate = Date.Now
         Catch
                myDate = Date.Now
         End Try
         Return myDate
      End Get
   End Property

   Public ReadOnly Property Description() As String Implements MapWindow.Interfaces.IPlugin.Description
      Get
            Return "COPAM GIS"
      End Get
   End Property

   <CLSCompliant(False)> _
   Public Sub Initialize(ByVal MapWin As MapWindow.Interfaces.IMapWin, ByVal ParentHandle As Integer) Implements MapWindow.Interfaces.IPlugin.Initialize


        mPublics.g_MW = MapWin
      m_MapWindowForm = System.Windows.Forms.Control.FromHandle(New IntPtr(ParentHandle))

   
      ''--------------------------------------------------------------------
      '' MAPWINDOW 4.8.3
      ''--------------------------------------------------------------------
        Dim ico As New System.Drawing.Bitmap(System.Windows.Forms.Application.StartupPath & "\Plugins\copam\IAMB.ico") ' (Me.GetType, "ToolbarIconNew.ico")
        m_Button = MapWin.Toolbar.AddButton(netGIS01Button1, ico)
  
      If m_Button Is Nothing Then
         MapWinUtility.Logger.Msg("Error initializing the COPAM Plug-in!", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Error")
      End If

      If Not g_MW Is Nothing Then
         If g_MW.Layers.NumLayers > 0 AndAlso LayerIsShapefile(g_MW.Layers.CurrentLayer) Then
            m_Button.Enabled = True
         Else
            m_Button.Enabled = True
         End If
      End If
        m_Button.Tooltip = "COPAM"
    m_Button.Enabled = True
   End Sub

   Public Sub ItemClicked(ByVal ItemName As String, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.ItemClicked
    If ItemName = netGIS01Button1 AndAlso LayerIsShapefile(g_MW.Layers.CurrentLayer) Then
      Handled = True
      mPublics.frmNetworkGen = New Form1 'frmNetGen '
          With mPublics.frmNetworkGen
                .Show()
            End With
    End If
  End Sub

   Public Sub LayerRemoved(ByVal Handle As Integer) Implements MapWindow.Interfaces.IPlugin.LayerRemoved
      If g_MW.Layers.NumLayers > 0 AndAlso LayerIsShapefile(g_MW.Layers.CurrentLayer) Then
         m_Button.Enabled = True
      Else
         m_Button.Enabled = True
      End If
   End Sub

   <CLSCompliant(False)> _
   Public Sub LayersAdded(ByVal Layers() As MapWindow.Interfaces.Layer) Implements MapWindow.Interfaces.IPlugin.LayersAdded
      LoadPluginDialog(Layers(0).Handle)
   End Sub

   Public Sub LayersCleared() Implements MapWindow.Interfaces.IPlugin.LayersCleared
      ' m_Button.Enabled = False
   End Sub

   Public Sub LayerSelected(ByVal Handle As Integer) Implements MapWindow.Interfaces.IPlugin.LayerSelected
      LoadPluginDialog(Handle)
   End Sub

   <CLSCompliant(False)> _
   Public Sub LegendDoubleClick(ByVal Handle As Integer, ByVal Location As MapWindow.Interfaces.ClickLocation, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.LegendDoubleClick

   End Sub

   <CLSCompliant(False)> _
   Public Sub LegendMouseUp(ByVal Handle As Integer, ByVal Button As Integer, ByVal Location As MapWindow.Interfaces.ClickLocation, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.LegendMouseUp

   End Sub

   <CLSCompliant(False)> _
   Public Sub LegendMouseDown(ByVal Handle As Integer, ByVal Button As Integer, ByVal Location As MapWindow.Interfaces.ClickLocation, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.LegendMouseDown

   End Sub

   Public Sub MapDragFinished(ByVal Bounds As System.Drawing.Rectangle, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.MapDragFinished

   End Sub

   Public Sub Message(ByVal msg As String, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.Message
      If msg = "DataGettingStart" Then
         Handled = True
         If LayerIsShapefile(g_MW.Layers.CurrentLayer) Then
            'mPublics.LayoutEditor = New frmBlockDesign(CType(g_MW.Layers(g_MW.Layers.CurrentLayer).GetObject, MapWinGIS.Shapefile), m_MapWindowForm)
            ' mPublics.pwaGIS01Editor = New frmDmaDesign '(CType(g_MW.Layers(g_MW.Layers.CurrentLayer).GetObject, MapWinGIS.Shapefile), m_MapWindowForm)
                With mPublics.frmNetworkGen
                    .Show()
                End With
         Else
            MapWinUtility.Logger.Msg("This layer is not a shapefile", MsgBoxStyle.Information)
         End If
      End If
   End Sub

   Public Sub MapExtentsChanged() Implements MapWindow.Interfaces.IPlugin.MapExtentsChanged

   End Sub

   Public Sub MapMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.MapMouseDown
      Dim xx, yy As Double
      Dim x1, y1, x2, y2 As Double
      g_MW.View.PixelToProj(x, y, xx, yy)

    End Sub

   Public Sub MapMouseMove(ByVal ScreenX As Integer, ByVal ScreenY As Integer, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.MapMouseMove
      Dim xx, yy, x1, x2, y1, y2 As Double
      g_MW.View.PixelToProj(ScreenX, ScreenY, xx, yy)
    End Sub

   Public Sub MapMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.MapMouseUp

   End Sub



   Public Sub ProjectLoading(ByVal ProjectFile As String, ByVal SettingsString As String) Implements MapWindow.Interfaces.IPlugin.ProjectLoading

   End Sub

   Public Sub ProjectSaving(ByVal ProjectFile As String, ByRef SettingsString As String) Implements MapWindow.Interfaces.IPlugin.ProjectSaving

   End Sub

   Public ReadOnly Property SerialNumber() As String Implements MapWindow.Interfaces.IPlugin.SerialNumber
      Get
         Return "OCGHOFNASNKGD1CAAAAKRIT"
      End Get
   End Property

   <CLSCompliant(False)> _
   Public Sub ShapesSelected(ByVal Handle As Integer, ByVal SelectInfo As MapWindow.Interfaces.SelectInfo) Implements MapWindow.Interfaces.IPlugin.ShapesSelected
    End Sub

   Public Sub Terminate() Implements MapWindow.Interfaces.IPlugin.Terminate
    g_MW.Toolbar.RemoveButton(netGIS01Button1)
   End Sub


   Private Function LayerIsShapefile(ByVal Handle As Integer) As Boolean
      If g_MW Is Nothing Then Return False
      If Handle < 0 Then Return False
      If Not g_MW.Layers.IsValidHandle(Handle) Then Return False
      Select Case g_MW.Layers(Handle).LayerType
         Case MapWindow.Interfaces.eLayerType.LineShapefile, MapWindow.Interfaces.eLayerType.PointShapefile, MapWindow.Interfaces.eLayerType.PolygonShapefile
            Return True
         Case Else
            Return False
      End Select
   End Function

   Private Sub LoadPluginDialog(ByVal Handle As Integer)
        'If g_MW Is Nothing Then
        '   MsgBox("The PWA function can only be used from MapWindow or if g_MW has been set (directly or with Initialize()).", MsgBoxStyle.Exclamation, "Unavailable Functionality")
        '   Return
        'End If

      'If LayerIsShapefile(Handle) Then
      '    If Not mPublics.pwaGIS01Editor Is Nothing Then
      '        ' the table editor is already open, reload it
      '        ' mPublics.pwaGIS01Editor.Initialize(CType(g_MW.Layers(g_MW.Layers.CurrentLayer).GetObject, MapWinGIS.Shapefile), m_MapWindowForm)
      '    End If
      'Else
      '    If Not mPublics.pwaGIS01Editor Is Nothing Then
      '        mPublics.pwaGIS01Editor.Hide()
      '        mPublics.pwaGIS01Editor.Close()
      '    End If
      '    m_Button.Enabled = True
      'End If

   End Sub

  

#Region "Tools"

   Private Function getShpID(ByVal shp As MapWinGIS.Shapefile, ByVal pt As MapWinGIS.Point, Optional ByVal ExtSize As Double = 0.1) As Integer
      Dim ShapeID() As Int32

      Dim ext As New MapWinGIS.Extents
      ext.SetBounds(pt.x - ExtSize, pt.y - ExtSize, 0, pt.x + ExtSize, pt.y + ExtSize, 0)
      shp.SelectShapes(ext, 0, MapWinGIS.SelectMode.INTERSECTION, ShapeID)

      If UBound(ShapeID) >= 0 Then
         getShpID = ShapeID(0)
      Else
         'getShpID = -1
         MapWinGeoProc.SpatialOperations.GetShapeNearestToPoint(shp, pt, getShpID, ExtSize)
      End If
   End Function

   Private Sub drawShape(ByVal shp As MapWinGIS.Shape, ByVal color As System.Drawing.Color)
      g_MW.View.Draw.NewDrawing(MapWinGIS.tkDrawReferenceList.dlSpatiallyReferencedList)
      For i As Integer = 0 To shp.numPoints - 2
         Dim x1 As Double = shp.Point(i).x
         Dim y1 As Double = shp.Point(i).y
         Dim x2 As Double = shp.Point(i + 1).x
         Dim y2 As Double = shp.Point(i + 1).y
         g_MW.View.Draw.DrawLine(x1, y1, x2, y2, 3, color)
      Next
   End Sub

#End Region

End Class
