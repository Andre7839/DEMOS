Option Strict Off
Option Explicit On
'Imports VB = Microsoft.VisualBasic

Module PublicVar_Sp

   Public Const dInf As Double = 1.0E+308 '"Infinity"
   Public Const sInf As Single = 1.0E+38 '/"Infinity" (single)
   Public Const lProgress_ms As Integer = 250 'How often should we update the progressbar (ms)

   Public GlobalFormHeight As Integer
   Public GlobalFormWidth As Integer
   Public blMouseMoveCalc As Boolean '=True if we are calculating Flow Paths via the mouse
   Public blAllowTileDiag As Boolean 'Only for tile instances - toggles diagonals.
   Public blSPTrees As Boolean 'Show Flow Path trees instead of a single path
   Public blPruneGraph As Boolean 'see clsFlowPathHyd InitList
   Public blBenchMark As Boolean 'Udate #1: =True if we are running the benchmark

   Public blStopProcess As Boolean 'Stop calculating (used by frmProgress)
   Public blUpdate As Boolean 'update the progressbar or not
   Public lElapsedPrevious As Integer 'elapsed ms until last progressbar update

   Private Const NON_RELATIVE As Boolean = False 'Refers to imported graphs (see bellow)
   Private Const RELATIVE As Boolean = True 'Refers to imported graphs (see bellow)
   Public iOrigin As Integer
   Public iDetinetion As Integer
   Public iBenchmark As Integer
   '++----------------------------------------------------------------------------
   '//   CALCULATE Flow path
   '++----------------------------------------------------------------------------
   'Public Function CalcShortestPaths(ByRef FlowPathHyd As clsFlowPathHyd, ByRef SimpleDraw As clsSimpleDraw, ByRef PipeNetwork As clsGraph, Optional ByVal Source As Integer = 0, Optional ByVal Destination As Integer = 0, Optional ByVal lAPSP As Integer = 0, Optional ByVal blCoord As Boolean = False) As Double
   Public Function CalcShortestPaths(ByRef FlowPathHyd As clsFlowPath, ByRef PipeNetwork As clsNetwork, Optional ByVal Source As Integer = 0, Optional ByVal Destination As Integer = 0, Optional ByVal lAPSP As Integer = 0, Optional ByVal blCoord As Boolean = False) As Double
      Dim i As Integer
      Dim j As Integer
      Dim lGNodes As Integer
      Dim lGEdges As Integer
      Dim dTemp As Double
      Dim Predecessors() As Integer
      Dim ST As String = ""
      With FlowPathHyd

         lGNodes = .GNodes
         lGEdges = .GEdges ' Update #1: Only for the benchmark (used in the progressbar)
         '--------------------------------
         'Simple Flow Path calculation
         '--------------------------------

         If Source And Source <> Destination Then

            If Not blMouseMoveCalc Then
               System.Windows.Forms.Application.DoEvents()
            End If

            j = iBenchmark 'Int(frmMDI.txtMDI(3))

            ReDim Predecessors(lGNodes)
            For i = 1 To j
               dTemp = .FlowPathHydHeap(Source, Destination)
            Next i
            .GetPredecessorsArray(Predecessors)
            ST = iDetinetion & "-" & iOrigin & ": "
            If Destination Then
               While Predecessors(Destination)
                  ST += CStr(Destination) & ","
                  Destination = Predecessors(Destination)
               End While
            End If
            ST += CStr(Destination)
            PrintLine(2, ST)
         End If
      End With

   End Function

   Public Sub getPipeNET(ByRef PipeNetwork As clsNetwork)
      With PipeNetwork
         For i As Integer = 1 To nLink
            .NewEdge(Link(i).NodeF, Link(i).NodeT, Link(i).L, Link(i).L)
         Next
         .TrimBuffers()
      End With
   End Sub

   Public Function createSPaths(ByRef FlowPathHyd As clsFlowPath, ByRef PipeNetwork As clsNetwork, ByVal Source As Integer, ByVal Destination As Integer) As PathProperties
      Try

         Dim i As Integer
         Dim j As Integer
         Dim lGNodes As Integer
         Dim lGEdges As Integer
         Dim dTemp As Double
         Dim Predecessors() As Integer
         Dim ST As String = ""
         Dim mpath As PathProperties

         With FlowPathHyd

            lGNodes = .GNodes
            lGEdges = .GEdges ' Update #1: Only for the benchmark (used in the progressbar)
            '--------------------------------
            'Simple Flow Path calculation
            '--------------------------------

            If Source And Source <> Destination Then

               j = iBenchmark 'Int(frmMDI.txtMDI(3))

               ReDim Predecessors(lGNodes)
               For i = 1 To j
                  dTemp = .FlowPathHydHeap(Source, Destination)
               Next i
               .GetPredecessorsArray(Predecessors)
               'ST = iDetinetion & "-" & iOrigin & ": "
               If Destination Then
                  While Predecessors(Destination)
                     'ST += CStr(Destination) & ","
                     mpath.nPt += 1
                     ReDim Preserve mpath.pNode(mpath.nPt)
                     mpath.pNode(mpath.nPt).id = Destination
                     Destination = Predecessors(Destination)

                  End While
               End If
               mpath.nPt += 1
               ReDim Preserve mpath.pNode(mpath.nPt)
               mpath.pNode(mpath.nPt).id = Destination

               'ST += CStr(Destination)
               'PrintLine(2, ST)
            End If
         End With
         Return mpath
      Catch ex As Exception
            MsgBox("No link was found")
      End Try

   End Function

End Module