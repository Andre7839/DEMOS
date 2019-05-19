Option Strict Off
Option Explicit On
Friend Class clsNetwork
    ' Pipe Network Class
    '
    Private Enum BufferConst
        UPDT_BUF = 0 ' Update the Edges potential UBound
        INIT_BUF = 1 ' Initialize the Edges
        TRIM_BUF = 2 ' Trim the Edges arrays to the actual Edges UBound
    End Enum

    Private Const lInitStep As Integer = 256 'Default buffer step
    Private lRedimStep As Integer 'This is the number of Edges "slots" we buffer each time the buffer is full.

    Private lGNodes As Integer 'Number of nodes in the graph (n)
    Private lGEdges As Integer 'Number of edges in the graph (m) (0<= m <=n^2)
    Private lGConnections As Integer 'Number of connections in the graph (c) (0<= c <=2*m)
    Private blInitXY As Boolean '=True, If the graph nodes are associated with XY coordinates

    Private Coords_X() As Single 'Node X coordinates. Index refers to node ID
    Private Coords_Y() As Single 'Node Y coordinates. Index refers to node ID
    Private Topology_From() As Integer 'Edge Start, or "From", node. Index refers to edge ID
    Private Topology_To() As Integer 'Edge End, or "To", node. Index refers to edge ID
    Private Cost_FromTo() As Double 'Cost of connection From -> To. Index refers to edge ID
    Private Cost_ToFrom() As Double 'Cost of connection To -> From. Index refers to edge ID

    '++----------------------------------------------------------------------------
    '//   Class
    '++----------------------------------------------------------------------------

    Public Sub Class_Initialize_Renamed()
        RedimStep = lInitStep
        lGNodes = 0
        lGEdges = 0
        blInitXY = False
        BufferEdges(BufferConst.INIT_BUF)
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    '++----------------------------------------------------------------------------
    '//   Public Properties and simple Methods
    '++----------------------------------------------------------------------------

    '----------------------------------
    '// Get Network properties one by one
    '----------------------------------

    Public ReadOnly Property CostFromTo(ByVal EdgeID As Integer) As Double
        Get
            If EdgeID <= lGEdges Then
                CostFromTo = Cost_FromTo(EdgeID)
            End If
        End Get
    End Property

    Public ReadOnly Property CostToFrom(ByVal EdgeID As Integer) As Double
        Get
            If EdgeID <= lGEdges Then
                CostToFrom = Cost_ToFrom(EdgeID)
            End If
        End Get
    End Property

    Public ReadOnly Property EdgeFrom(ByVal EdgeID As Integer) As Integer
        Get
            If EdgeID <= lGEdges Then
                EdgeFrom = Topology_From(EdgeID)
            End If
        End Get
    End Property

    Public ReadOnly Property EdgeTo(ByVal EdgeID As Integer) As Integer
        Get
            If EdgeID <= lGEdges Then
                EdgeTo = Topology_To(EdgeID)
            End If
        End Get
    End Property

    Public ReadOnly Property GConnections() As Integer
        Get
            GConnections = lGConnections
        End Get
    End Property

    Public ReadOnly Property GEdges() As Integer
        Get
            GEdges = lGEdges
        End Get
    End Property

    Public ReadOnly Property GNodes() As Integer
        Get
            GNodes = lGNodes
        End Get
    End Property

    Public ReadOnly Property NodeXcoord(ByVal NodeID As Integer) As Single
        Get
            If NodeID <= lGNodes Then
                NodeXcoord = Coords_X(NodeID)
            End If
        End Get
    End Property

    Public ReadOnly Property NodeYcoord(ByVal NodeID As Integer) As Single
        Get
            If NodeID <= lGNodes Then
                NodeYcoord = Coords_Y(NodeID)
            End If
        End Get
    End Property


    Public Property RedimStep() As Integer
        Get
            RedimStep = lRedimStep
        End Get
        Set(ByVal Value As Integer)
            lRedimStep = IIf(Value, System.Math.Abs(Value), lInitStep)
            If lGEdges = 0 Then
                BufferEdges(BufferConst.INIT_BUF)
            End If
        End Set
    End Property

    '------------------------------------
    '// Get complete arrays of properties
    '------------------------------------

    Public Sub GetCoords_X(ByRef TheArray() As Single)
        'TheArray = VB6.CopyArray(Coords_X)
        TheArray = Coords_X
    End Sub

    Public Sub GetCoords_Y(ByRef TheArray() As Single)
        'TheArray = VB6.CopyArray(Coords_Y)
        TheArray = Coords_Y
    End Sub

    Public Sub GetCost_FromTo(ByRef TheArray() As Double)
        'TheArray = VB6.CopyArray(Cost_FromTo)
        TheArray = Cost_FromTo
    End Sub

    Public Sub GetCost_ToFrom(ByRef TheArray() As Double)
        'TheArray = VB6.CopyArray(Cost_ToFrom)
        TheArray = Cost_ToFrom
    End Sub

    Public Sub GetTopology_From(ByRef TheArray() As Integer)
        'TheArray = VB6.CopyArray(Topology_From)
        TheArray = Topology_From
    End Sub

    Public Sub GetTopology_To(ByRef TheArray() As Integer)
        'TheArray = VB6.CopyArray(Topology_To)
        TheArray = Topology_To
    End Sub



    '++----------------------------------------------------------------------------
    '//   BUFFER EDGES
    '++----------------------------------------------------------------------------
    '//   - It is used to handle how frequently the edge arrays are redimensioned.
    '++----------------------------------------------------------------------------
    Public Sub TrimBuffers()
        BufferEdges(BufferConst.TRIM_BUF) '/Partial interface to BufferEdges
    End Sub

    Private Sub BufferEdges(Optional ByVal lOp As BufferConst = BufferConst.UPDT_BUF)
        Dim i As Integer

        If lOp = BufferConst.INIT_BUF Then '/Initialize
            ReDim Topology_From(lRedimStep)
            ReDim Topology_To(lRedimStep)
            ReDim Cost_FromTo(lRedimStep)
            ReDim Cost_ToFrom(lRedimStep)
        Else
            If lOp = BufferConst.UPDT_BUF Then
                i = UBound(Topology_From) + lRedimStep '/Update
            ElseIf lOp = BufferConst.TRIM_BUF Then
                i = lGEdges '/Trim
            End If
            ReDim Preserve Topology_From(i)
            ReDim Preserve Topology_To(i)
            ReDim Preserve Cost_FromTo(i)
            ReDim Preserve Cost_ToFrom(i)
        End If

    End Sub

    '++----------------------------------------------------------------------------
    '//   ADD A NEW EDGE
    '++----------------------------------------------------------------------------
    '//   - Adds a new edge to the graph, redimensioning the arrays if needed.
    '//   - This routine also maintains an upper bound in the number of nodes.
    '//   - Connection Cost = - 1 means impassable.
    '++----------------------------------------------------------------------------
    Public Sub NewEdge(ByVal FromNode As Integer, ByVal ToNode As Integer, ByVal CostFromTo As Double, ByVal CostToFrom As Double)

        lGEdges = lGEdges + 1
        If Not CostFromTo Then '/Not impassable.
            lGConnections = lGConnections + 1 ' increment connections.
        End If
        If Not CostToFrom Then '/Not impassable.
            lGConnections = lGConnections + 1 ' increment connections.
        End If

        If lGEdges Mod lRedimStep = 0 Then '/We must redimension
            BufferEdges()
        End If

        Topology_From(lGEdges) = FromNode '/Set the edge properties
        Topology_To(lGEdges) = ToNode '
        Cost_FromTo(lGEdges) = CostFromTo '
        Cost_ToFrom(lGEdges) = CostToFrom '

        If lGNodes < FromNode Then '/Update the number of nodes
            lGNodes = FromNode '
        End If '
        If lGNodes < ToNode Then '
            lGNodes = ToNode '
        End If '

    End Sub

    '++----------------------------------------------------------------------------
    '//   UPDATE AN EDGE
    '++----------------------------------------------------------------------------
    '//   - Updates an edge's properties, adjusting the connection count appropriately
    '++----------------------------------------------------------------------------
    Public Sub UpdateEdge(ByVal EdgeID As Integer, Optional ByVal FromNode As Integer = 0, Optional ByVal ToNode As Integer = 0, Optional ByVal CostFromTo As Double = -2, Optional ByVal CostToFrom As Double = -2)

        If FromNode Then
            Topology_From(EdgeID) = FromNode '/Update "From" Node ID
        End If

        If ToNode Then
            Topology_To(EdgeID) = ToNode '/Update "To" Node ID
        End If

        If CostFromTo <> -2 Then '/Update FromTo costs & connection count
            If (Not CostFromTo) And Cost_FromTo(EdgeID) = -1 Then
                lGConnections = lGConnections + 1 'Previous=impassable, new=passable
            End If
            If CostFromTo = -1 And (Not Cost_FromTo(EdgeID)) Then
                lGConnections = lGConnections - 1 'Previous=passable, new=impassable
            End If
            Cost_FromTo(EdgeID) = CostFromTo 'New FromTo connection cost
        End If

        If CostToFrom <> -2 Then '/Update ToFrom costs & connection count
            If (Not CostToFrom) And Cost_ToFrom(EdgeID) = -1 Then
                lGConnections = lGConnections + 1 'Previous=impassable, new=passable
            End If
            If CostToFrom = -1 And (Not Cost_ToFrom(EdgeID)) Then
                lGConnections = lGConnections - 1 'Previous=passable, new=impassable
            End If
            Cost_ToFrom(EdgeID) = CostToFrom 'New ToFrom connection cost
        End If

    End Sub

    '++----------------------------------------------------------------------------
    '//   DELETE AN EDGE
    '++----------------------------------------------------------------------------
    '//   - Swaps deleted edge with the last edge and adjust connection and edge count appropriately
    '++----------------------------------------------------------------------------
    Public Sub DeleteEdge(ByVal EdgeID As Integer)

        If Not Cost_FromTo(EdgeID) Then '/The connection existed (was passable)
            lGConnections = lGConnections - 1
        End If
        If Not Cost_ToFrom(EdgeID) Then '/The connection existed (was passable)
            lGConnections = lGConnections - 1
        End If
        If EdgeID < lGEdges Then '/Swap with last edge
            Topology_From(EdgeID) = Topology_From(lGEdges) ' (Note: we are not storing edges
            Topology_To(EdgeID) = Topology_To(lGEdges) '  in increasing ID order - in
            Cost_FromTo(EdgeID) = Cost_FromTo(lGEdges) '  contrast with the nodes.)
            Cost_ToFrom(EdgeID) = Cost_ToFrom(lGEdges) '
        End If '
        lGEdges = lGEdges - 1

    End Sub

    '++----------------------------------------------------------------------------
    '//   ADD NODES / UPDATE NODE'S XY COORDINATES
    '++----------------------------------------------------------------------------
    '//   - It also ReDim Preserve the array of nodes, if we exceed the UBound.
    '//     '(This shouldn't happen if we load the edges first. It is useful however,
    '//       if we want to add new nodes later, e.g. interactively on screen).
    '++----------------------------------------------------------------------------
    Public Sub SetNodeCoordinates(ByVal NodeID As Integer, ByVal Xcoord As Single, ByVal Ycoord As Single)
        If lGNodes < NodeID Then
            lGNodes = NodeID
        End If
        If Not blInitXY Then
            ReDim Coords_X(lGNodes) '/Initialize the arrays
            ReDim Coords_Y(lGNodes)
            blInitXY = True
        Else
            If UBound(Coords_X) < lGNodes Then
                ReDim Preserve Coords_X(lGNodes) '/New nodes have been added
                ReDim Preserve Coords_Y(lGNodes)
            End If
        End If
        Coords_X(NodeID) = Xcoord
        Coords_Y(NodeID) = Ycoord
    End Sub

    '++----------------------------------------------------------------------------
    '//   FIX THE GRAPH - (READ: BULK DELETE NODES AND ASSOCIATED EDGES AND FIX ID MESS :) )
    '++----------------------------------------------------------------------------
    '//   - NOTE: This routine is rather inefficient (it should use an adjacency list for speed up).
    '//     In addition, it potentially messes up with the ID's initially imported.
    '//
    '//   - It finds and deletes all edges that have at least one of their endpoint nodes marked
    '//     as deleted (<> 0). Nodes marked as deleted are deleted as well (duh!..).
    '//     It then assigns new IDs to the remaining nodes in order for them to be in the new
    '//     (1,...,lGNodes) interval and updates all edges with the new node IDs.
    '//
    '//     This means that the new node with ID=12345 is probably different than the
    '//     node with the same ID before this routine was applied.
    '++----------------------------------------------------------------------------
    Public Sub FixGraph(ByRef DeleteNodes() As Short)
        Dim i As Integer
        Dim j As Integer
        Dim MapNewNodes() As Integer
        ReDim MapNewNodes(lGNodes)
        ReDim Preserve DeleteNodes(lGNodes)
        i = 1
        j = 0
        Do
            If DeleteNodes(Topology_From(i)) Or DeleteNodes(Topology_To(i)) Then
                DeleteEdge((i))
            Else
                i = i + 1
            End If
        Loop Until i > lGEdges
        BufferEdges(BufferConst.TRIM_BUF)
        j = 0
        For i = 1 To lGNodes
            If DeleteNodes(i) = 0 Then
                j = j + 1
                MapNewNodes(i) = j
                Coords_X(j) = Coords_X(i)
                Coords_Y(j) = Coords_Y(i)
            End If
        Next i
        lGNodes = j
        ReDim Preserve Coords_X(lGNodes)
        ReDim Preserve Coords_Y(lGNodes)
        'Debug.Print Coords_X(j) & ", " & Coords_Y(j)
        For i = 1 To lGEdges
            UpdateEdge(i, MapNewNodes(Topology_From(i)), MapNewNodes(Topology_To(i)))
        Next i
    End Sub

End Class