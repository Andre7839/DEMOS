Option Strict Off
Option Explicit On

'Imports System.Runtime.InteropServices

Friend Class clsFlowPath
   ' STANDARD FlowPathHyd's ALGORITHM USING A PRIORITY QUEUE IMPLEMENTED AS A BINARY HEAP
   '
   Private Const dInf As Double = 1.0E+308 '"Infinity".

   Private blListInit As Boolean 'List has been initialized.

   Private lGNodes As Integer 'Number of nodes (vertices) in the graph.
   Private lGEdges As Integer 'Number of edges (arcs) in the graph (max.=lGNodes^2).
   Private lGConnections As Integer 'Number of Connections in the graph (max.=2*lGEdges).

   Private lHeapSize As Integer 'Number of elements in the priority queue (max.=lGNodes).

   Private ListPtr() As Integer 'Holds each node's UBound position in the Adjacency List.
   Private ListAdjNode() As Integer 'For each node, it holds the ID's of the linked nodes.
   Private ListAdjCost() As Double 'For each node, it holds the connection cost of the linked nodes.
   Private dCost() As Double 'Cost of travelling to each node. Array Index refers to node ID.
   Private lPredecessor() As Integer 'Keeps track of the node ID via which we rach each destination node. Array Index refers to destination's ID.
   Private HeapID() As Integer 'Node ID of each element in the priority queue. Array index refers to HeapPos()
   Private HeapPos() As Integer 'Position of each node in the priority queue.

   'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
   'Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Long, ByVal Length As Integer)
   '<DllImport("kernel32.dll")> _
   '  Public Shared Sub ZeroMemory(ByRef Destination As Long, ByVal Length As Integer)
   'End Sub

   '++----------------------------------------------------------------------------
   '//   Public Properties and simple Methods
   '++----------------------------------------------------------------------------

   '//Number of Nodes in the adjacency list.
   Public ReadOnly Property GNodes() As Integer
      Get
         GNodes = lGNodes
      End Get
   End Property

   '//Number of Edges in the adjacency list.
   Public ReadOnly Property GEdges() As Integer
      Get
         GEdges = lGEdges
      End Get
   End Property

   '//Number of Connections in the adjacency list.
   Public ReadOnly Property GConnections() As Integer
      Get
         GConnections = lGConnections
      End Get
   End Property

   '//Get the precalculated Predecessor's ID on the path from Source to lDestination.
   Public ReadOnly Property Predecessor(ByVal lDestination As Integer) As Integer
      Get
         Predecessor = lPredecessor(lDestination)
      End Get
   End Property

   '//Get the precalculated Cost value from Source to lDestination.
   Public ReadOnly Property Cost(ByVal lDestination As Integer) As Double
      Get
         Cost = dCost(lDestination)
      End Get
   End Property

   '//Get the array of Predecessors' IDs on the path from Source to all other nodes
   ' (index refers to Destination ID).
   Public Sub GetPredecessorsArray(ByRef TheArray() As Integer)
      'TheArray = VB6.CopyArray(lPredecessor)
      TheArray = lPredecessor
   End Sub

   '//Get the array of Costs from Source to all other nodes
   ' (index refers to Destination ID).
   Public Sub GetCostArray(ByRef TheArray() As Double)
      'TheArray = VB6.CopyArray(dCost)
      TheArray = dCost
   End Sub

   '//Get the ratio of the connections traversed during last Flow Paths calculation
   Public Function GraphExplored() As Double
      Dim lExplored As Integer 'Counter of explored connections

      If blListInit And lGConnections Then '/Adjacency List is Initialised
         ' and at least one connection exists.
         For i As Integer = 1 To lGNodes '/For every node in the graph..
            If dCost(i) <> dInf Then '/If a path from the source exists...
               lExplored = lExplored + ListPtr(i) - ListPtr(i - 1) '/add nr. of links whose origin is i.
            End If
         Next i
         GraphExplored = lExplored / lGConnections 'Ratio of explored VS total connections.
      End If
   End Function


   '++/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
   '//
   '//   ADJACENCY LIST STRUCTURE (sort of a linked-list implementation)
   '//
   '++\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

   '++----------------------------------------------------------------------------
   '//   ADJACENCY LIST INITIALIZATION
   '++----------------------------------------------------------------------------
   '//   -The adjacency list is a way of representing the network topology:
   '//    For each node, it allows sequential access to all the nodes that are
   '//    directly linked to it, retrieving the IDs and respective connection costs.
   '//
   '//   -This routine must be run prior calculating any Flow Paths.
   '//    It is also the only routine in this class that requires an external
   '//    object reference (PipeNetwork). Instead of PipeNetwork, you could also pass only
   '//    the Cost and Topology arrays, in order to easily include the class
   '//    in an other application / library.
   '++----------------------------------------------------------------------------

   Public Function InitList(ByRef PipeNetwork As clsNetwork, Optional ByRef blPrune As Boolean = False) As String
      Dim i As Integer
      Dim lF As Integer 'Temp variable, refers to TopoFrom()
      Dim lT As Integer 'Temp variable, refers to TopoTo()
      Dim lNodeLBound() As Integer 'This is a temp pointer for node insertion.
      Dim CostFromTo() As Double 'A copy of PipeNetwork's respective array.
      Dim CostToFrom() As Double 'A copy of PipeNetwork's respective array.
      Dim TopoFrom() As Integer 'A copy of PipeNetwork's respective array.
      Dim TopoTo() As Integer 'A copy of PipeNetwork's respective array.
      Dim DeleteNodes() As Short 'Used only if blPrune=True.
      Dim blHasDangles As Boolean '=True if PipeNetwork contains nodes whose connectivity is < 2 (could be any positive number).

      With PipeNetwork
         lGNodes = .GNodes '/Get the nr. of graph nodes.
         lGEdges = .GEdges '/Get the nr. of graph edges.
         lGConnections = .GConnections '/Get the nr. of graph connections.

         'UPGRADE_WARNING: Lower bound of array CostFromTo was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
         ReDim CostFromTo(lGEdges) '/Initialize the temp arrays to hold
         'UPGRADE_WARNING: Lower bound of array CostToFrom was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
         ReDim CostToFrom(lGEdges) ' PipeNetwork Data.
         'UPGRADE_WARNING: Lower bound of array TopoFrom was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
         ReDim TopoFrom(lGEdges) '
         'UPGRADE_WARNING: Lower bound of array TopoTo was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
         ReDim TopoTo(lGEdges) '

         .GetCost_FromTo(CostFromTo) '/Copy each temp array from PipeNetwork
         .GetCost_ToFrom(CostToFrom) ' to access the data more efficiently.
         .GetTopology_From(TopoFrom) '
         .GetTopology_To(TopoTo) '
      End With
      '/Initialize the 'pointer' array to the Adjacency List.
      ReDim ListPtr(lGNodes)

      For i = 1 To lGEdges '/Looping through PipeNetwork's edges, we
         If Not CostFromTo(i) Then ' count every nodes connectivity.
            ListPtr(TopoFrom(i)) = ListPtr(TopoFrom(i)) + 1 ' Edges Cost=-1 means the connection
         End If ' is impassable.
         If Not CostToFrom(i) Then '
            ListPtr(TopoTo(i)) = ListPtr(TopoTo(i)) + 1 '
         End If '
      Next i '

      If blPrune Then '/This optional part is useful if we
         'UPGRADE_WARNING: Lower bound of array DeleteNodes was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
         ReDim DeleteNodes(lGNodes) ' want to mark the nodes whose
         For i = 1 To lGNodes ' connectivity is less than a specified bound.
            If ListPtr(i) < 2 Then ' (Here, nodes with connectivity<2,
               blHasDangles = True ' i.e. belonging in sub-trees).
               DeleteNodes(i) = 1 '
            End If '
         Next i '
         If blHasDangles Then ' We remove nodes marked as Deleted
            PipeNetwork.FixGraph(DeleteNodes) ' in a recursive manner.

            Erase CostFromTo '//Update #1
            Erase CostToFrom '/   If we do not erase
            Erase TopoTo '/   the arrays, the recursion will
            Erase TopoFrom '/   consume more and more resources
            Erase ListPtr '/   before it is finished.
            Erase DeleteNodes '//End update

            InitList(PipeNetwork, blPrune) '
            Exit Function '
         End If '
      End If '

      'UPGRADE_WARNING: Lower bound of array ListAdjNode was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
      ReDim ListAdjNode(lGConnections) '/Initialization of the Adjacency List.
      'UPGRADE_WARNING: Lower bound of array ListAdjCost was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
      ReDim ListAdjCost(lGConnections) '
      'UPGRADE_WARNING: Lower bound of array lNodeLBound was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
      ReDim lNodeLBound(lGNodes) '/Initialization of the temp pointer array.

      For i = 1 To lGNodes '/We make the array of pointers incremetal
         ListPtr(i) = ListPtr(i) + ListPtr(i - 1) ' and initialise each node's Lower Bound
         lNodeLBound(i) = ListPtr(i - 1) ' for insertion in the List.
      Next i

      For i = 1 To lGEdges '/This is the construction of the actuall List.
         lF = TopoFrom(i) ' For each connection found, the respective
         lT = TopoTo(i) ' nodes' LBound is incremented by 1, and the
         If Not CostFromTo(i) Then ' adjacent node's ID and connection cost are
            lNodeLBound(lF) = lNodeLBound(lF) + 1 ' entered into that position.
            ListAdjNode(lNodeLBound(lF)) = lT '
            ListAdjCost(lNodeLBound(lF)) = CostFromTo(i) '
         End If '
         If Not CostToFrom(i) Then '
            lNodeLBound(lT) = lNodeLBound(lT) + 1 '
            ListAdjNode(lNodeLBound(lT)) = lF '
            ListAdjCost(lNodeLBound(lT)) = CostToFrom(i) '
         End If '
      Next i
      '
      '    for i As Integer = 1 To lGEdges
      '        lF = TopoFrom(i)
      '        lT = TopoTo(i)
      '        Debug.Print lF & " - " & lT
      '    Next

      'UPGRADE_WARNING: Lower bound of array dCost was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
      ReDim dCost(lGNodes) '/Initialization of FlowPathHyd and
      'UPGRADE_WARNING: Lower bound of array lPredecessor was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
      ReDim lPredecessor(lGNodes) ' Heap related arrays.
      'UPGRADE_WARNING: Lower bound of array HeapID was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
      ReDim HeapID(lGNodes) '
      'UPGRADE_WARNING: Lower bound of array HeapPos was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
      ReDim HeapPos(lGNodes) '

      blListInit = True 'We are ready to start computing Flow Paths!

   End Function

   '++/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
   '//   FlowPathHyd ALGORITHM USING A PRIORITY QUEUE IMPLEMENTED AS A BINARY HEAP
   '//
   '//   -Step 1: Initialization ~ O(lGNodes).
   '//   -Step 2: FindMin, DeleteRoot ~ O(log(lGNodes)).
   '//   -Step 3: Relax tentative costs ~ O(lLocalConnections*log(lGNodes)).
   '//   -Step 4: Repeat Step 2, Step 3 until Heap is empty or Destination Reached.
   '//
   '//   Amortised worst case complexity ~ O(lGEdges*log(lGNodes)).
   '++\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

   '++----------------------------------------------------------------------------
   '//   FlowPathHyd INITIALIZATION ~ complexity O(lGNodes)
   '++----------------------------------------------------------------------------
   '//   -What does it do?
   '//   -First we set the array of costs equal to "Infinity", and clear the
   '//    contents of the rest of the arrays used. Then we initialize the Heap by
   '//    inserting the Source node in the Root. Finally we call the main algorithm.
   '++----------------------------------------------------------------------------
   Public Function FlowPathHydHeap(ByRef Source As Integer, Optional ByRef Destination As Integer = 0) As Double
      Dim lNodes As Integer

      FlowPathHydHeap = -1 'In case something goes wrong.
      If blListInit Then
         If Source > 0 And Source <= lGNodes Then
            If Destination >= 0 And Destination <= lGNodes Then

               For lNodes = 1 To lGNodes '/Set cost array = Infinity.
                  dCost(lNodes) = dInf '
               Next lNodes '
               dCost(Source) = 0 '/Set the cost to the Source=0
               '(this could also be optional).

               lNodes = lGNodes * 4 '/Size of each array in bytes.
               For iZ As Integer = 1 To UBound(lPredecessor) 'lNodes
                  lPredecessor(iZ) = 0 'ZeroMemory(lPredecessor(1), lNodes) '/Clear the arrays.
                  HeapID(iZ) = 0       'ZeroMemory(HeapID(1), lNodes) '
                  HeapPos(iZ) = 0      'ZeroMemory(HeapPos(1), lNodes) '
               Next

               lHeapSize = 1 '/Initialization of the Heap
               HeapID(lHeapSize) = Source ' by inserting the Source
               HeapPos(Source) = lHeapSize ' as the root node.

               FlowPathHydHeap = FlowPathHydMain(Destination) '/Call the main algorithm.

            End If
         End If
      End If
   End Function

   '++----------------------------------------------------------------------------
   '//   MAIN ALGORITHM ~ complexity O(lGEdges*log(lGNodes))
   '++----------------------------------------------------------------------------
   '//   -What does it do?
   '//   -First, it finds the unscanned node closest to the Source. This is the
   '//    root of the Heap, HeapID(1). If this is the target destination, we stop,
   '//    else we delete the Root:
   '//    This is done by swapping the Root (which is now lLastScanned) with the
   '//    lLastElement of the heap. Then, we perform a Shift Down on the root since
   '//    it probably violates the heap property (ie. it's key [cost in our case]
   '//    is probably greater than one of it's children).
   '//    Finally, we try improving the costs found so far by considering the
   '//    connections that come out of lLastScanned node.
   '++----------------------------------------------------------------------------
   Private Function FlowPathHydMain(ByVal Destination As Integer) As Double
      Dim lLastScanned As Integer 'The (until now unscanned) node's ID which is closest to the Source.
      Dim lLastElement As Integer 'ID of the last element in the priority queue.

      Do  '//Main Loop.

         lLastScanned = HeapID(1) '/lLastScanned node's ID = Root ID.
         If lLastScanned = Destination Then '/Check if we have reached our
            FlowPathHydMain = dCost(Destination) ' Destination.
            Exit Do ' If Destination = 0
         End If ' we simply calculate all paths.

         lLastElement = HeapID(lHeapSize) '/We delete the Heap's (now scanned)
         HeapID(1) = lLastElement ' Root element, by swapping it with the
         HeapPos(lLastElement) = 1 ' Heap's lLastElement, decreasing the
         lHeapSize = lHeapSize - 1 ' HeapSize by 1
         ' and shifting the new Root down to
         HeapShiftDown(lLastElement) ' an appropriate position.

         FlowPathHydRelaxCosts(lLastScanned) '/We try improving the costs found so far.

      Loop While lHeapSize '//While the Heap is not empty.

      FlowPathHydMain = dCost(lLastScanned)
   End Function

   '++----------------------------------------------------------------------------
   '//   RELAX COSTS ~ complexity O((LastCon-FirstCon)*log(lGNodes))
   '++----------------------------------------------------------------------------
   '//   -What does it do?
   '//   -For each node adjacent to lLastScannedNode (lAdjNode), we try relaxing
   '//    it's so-far Best Known ("tentative") Cost (dCost(lAdjNode)).
   '//    If that was an improvement, we update dCost(lAdjNode), insert
   '//    lAdjNode in the Heap (if it wasn't already there) and Percolate lAdjNode
   '//    to it's final Heap Position.
   '//
   '//    We could also use an array to keep track of which nodes have reached
   '//    their optimal costs (i.e. Scanned), to prevent summing and comparing
   '//    double Vars. However, I found that the overhead caused by this, actually
   '//    slows the procedure down marginally, even for dense graphs.
   '++----------------------------------------------------------------------------
   Private Sub FlowPathHydRelaxCosts(ByVal lLastScanned As Integer)
      Dim i As Integer
      Dim FirstCon As Integer 'LBound in the Adjacency List.
      Dim LastCon As Integer 'Ubound in the Adjacency List.
      Dim lAdjNode As Integer 'ID of each of lLastScanned node's adjacent nodes.
      Dim dCostTemp As Double 'Optimal cost from Source node to lLastScanned node.
      Dim dLastScanned As Double 'Cost from Source node to lAdjNode node via lLastScanned node.

      dLastScanned = dCost(lLastScanned) '/These assignments should
      FirstCon = ListPtr(lLastScanned - 1) + 1 ' save some overhead later on.
      LastCon = ListPtr(lLastScanned) '

      For i = FirstCon To LastCon '/Check each of LastScannedNode's direct connections.

         lAdjNode = ListAdjNode(i) '/These assignments should
         dCostTemp = ListAdjCost(i) + dLastScanned ' save some overhead later on.

         If dCostTemp < dCost(lAdjNode) Then '/If we have found a shorter path...
            dCost(lAdjNode) = dCostTemp ' Update the current best cost.
            lPredecessor(lAdjNode) = lLastScanned ' Update the last node we have to reach
            ' prior reaching lAdjNode (for drawing purposes).

            If HeapPos(lAdjNode) Then '/lAdjNode is in the Heap:
               HeapPercolate(lAdjNode) ' Increase lAdjNode's Priority.

            Else '/lAdjNode is not in the Heap:
               lHeapSize = lHeapSize + 1 ' Increase the lHeapSize by 1.
               HeapID(lHeapSize) = lAdjNode ' Insert lAdjNode as the last element.
               HeapPos(lAdjNode) = lHeapSize '
               HeapPercolate(lAdjNode) ' Increase AdjNode's Priority.
            End If
         End If
      Next i
   End Sub

   '++----------------------------------------------------------------------------
   '//   PERCOLATE ~ complexity O(log(lHeapSize))
   '++----------------------------------------------------------------------------
   '//   -What does it do?
   '//   -It maintains the Heap property, i.e. ensuring that the target Child's
   '//    key (cost in our case) is greater than it's respective Parent's key.
   '//
   '//    If the child's key is smaller, we swap Parent and Child and continue
   '//    the process until no further swaps are possible.
   '//
   '//    In a binary Heap, ith Child's Parent is located at i\2
   '//
   '//    Also known as: Decrease Key, Increase Priority, Bubble Up or Shift Up.
   '++----------------------------------------------------------------------------
   Private Sub HeapPercolate(ByVal lChildID As Integer)
      Dim lParentID As Integer 'Parent's ID at each iteration.
      Dim lParentPos As Integer 'Parent's position in the Heap at each iteration.
      Dim lChildPos As Integer 'Child's position in the Heap at each iteration.
      Dim dChildCost As Double 'Child's updated key (cost from Source).

      dChildCost = dCost(lChildID)
      lChildPos = HeapPos(lChildID) '/Current Child's position.

      Do Until lChildPos = 1 '//Until we reach the Root...
         lParentPos = lChildPos \ 2 '/Parent's Current position.
         lParentID = HeapID(lParentPos) '/Parent's ID.

         If dCost(lParentID) > dChildCost Then '/Child is violating the Heap property.
            HeapID(lChildPos) = lParentID ' Update the HeapID at Child's position to hold Parent's ID.
            HeapPos(lParentID) = lChildPos ' Update Parent's position.
            lChildPos = lParentPos ' This is the Child's new position.
         Else '
            Exit Do '/Child is NOT violating the Heap property.
         End If '
      Loop  '//Try shifting the Child further UP the Heap.

      HeapID(lChildPos) = lChildID '/Establish the Child in it's final position.
      HeapPos(lChildID) = lChildPos '
   End Sub

   '++----------------------------------------------------------------------------
   '//   SHIFT DOWN ~ complexity O(log(lHeapSize))
   '++----------------------------------------------------------------------------
   '//   -What does it do?
   '//   -It maintains the Heap property, i.e. ensuring that the target parent's
   '//    key (cost in our case) is smaller than it's respective children keys.
   '//
   '//    If the parent's key is greater, we swap Parent with the Child that has the
   '//    smallest key and continue the process until no further swaps are possible.
   '//
   '//    In a binary Heap, ith Parent's Left and Right Children are located
   '//    at i*2 and i*2+1 respectively (here we do Left=i+i and Right=Left+1).
   '//
   '//    Also known as: Increase Key, Decrease Priority or Bubble Down.
   '++----------------------------------------------------------------------------
   Private Sub HeapShiftDown(ByVal lParentID As Integer)
      Dim lChildID As Integer 'Selected Child's ID at each iteration.
      Dim lParentPos As Integer 'Parent's position in the Heap at each iteration.
      Dim lLeftPos As Integer 'Left Child's position in the Heap at each iteration.
      Dim lRightPos As Integer 'Right Child's position in the Heap at each iteration.
      Dim dLeftCost As Double 'Left Child's key (cost from Source).
      Dim dParentCost As Double 'Parent's key (cost from Source).

      dParentCost = dCost(lParentID)
      lParentPos = HeapPos(lParentID) '/Current Parent's position.
      lLeftPos = lParentPos + lParentPos '/Current Left Child's position.

      Do Until lLeftPos > lHeapSize '//Until Left position is out of bounds...
         lRightPos = lLeftPos + 1 '/Current Right Child's position.
         dLeftCost = dCost(HeapID(lLeftPos)) '/Current Left Child's key.

         If lRightPos <= lHeapSize Then '//If Right position is inside bounds...

            If dLeftCost < dParentCost Then '/Parent is violating the Heap property.
               If dLeftCost > dCost(HeapID(lRightPos)) Then ' We consider as "Left" Child the Child
                  lLeftPos = lRightPos ' with the minimum key.
               End If '
            Else '
               If dParentCost > dCost(HeapID(lRightPos)) Then '/Parent is violating the Heap property.
                  lLeftPos = lRightPos ' We consider as "Left" Child the Child
                  ' with the minimum key.

               Else '/Parent is NOT violating the Heap property.
                  Exit Do '
               End If
            End If
         Else '//If Right position is outside bounds...
            If dParentCost <= dLeftCost Then '/Parent is NOT violating the Heap property.
               Exit Do '
            End If
         End If
         '/Perform the swap.
         lChildID = HeapID(lLeftPos) 'Child's ID (could actually be Left's or Right's)
         HeapID(lParentPos) = lChildID 'Update the HeapID at Parent's position to hold Child's ID.
         HeapPos(lChildID) = lParentPos 'Update Child's position.

         lParentPos = lLeftPos 'Current Parent's position.
         lLeftPos = lLeftPos + lLeftPos 'Current Left Child's position.
      Loop  '//Try shifting the Parent further DOWN the Heap.

      HeapID(lParentPos) = lParentID '/Establish the Parent in it's final position.
      HeapPos(lParentID) = lParentPos '
   End Sub
End Class