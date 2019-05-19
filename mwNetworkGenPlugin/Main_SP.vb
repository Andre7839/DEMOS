Option Strict Off
Option Explicit On

Module Main_SP
   Private MyFile As String 'The last file imported

   Private blInit As Boolean '=True if a graph has been initialized
   Private blCoord As Boolean '=True if node coordinates exist
   Private blIsRunning As Boolean '=True if algorithm is running
   Private blZoomStart As Boolean
   Public FlowPathHyd As clsFlowPath
   Public PipeNetwork As clsNetwork
   Public OutFile As String
   Public NetFile As String
   'Public nPAth As Short
   '
   'Main Program
   '
   Public Sub Main() '(InputFile As String)
      '
      ' initial Data
      '
      'initial()

      Dim st As String
      iBenchmark = 1
      ''--------------------------------------------------------------
      PipeNetwork = New clsNetwork
      FlowPathHyd = New clsFlowPath

      blInit = False
      blCoord = False

      getPipeNET(PipeNetwork)
      FlowPathHyd.InitList(PipeNetwork, blPruneGraph)

      blCoord = True
      blInit = True
      '--------------------------------------------------------------
    End Sub


    Public Const gram As Double = 9.806
    Public Const g As Double = 9.806
    Public maxQ As Double = -100000000000
    Public minQ As Double = 100000000000
    Public LoadNetwork As Boolean = False
    Public HashLinkID As New Hashtable()
    Public HashNodeID As New Hashtable()

    Public Structure NodalPathProperties
        Dim id As Integer
        Dim x As Double
        Dim y As Double
        Dim z As Double
        Dim kd As Double
        Dim Type As String
        Dim Condition As String
    End Structure

    Public Structure PathProperties
        Dim nPt As Integer
        Dim pNode() As NodalPathProperties
    End Structure

    Public Structure NodeProperties
        Dim id As Integer
        Dim x As Double
        Dim y As Double
        Dim z As Double
        Dim h As Double
        Dim Q As Double
        Dim A As Double
        'Dim kd As Double
        Dim Type As String
        'Dim Condition As String
        'Dim hf As Double
        'Dim used As Boolean
        'Dim BEndVal As Integer
        'Dim UsedCount As Integer
    End Structure

    Public Structure LoopProperties
        Dim id As Integer
        Dim NodeT As Integer
        Dim NodeF As Integer
        Dim L As Double
        Dim DH As Double
        Dim beta As Double
    End Structure

    Public Structure LinkProperties
        Dim id As Integer
        Dim NodeT As Integer
        Dim NodeF As Integer
        Dim L As Double
        Dim Discharge As Double
        Dim diam As Double
        Dim rough As Double

        'Dim diameter As Double
        'Dim Area As Double
        'Dim c As Double
        'Dim Velocity As Double
        'Dim J As Double
        'Dim AccumuUsed As Integer
    End Structure

    Public Structure pipecostproperties
        Dim price As Single
        Dim diam As Single
        Dim rough As Single
    End Structure

    Public nNode As Integer
    Public nLink As Integer
    Public nPath As Integer


    Public Node() As NodeProperties
    Public Link() As LinkProperties
    Public Path() As PathProperties
    Public Pipe_C() As pipecostproperties
    Public loop_C() As LoopProperties
    Public Newnode() As newNodeProperties

    Public Structure newNodeProperties
        Dim id As Integer
        Dim z As Double
        Dim q As Double
        Dim Fnode As Integer ' first node
        Dim Snode As Integer 'second node
    End Structure
    Public BranchEnd As Integer

    Public fDebug As Integer = FreeFile()
    Public oldPath As String


    '
    Function getPathId(ByVal n1 As Integer) As Integer
        For i As Integer = 1 To nPath
            'nl1 = Sheets("path").Cells(2, i + 1)
            Dim nl1 As Integer = Path(i).pNode(1).id
            If nl1 = n1 Then 'And nl2 = n2 Then
                getPathId = i 'Sheets("path").Cells(2, i)
                Exit Function
            End If
        Next
    End Function

    Private Function getLinkId(ByVal n1 As Integer, ByVal n2 As Integer) As Integer
        Dim CHK As String = HashLinkID(n1 & "-" & n2)
        If CHK <> "" Then
            Return CInt(CHK)
        Else
            Return CInt(-1)
        End If
    End Function
    '
    Function getNodeId(ByVal n1 As Integer) As Integer
        For i As Integer = 1 To nNode
            Dim nl1 As Integer = Node(i).id  'Sheets("Node").Cells(i + 2, 1)
            If nl1 = n1 Then
                getNodeId = i 'Node(i).id 'Sheets("node").Cells(i + 2, 1)
                Exit Function
            End If
        Next
    End Function

    Function getNodeCoord(ByVal n1 As Integer) As POINTAPI
        For i As Integer = 1 To nNode
            Dim nl1 As Integer = Node(i).id  'Sheets("Node").Cells(i + 2, 1)
            If nl1 = n1 Then
                getNodeCoord.X = Node(i).x
                getNodeCoord.Y = Node(i).y
                Exit Function
            End If
        Next
    End Function
    '
    Sub initial()
        If LoadNetwork = False Then
            LoadNetwork = True
        Else
            Exit Sub
        End If
        Dim Fx As Integer = FreeFile()
        'FileOpen(Fx, My.Application.Info.DirectoryPath & "\network_02.txt", OpenMode.Input)
        'FileOpen(Fx, "f:\SEATEC_Project\Web-GIS\2012-06-05\PWA.NET\bin\PWANET.txt", OpenMode.Input)
        FileOpen(Fx, "F:\SEATEC_Project\Web-GIS\2012-07-27 mwNetworkGeneration\example_data\PWANET2.txt", OpenMode.Input)
        Dim st As String
        Dim a As Object
        'Node number
        st = LineInput(Fx) : a = Split(st, vbTab)
        nNode = CInt(a(1))
        'Link number
        st = LineInput(Fx) : a = Split(st, vbTab)
        nLink = CInt(a(1))
        '
        ReDim Node(nNode)
        ReDim Link(nLink)

        'Form1.txtMeter.Text = nNode
        'Form1.txtPipe.Text = nLink
        ' get Node data
        '
        st = LineInput(Fx) 'Node title
        For i As Integer = 1 To nNode
            st = LineInput(Fx) : a = Split(st, vbTab)
            Node(i).id = CInt(a(0))
            Node(i).x = CDbl(a(1))
            Node(i).y = CDbl(a(2))
            Node(i).z = CDbl(a(3))
            Node(i).Q = CDbl(a(4))
            'Node(i).Type = CStr(a(5))
            'Node(i).Condition = CStr(a(6))
            'Node(i).used = False
        Next            '
        ' get Link data
        '
        st = LineInput(Fx) 'Link title
        For i As Integer = 1 To nLink
            st = LineInput(Fx) : a = Split(st, vbTab)
            Link(i).id = CInt(a(0))
            Link(i).NodeF = CInt(a(1))
            Link(i).NodeT = CInt(a(2))
            Link(i).L = ((Node(Link(i).NodeF).x - Node(Link(i).NodeT).x) ^ 2 + (Node(Link(i).NodeF).y - Node(Link(i).NodeT).y) ^ 2) ^ 0.5
            'Link(i).diameter = CDbl(a(4))
            'Link(i).Area = Math.PI / 4 * (Link(i).diameter / 100) ^ 2 'CDbl(a(5))
            'Link(i).c = 150 'CDbl(a(6))

        Next
        FileClose(Fx)
        '--------------------------------------------------------
        ' Create HashTable for Link ID and Node ID
        '
    End Sub

End Module