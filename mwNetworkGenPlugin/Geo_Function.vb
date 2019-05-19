Module Geo_Function

    Public Const PI = 3.14159265

    Structure POINTAPI
        Dim X As Double 'Long
        Dim Y As Double 'Long
        Dim Z As Double 'Long
    End Structure

    Structure LineType
        Dim Pt1 As POINTAPI
        Dim Pt2 As POINTAPI
        Dim Dx As Double
        Dim DY As Double
    End Structure

    Structure CircleType
        Dim pt As POINTAPI
        Dim R As Double
    End Structure

    Structure PathType
        Dim nPt As Integer
        Dim pt() As POINTAPI
    End Structure

    Public LL As LineType
    Public BcPath As PathType
    Public BcPathOffset As PathType

    '{ID:GEO-001}
    Function LineGIS(ByVal X1 As Double, ByVal Y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal Lwidth As Long, ByVal LColor As System.Drawing.Color) As Boolean
        Dim hDraw
        hDraw = g_MW.View.Draw.NewDrawing(1)
        g_MW.View.Draw.DrawLine(X1, Y1, x2, y2, Lwidth, LColor)
        LineGIS = True
    End Function

   Function LineGIS(ByVal L1 As LineType, ByVal Lwidth As Long, ByVal LColor As System.Drawing.Color) As Boolean
      Dim hDraw
      Dim x1 As Double = L1.Pt1.X
      Dim y1 As Double = L1.Pt1.Y
      Dim x2 As Double = L1.Pt2.X
      Dim y2 As Double = L1.Pt2.Y
      hDraw = g_MW.View.Draw.NewDrawing(1)
      g_MW.View.Draw.DrawLine(X1, Y1, x2, y2, Lwidth, LColor)
      LineGIS = True
   End Function

    '{ID:GEO-002}
    Function PointGIS(ByVal X1 As Double, ByVal Y1 As Double, ByVal Pwidth As Long, ByVal PColor As System.Drawing.Color) As Boolean
        Dim hDraw
        hDraw = g_MW.View.Draw.NewDrawing(1)
        g_MW.View.Draw.DrawCircle(X1, Y1, Pwidth, PColor, True)
        PointGIS = True
    End Function

    '{ID:GEO-003}
    Public Sub Centroid(ByVal BoundaryPath As PathType, ByRef CentroidPoint As POINTAPI)
        Dim pt As Integer
        Dim second_factor As Double
        Dim polygon_area As Double

        'Add the first point at the end of the array.
        ReDim Preserve BoundaryPath.pt(BoundaryPath.nPt + 1)
        BoundaryPath.pt(BoundaryPath.nPt + 1) = BoundaryPath.pt(1)
        Dim X, Y
        ' Find the centroid.
        X = 0
        Y = 0
        For pt = 1 To BoundaryPath.nPt - 1 ' m_NumPoints
            second_factor = _
                BoundaryPath.pt(pt).X * BoundaryPath.pt(pt + 1).Y - _
                BoundaryPath.pt(pt + 1).X * BoundaryPath.pt(pt).Y
            X = X + (BoundaryPath.pt(pt).X + BoundaryPath.pt(pt + 1).X) * second_factor
            Y = Y + (BoundaryPath.pt(pt).Y + BoundaryPath.pt(pt + 1).Y) * second_factor
        Next pt

        ' Divide by 6 times the polygon's area.
        polygon_area = PathArea(BoundaryPath)               '{GEO-004}
        X = X / 6 / polygon_area
        Y = Y / 6 / polygon_area

        ' If the values are negative, the polygon is
        ' oriented counterclockwise. Reverse the signs.
        If X < 0 Then
            X = -X
            Y = -Y
        End If
        CentroidPoint.X = X
        CentroidPoint.Y = Y
        '    frmMap.MAP.CurrentX = X
        '    frmMap.MAP.CurrentY = Y
        '    frmMap.MAP.Print "Centroid of Planted Area"
    End Sub

    '{ID:GEO-004}
    Function PathArea(ByVal mPath As PathType)
        Dim area
        Dim pt
        area = 0
        mPath.pt(mPath.nPt + 1) = mPath.pt(1)
        For pt = 1 To mPath.nPt
            area = area + ((mPath.pt(pt).X * mPath.pt(pt + 1).Y) - (mPath.pt(pt + 1).X * mPath.pt(pt).Y)) / 2
        Next pt
        If area < 0 Then area = Math.Abs(area)
        PathArea = area
    End Function

    '{ID:GEO-005}
    Function Length(ByVal mLine As LineType)
        Length = (diffX(mLine) ^ 2 + diffY(mLine) ^ 2) ^ 0.5 '{GEO-009} {GEO-010}
    End Function

    '{ID:GEO-006}
    Function PathLength(ByVal Pth As PathType)
        Dim LP As LineType
        Dim I
        PathLength = 0
        For I = 1 To Pth.nPt - 1
            LP.Pt1.X = Pth.pt(I).X
            LP.Pt1.Y = Pth.pt(I).Y
            LP.Pt2.X = Pth.pt(I + 1).X
            LP.Pt2.Y = Pth.pt(I + 1).Y
            PathLength = PathLength + Length(LP) '{GEO-005}
        Next
    End Function

    '{ID:GEO-007}
    Function PointInPath(ByVal Pth As PathType, ByVal Dist As Double, ByVal ID As Integer) As POINTAPI
        Dim LP As LineType
        Dim PL1, PL2
        If Dist = 0 Or Dist = PathLength(Pth) Then  '{GEO-006}
            PointInPath = Pth.pt(1)
            Exit Function
        End If
        If Dist > 0 And Dist < PathLength(Pth) Then  '{GEO-006}
            PL1 = 0
            PL2 = 0
            Dim I
            For I = 1 To Pth.nPt - 1
                LP.Pt1.X = Pth.pt(I).X
                LP.Pt1.Y = Pth.pt(I).Y
                LP.Pt2.X = Pth.pt(I + 1).X
                LP.Pt2.Y = Pth.pt(I + 1).Y
                PL2 = PL2 + Length(LP)                    '{GEO-005}
                If PL1 <= Dist And Dist < PL2 Then
                    Dim subDist
                    subDist = Dist - PL1
                    PointInPath = PointInLine(LP, subDist)  '{GEO-008}
                    ID = I
                    Exit Function
                End If
                PL1 = PL2
            Next
        Else
            'MsgBox "Function Error"
        End If
    End Function

    '{ID:GEO-008}
    Function PointInLine(ByVal L As LineType, ByVal subDist As Double) As POINTAPI
        PointInLine.X = subDist / Length(L) * (L.Pt2.X - L.Pt1.X) + L.Pt1.X '{GEO-005}
        PointInLine.Y = subDist / Length(L) * (L.Pt2.Y - L.Pt1.Y) + L.Pt1.Y '{GEO-005}
    End Function

    '{ID:GEO-009}
    Function diffX(ByVal mLine As LineType)
        diffX = mLine.Pt2.X - mLine.Pt1.X
    End Function

    '{ID:GEO-010}
    Function diffY(ByVal mLine As LineType)
        diffY = mLine.Pt2.Y - mLine.Pt1.Y
    End Function

    '{ID:GEO-011}
    Function AzmAngle(ByVal mLine As LineType)
        If diffX(mLine) = 0 Or diffY(mLine) = 0 Then      '{GEO-009} {GEO-010}
            'Vertical Case
            If diffX(mLine) = 0 Then                        '{GEO-009}
                If mLine.Pt1.Y > mLine.Pt2.Y Then AzmAngle = 0
                If mLine.Pt1.Y < mLine.Pt2.Y Then AzmAngle = 180
            End If
            'Horizontal Case
            If diffY(mLine) = 0 Then                        '{GEO-010}
                If mLine.Pt1.X > mLine.Pt2.X Then AzmAngle = 270
                If mLine.Pt1.X < mLine.Pt2.X Then AzmAngle = 90
            End If
        Else
            'Angle Computation
            Dim angR As Double
            angR = Math.Atan(diffY(mLine) / diffX(mLine))         '{GEO-010}   {GEO-009}
            If mLine.Pt1.Y < mLine.Pt2.Y Then
                If mLine.Pt1.X < mLine.Pt2.X Then
                    'Quadrant 1
                    AzmAngle = 90 - angR * 180 / PI
                Else
                    'Qurdrant 2
                    AzmAngle = 270 - angR * 180 / PI
                End If
            Else
                If mLine.Pt1.X < mLine.Pt2.X Then
                    'Qurdrant 3
                    AzmAngle = 90 - angR * 180 / PI
                Else
                    'Qurdrant 4
                    AzmAngle = 270 - angR * 180 / PI
                End If
            End If
        End If
    End Function

    '{ID:GEO-013}
    Function PolygonArea(ByVal nPoints As Integer, ByVal R_Point() As POINTAPI) As Double
        Dim area
        Dim pt
        area = 0
        R_Point(nPoints + 1) = R_Point(1)
        For pt = 1 To nPoints
            area = area + ((R_Point(pt).X * R_Point(pt + 1).Y) - (R_Point(pt + 1).X * R_Point(pt).Y)) / 2
        Next pt
        If area < 0 Then area = Math.Abs(area)
        'PolygonArea = area
        Return area
    End Function

    '{ID:GEO-014}
    Function MidLine(ByVal mLine As LineType) As POINTAPI
        MidLine.X = (mLine.Pt2.X + mLine.Pt1.X) / 2
        MidLine.Y = (mLine.Pt2.Y + mLine.Pt1.Y) / 2
        MidLine.Z = (mLine.Pt2.Z + mLine.Pt1.Z) / 2
        Return MidLine
    End Function

    '{ID:GEO-015}
    Public Function Line_Line(ByVal a As LineType, ByVal b As LineType, ByVal desx As Double, ByVal desy As Double) As Boolean
        On Error GoTo err1
        Dim A1 As Double
        Dim A2 As Double
        Dim b1 As Double
        Dim b2 As Double
        Dim C1 As Double
        Dim C2 As Double
        Dim F As Boolean

        a.Dx = a.Pt2.X - a.Pt1.X
        a.DY = a.Pt2.Y - a.Pt1.Y
        b.Dx = b.Pt2.X - b.Pt1.X
        b.DY = b.Pt2.Y - b.Pt1.Y

        A1 = a.Dx        'a.xt
        A2 = a.DY        'a.yt
        b1 = b.Dx * -1   '-b.xt
        b2 = b.DY * -1   '-b.yt
        C1 = a.Pt1.X - b.Pt1.X
        C2 = a.Pt1.Y - b.Pt1.Y
        Dim T, S As Double
        T = 0
        S = 0
        If b1 = 0 Then
            If A1 <> 0 Then T = -C1 / A1
        Else
            T = ((b2 * C1) / b1 - C2) / (A2 - (b2 * A1) / b1)
        End If
        desx = a.Pt1.X + a.Dx * T
        desy = a.Pt1.Y + a.DY * T
        If Point_Line(desx, desy, a) And Point_Line(desx, desy, b) Then F = True '{GEO-016}

        Line_Line = F
        Return Line_Line

        Exit Function
err1:
        Line_Line = False
    End Function

    '{ID:GEO-016}
    Public Function Point_Line(ByVal X As Double, ByVal Y As Double, ByVal P As LineType) As Boolean
        Dim t1 As Double
        Dim t2 As Double
        Dim op As Boolean
        Dim T As Double
        P.Dx = P.Pt2.X - P.Pt1.X
        P.DY = P.Pt2.Y - P.Pt1.Y
        If P.Dx = 0 Then
            T = (Y - P.Pt1.Y) / P.DY
            If T <= 1 And T >= 0 And X = P.Pt1.X Then op = True
        End If
        If P.DY = 0 Then
            T = (X - P.Pt1.X) / P.Dx
            If T <= 1 And T >= 0 And Y = P.Pt1.Y Then op = True
        End If
        If P.Dx <> 0 And P.DY <> 0 Then
            t1 = (X - P.Pt1.X) / P.Dx
            t2 = (Y - P.Pt1.Y) / P.DY
            If Math.Abs(t1 - t2) <= 0.5 And t1 <= 1 And t1 >= 0 Then op = True
        End If
        Point_Line = op
    End Function

    '{ID:GEO-017}

    Public Function PointInPolygon(ByVal Poly() As POINTAPI, ByVal Xray As Double, ByVal YofRay As Double) As Boolean
        Dim X As Long
        Dim PolyCount As Long
        Dim NumSidesCrossed As Long
        Dim LenOfSide As Double
        Dim CrossPt As POINTAPI
        CrossPt.X = Xray
        PolyCount = 1 + UBound(Poly) - LBound(Poly)
        For X = LBound(Poly) To UBound(Poly)
            If Poly(X).X > Xray Xor Poly((X + 1) Mod PolyCount).X > Xray Then
                CrossPt.Y = Y_at_X_Ray(Xray, Poly(X), Poly((X + 1) Mod PolyCount))
                If CrossPt.Y > YofRay Then
                    LenOfSide = PtDist(Poly(X), Poly((X + 1) Mod PolyCount))
                    If LenOfSide > PtDist(Poly(X), CrossPt) And _
                      LenOfSide > PtDist(Poly((X + 1) Mod PolyCount), CrossPt) Then
                        NumSidesCrossed = NumSidesCrossed + 1
                    End If
                End If
            End If
        Next
        If NumSidesCrossed Mod 2 Then PointInPolygon = True Else PointInPolygon = False
    End Function

    Private Function Y_at_X_Ray(ByVal Xray As Double, _
      ByVal p1 As POINTAPI, _
      ByVal p2 As POINTAPI) As Double
        Dim m As Double
        Dim b As Double
        m = (p2.Y - p1.Y) / (p2.X - p1.X)
        b = (p1.Y * p2.X - p1.X * p2.Y) / (p2.X - p1.X)
        Y_at_X_Ray = m * Xray + b
    End Function

    Private Function PtDist(ByVal p1 As POINTAPI, ByVal p2 As POINTAPI) As Double
        PtDist = Math.Sqrt((p2.Y - p1.Y) * (p2.Y - p1.Y) + _
        (p2.X - p1.X) * (p2.X - p1.X))
    End Function

    '{ID:GEO-018}
    Function Intersect(ByVal p1 As POINTAPI, ByVal p2 As POINTAPI, ByVal p3 As POINTAPI, ByVal p4 As POINTAPI) As Boolean
        Intersect = (((CCW(p1, p2, p3) * CCW(p1, p2, p4)) <= 0) And ((CCW(p3, p4, p1) * CCW(p3, p4, p2) <= 0)))  '{GEO-019}
    End Function

    '{ID:GEO-019}
    Function CCW(ByVal p0 As POINTAPI, ByVal p1 As POINTAPI, ByVal p2 As POINTAPI) As Long

        Dim dx1 As Long, dx2 As Long
        Dim dy1 As Long, dy2 As Long
        dx1 = p1.X - p0.X
        dx2 = p2.X - p0.X

        dy1 = p1.Y - p0.Y
        dy2 = p2.Y - p0.Y

        If (dx1 * dy2 > dy1 * dx2) Then
            CCW = 1
        Else
            CCW = -1
        End If
    End Function

    '{ID:GEO-020}
    Function ParallelLine(ByVal L1 As LineType, ByVal L2 As LineType) As Boolean
        ParallelLine = False
        Dim Ang1, Ang2
        Ang1 = AzmAngle(L1)                                  '{GEO-011}
        Ang2 = AzmAngle(L2)                                  '{GEO-011}
        If Math.Abs(Ang1 - Ang2) <= 1 Then ParallelLine = True
        If Math.Abs(Math.Abs(Ang1 - Ang2) - 180) <= 1 Then ParallelLine = True
        'frmEvalutionTable.grdGA4Mainline.TextMatrix(2, 10) = Format(Ang1, "0.00")
        'frmEvalutionTable.grdGA4Mainline.TextMatrix(3, 10) = Format(Ang2, "0.00")
    End Function

    '{ID:GEO-021}
    Function PerpendicularLine(ByVal L1 As LineType, ByVal L2 As LineType) As Boolean
        PerpendicularLine = False
        Dim Ang1, Ang2
        Ang1 = AzmAngle(L1)                                  '{GEO-011}
        Ang2 = AzmAngle(L2)                                  '{GEO-011}
        If Math.Abs(Math.Abs(Ang1 - Ang2) - 90) <= 1 Then PerpendicularLine = True
        If Math.Abs(Math.Abs(Ang1 - Ang2) - 270) <= 1 Then PerpendicularLine = True
    End Function

    '{ID:GEO-022}
    Function Offset(ByVal Dx As Double, ByVal p1 As POINTAPI, ByVal p2 As POINTAPI) As POINTAPI
        Dim C1 As CircleType
        Dim p3 As POINTAPI
        Dim LOffset As LineType
        Dim R As Double
        Dim Chk As Integer
        LOffset.Pt1 = p1
        LOffset.Pt2 = p2
        R = Length(LOffset)                                  '{GEO-005}
        C1.pt = p1
        C1.R = R + Dx
        LOffset.Pt1 = p1
        LOffset.Pt2 = DLengthPosition(p1, p2)                '{GEO-023}
        Chk = Line_Circle(LOffset, C1, Offset.X, Offset.Y, p3.X, p3.Y)  '{GEO-024}
        If Chk <= 0 Then
            Offset = p2
        End If
    End Function

    '{ID:GEO-023}
    Function DLengthPosition(ByVal p1 As POINTAPI, ByVal p2 As POINTAPI) As POINTAPI
        DLengthPosition.X = p1.X + (p2.X - p1.X) * 2
        DLengthPosition.Y = p1.Y + (p2.Y - p1.Y) * 2
    End Function

    '{ID:GEO-024}
    Public Function Line_Circle(ByVal P As LineType, ByVal k As CircleType, ByRef desx1 As Double, ByRef desy1 As Double, ByRef desx2 As Double, ByRef desy2 As Double) As Integer
        Dim a As Double
        Dim X1 As Double
        Dim x2 As Double
        Dim Y1 As Double
        Dim y2 As Double
        Dim point1_exist As Boolean
        Dim point2_exist As Boolean
        Dim points As Integer ' points of intersection
        Dim intersection As Boolean
        Dim Dx, DY, DR, d
        X1 = P.Pt1.X - k.pt.X
        x2 = P.Pt2.X - k.pt.X          'p.Pt2.x=(p.xn + p.xt)
        Y1 = P.Pt1.Y - k.pt.Y
        y2 = P.Pt2.Y - k.pt.Y          'p.Pt2.y=(p.yn + p.yt)
        Dx = x2 - X1
        DY = y2 - Y1
        DR = Math.Sqrt(Dx ^ 2 + DY ^ 2)
        d = X1 * y2 - x2 * Y1

        If (k.R ^ 2 * DR ^ 2 - d ^ 2) >= 0 Then intersection = True

        If intersection = True Then
            If (k.R ^ 2 * DR ^ 2 - d ^ 2) < 0 Then
                a = 0
            Else
                a = Math.Sqrt(k.R ^ 2 * DR ^ 2 - d ^ 2)
            End If
            desx1 = (d * DY + My_Sgn(DY) * Dx * a) / DR ^ 2 + k.pt.X     '{GEO-025}
            desy1 = (-d * Dx + Math.Abs(DY) * a) / DR ^ 2 + k.pt.Y
            desx2 = (d * DY - My_Sgn(DY) * Dx * a) / DR ^ 2 + k.pt.X     '{GEO-025}
            desy2 = (-d * Dx - Math.Abs(DY) * a) / DR ^ 2 + k.pt.Y
        End If
        point1_exist = Point_Line(desx1, desy1, P)                       '{GEO-016}
        point2_exist = Point_Line(desx2, desy2, P)                       '{GEO-016}
        If point1_exist And point2_exist Then
            points = 2
        Else
            points = 0
            If point1_exist Then points = 1
            If point2_exist Then points = 1 : desx1 = desx2 : desy1 = desy2
        End If

        Line_Circle = points

    End Function

    '{ID:GEO-025}
    Public Function My_Sgn(ByVal X) As Integer
        If X < 0 Then
            My_Sgn = -1
        Else
            My_Sgn = 1
        End If
    End Function

    '{ID:GEO-026}
    Function PerpendicularPoint(ByVal L1 As LineType, ByVal L2 As Double) As POINTAPI
        Dim Ang1, Ang2
        Dim mp As POINTAPI
        Dim p1 As POINTAPI, p2 As POINTAPI, LL1 As LineType, LL2 As LineType

        Ang1 = AzmAngle(L1)                                 '{GEO-011}
        mp = MidLine(L1)                                    '{GEO-014}
        p1.X = Math.Sin((Ang1 - 90) * PI / 180) * L2 + mp.X
        p1.Y = Math.Cos((Ang1 - 90) * PI / 180) * L2 + mp.Y

        p2.X = Math.Sin((Ang1 + 90) * PI / 180) * L2 + mp.X
        p2.Y = Math.Cos((Ang1 + 90) * PI / 180) * L2 + mp.Y

        LL1.Pt1 = p1 : LL1.Pt2 = CenPTA
        LL2.Pt1 = p2 : LL2.Pt2 = CenPTA

        If Length(LL1) > Length(LL2) Then                   '{GEO-005}
            PerpendicularPoint = p2
            'LineGIS p1.X, p1.Y, mp.X, mp.Y, 1, vbRed
        Else
            PerpendicularPoint = p1
            'LineGIS p2.X, p2.Y, mp.X, mp.Y, 1, vbRed
        End If

    End Function

    '{ID:GEO-027}
    Sub Perpend2Point(ByVal L1 As LineType, ByVal mp As POINTAPI, ByVal L2 As Double, ByRef LP As LineType)
        Dim Ang1
        Ang1 = AzmAngle(L1) + 90                             '{GEO-011}
        LP.Pt1.X = Math.Sin((Ang1) * PI / 180) * L2 + mp.X
        LP.Pt1.Y = Math.Cos((Ang1) * PI / 180) * L2 + mp.Y
        LP.Pt2.X = Math.Sin((Ang1 + 180) * PI / 180) * L2 + mp.X
        LP.Pt2.Y = Math.Cos((Ang1 + 180) * PI / 180) * L2 + mp.Y
    End Sub


    '{ID:GEO-028}
    Function PointTodist(ByVal Pth As PathType, ByVal pp1 As POINTAPI) As Double
        Dim LP As LineType
        Dim PLength As Double
        Dim Chk As Boolean
        Dim Dist As Double
        Dim I
        Dist = 0
        PLength = PathLength(Pth)                        '{GEO-006}
        PointTodist = -1
        For I = 1 To Pth.nPt - 1
            LP.Pt1.X = Pth.pt(I).X
            LP.Pt1.Y = Pth.pt(I).Y
            LP.Pt2.X = Pth.pt(I + 1).X
            LP.Pt2.Y = Pth.pt(I + 1).Y
            If Point_Line(pp1.X, pp1.Y, LP) Then             '{GEO-016}
                LP.Pt2 = pp1
                PointTodist = Dist + Length(LP)              '{GEO-005}
                Exit Function
            Else
                Dist = Dist + Length(LP)                     '{GEO-005}
            End If
        Next
    End Function

    '{ID:GEO-029}
    Function pointOnBC(ByVal LL As LineType, ByVal bc As PathType, ByVal pp1 As POINTAPI, ByVal pp2 As POINTAPI) As Boolean
        Dim p1 As POINTAPI
        Dim p2 As POINTAPI, X, Y
        Dim LineEadge As LineType
        Dim I, j, k
        Dim check As Boolean
        check = True
        For I = 1 To bc.nPt - 1
            p1 = BcPath.pt(I)
            p2 = BcPath.pt(I + 1)
            LineEadge.Pt1 = p1
            LineEadge.Pt2 = p2
            If check = True Then
                If Line_Line(LL, LineEadge, X, Y) = True Then                       '{GEO-015}
                    pp1.X = X
                    pp1.Y = Y
                    check = False
                End If
            Else
                If Line_Line(LL, LineEadge, X, Y) = True Then                       '{GEO-015}
                    pp2.X = X
                    pp2.Y = Y
                    pointOnBC = True
                End If
            End If
        Next
    End Function
    '
    '  Find point on planted edge
    '  2008-11-18
    '{ID:GEO-030}
    Function DistToPoint(ByVal Pth As PathType, ByVal Dist As Double) As POINTAPI
        Dim LP As LineType
        Dim PLength As Double
        Dim PL1, PL2
        PLength = PathLength(Pth)         '{GEO-006}
        If Dist = 0 Or Dist = PLength Then
            DistToPoint = Pth.pt(1)
            Exit Function
        End If
        If Dist > 0 And Dist < PLength Then
            PL1 = 0
            PL2 = 0
            Dim I
            For I = 1 To Pth.nPt - 1
                LP.Pt1.X = Pth.pt(I).X
                LP.Pt1.Y = Pth.pt(I).Y
                LP.Pt2.X = Pth.pt(I + 1).X
                LP.Pt2.Y = Pth.pt(I + 1).Y
                PL2 = PL2 + Length(LP)                     '{GEO-005}
                If PL1 <= Dist And Dist < PL2 Then
                    Dim subDist
                    subDist = Dist - PL1
                    DistToPoint = PointInLine(LP, subDist)   '{GEO-008}
                    Exit Function
                End If
                PL1 = PL2
            Next
        Else
            'MsgBox "Function Error"
        End If

    End Function


    '********************************************************************************
    '********************************************************************************
    '
    ' UNDEFINE ID
    '
    '********************************************************************************
    '********************************************************************************



    Function diffZ(ByVal mLine As LineType)
        diffZ = mLine.Pt2.Z - mLine.Pt1.Z
    End Function


    Function destinationPiont(ByVal ang As Double, ByVal OriginalPoint As POINTAPI, ByVal L As Double) As POINTAPI
        Dim angR
        angR = ang * PI / 180
        destinationPiont.Y = Math.Cos(angR) * L + OriginalPoint.Y
        destinationPiont.X = Math.Sin(angR) * L + OriginalPoint.X
        'destinationPiont.X = Math.sin(angR) * L + OriginalPoint.Y
    End Function
    '
    'Function for convert the angle from Azimut to 4 qurdant angle.
    '
    Function Azm2Qurdrant(ByVal Azm) As Double
        If Azm >= 0 And Azm < 90 Then Azm2Qurdrant = 90 - Azm
        If Azm >= 90 And Azm < 180 Then Azm2Qurdrant = 360 - (Azm - 90)
        If Azm >= 180 And Azm < 270 Then Azm2Qurdrant = 270 - (Azm - 180)
        If Azm >= 270 And Azm < 360 Then Azm2Qurdrant = 90 + (360 - Azm)
    End Function
    '
    'This function is for find the intersection point of 2 circles.
    '
    Public Function Circle_Circle(ByVal k As CircleType, ByVal L As CircleType, ByVal desx1 As Double, ByVal desy1 As Double, ByVal desx2 As Double, ByVal desy2 As Double) As Integer
        Dim d As Double
        Dim X As Double
        Dim Y As Double
        Dim points As Integer
        Dim Dx, DY
        Dx = k.pt.X - L.pt.X
        DY = k.pt.Y - L.pt.Y
        d = Math.Sqrt(Dx ^ 2 + DY ^ 2)
        If d <> 0 Then
            X = (d ^ 2 - L.R ^ 2 + k.R ^ 2) / (2 * d)
            If (k.R ^ 2 - X ^ 2) > 0 Then Y = Math.Sqrt(k.R ^ 2 - X ^ 2)
            desx1 = k.pt.X - (X * (Dx / d) + Y * (DY / d))
            desy1 = k.pt.Y - (-Y * (Dx / d) + X * (DY / d))
            desx2 = k.pt.X - (X * (Dx / d) - Y * (DY / d))
            desy2 = k.pt.Y - (Y * (Dx / d) + X * (DY / d))
        End If
        If d < (k.R + L.R) And d > Math.Abs(k.R - L.R) Then points = 2
        If d = (k.R + L.R) Then points = 1

        Circle_Circle = points

    End Function


#Region "Minimum Distance between a Point and a Line"

    Public Function lineMagnitude(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
        lineMagnitude = ((x2 - x1) ^ 2 + (y2 - y1) ^ 2) ^ 0.5
    End Function

    Public Function DistancePointLine(ByVal px As Double, ByVal py As Double, ByVal x1 As Double, ByVal y1 As Double, _
                                      ByVal x2 As Double, ByVal y2 As Double, ByRef xx As Double, ByRef yy As Double) As Double
        ' Returns 9999 on 0 denominator conditions.
        Dim LineMag As Double, u As Double
        Dim ix As Double, iy As Double ' intersecting point

        LineMag = lineMagnitude(x1, y1, x2, y2)
        If LineMag < 0.00000001 Then
            xx = px
            yy = py
            DistancePointLine = 99999999 : Exit Function
        End If

        xx = px
        yy = py
        u = (((px - x1) * (x2 - x1)) + ((py - y1) * (y2 - y1)))
        u = u / (LineMag * LineMag)
        If u < 0.00001 Or u > 1 Then
            '// closest point does not fall within the line segment, take the shorter distance
            '// to an endpoint
            ix = lineMagnitude(px, py, x1, y1)
            iy = lineMagnitude(px, py, x2, y2)
            If ix > iy Then DistancePointLine = iy Else DistancePointLine = ix
        Else
            ' Intersecting point is on the line, use the formula
            ix = x1 + u * (x2 - x1)
            iy = y1 + u * (y2 - y1)
            DistancePointLine = lineMagnitude(px, py, ix, iy)
        End If
        xx = ix
        yy = iy
    End Function

#End Region



    Function lineIntersection(ByVal L1 As LineType, ByVal L2 As LineType, ByRef pt As POINTAPI) As Boolean
        'Function lineIntersection(ByVal Ax As Double, ByVal Ay As Double, ByVal Bx As Double, ByVal By As Double, ByVal Cx As Double, ByVal Cy As Double, _
        ' ByVal Dx As Double, ByVal Dy As Double, ByRef X As Double, ByRef Y As Double) As Boolean

        Dim distAB As Double, theCos As Double, theSin As Double, newX As Double, ABpos As Double
        Dim Ax As Double = L1.Pt1.X
        Dim Ay As Double = L1.Pt1.Y
        Dim Bx As Double = L1.Pt2.X
        Dim By As Double = L1.Pt2.Y

        Dim Cx As Double = L2.Pt1.X
        Dim Cy As Double = L2.Pt1.Y
        Dim Dx As Double = L2.Pt2.X
        Dim Dy As Double = L2.Pt2.Y

        '  Fail if either line is undefined.
        If Ax = Bx AndAlso Ay = By OrElse Cx = Dx AndAlso Cy = Dy Then
            Return False
        End If

        '  (1) Translate the system so that point A is on the origin.
        Bx -= Ax
        By -= Ay
        Cx -= Ax
        Cy -= Ay
        Dx -= Ax
        Dy -= Ay

        '  Discover the length of segment A-B.
        distAB = (Bx * Bx + By * By) ^ 0.5

        '  (2) Rotate the system so that point B is on the positive X axis.
        theCos = Bx / distAB
        theSin = By / distAB
        newX = Cx * theCos + Cy * theSin
        Cy = Cy * theCos - Cx * theSin
        Cx = newX
        newX = Dx * theCos + Dy * theSin
        Dy = Dy * theCos - Dx * theSin
        Dx = newX

        '  Fail if the lines are parallel.
        If Cy = Dy Then
            Return False
        End If

        '  (3) Discover the position of the intersection point along line A-B.
        ABpos = Dx + (Cx - Dx) * Dy / (Dy - Cy)

        '  (4) Apply the discovered position to line A-B in the original coordinate system.
        pt.X = Ax + ABpos * theCos
        pt.Y = Ay + ABpos * theSin

        '  Success.
        Return True
    End Function
    Public Sub FindLineIntersection( _
ByVal x11 As Double, ByVal y11 As Double, _
ByVal x12 As Double, ByVal y12 As Double, _
ByVal x21 As Double, ByVal y21 As Double, _
ByVal x22 As Double, ByVal y22 As Double, _
ByRef inter_x As Double, ByRef inter_y As Double, _
ByRef inter_x1 As Double, ByRef inter_y1 As Double, _
ByRef inter_x2 As Double, ByRef inter_y2 As Double)
        Dim dx1 As Double
        Dim dy1 As Double
        Dim dx2 As Double
        Dim dy2 As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim denominator As Double

        ' Get the segments' parameters.
        dx1 = x12 - x11
        dy1 = y12 - y11
        dx2 = x22 - x21
        dy2 = y22 - y21

        ' Solve for t1 and t2.
        On Error Resume Next
        denominator = (dy1 * dx2 - dx1 * dy2)
        t1 = ((x11 - x21) * dy2 + (y21 - y11) * dx2) / _
            denominator
        If Err.Number <> 0 Then
            ' The lines are parallel.
            inter_x = 1.0E+38 : inter_y = 1.0E+38
            inter_x1 = 1.0E+38 : inter_y1 = 1.0E+38
            inter_x2 = 1.0E+38 : inter_y2 = 1.0E+38
            Exit Sub
        End If
        On Error GoTo 0
        t2 = ((x21 - x11) * dy1 + (y11 - y21) * dx1) / _
            -denominator

        ' Find the point of intersection.
        inter_x = x11 + dx1 * t1
        inter_y = y11 + dy1 * t1

        ' Find the closest points on the segments.
        If t1 < 0 Then
            t1 = 0
        ElseIf t1 > 1 Then
            t1 = 1
        End If
        If t2 < 0 Then
            t2 = 0
        ElseIf t2 > 1 Then
            t2 = 1
        End If
        inter_x1 = x11 + dx1 * t1
        inter_y1 = y11 + dy1 * t1
        inter_x2 = x21 + dx2 * t2
        inter_y2 = y21 + dy2 * t2
    End Sub

    Public Function SegmentsIntersect(ByVal L1 As LineType, ByVal L2 As LineType, ByRef pt As POINTAPI) As Boolean
        Dim dx As Double
        Dim dy As Double
        Dim da As Double
        Dim db As Double
        Dim t As Double

        Dim x1 As Double = L1.Pt1.X
        Dim y1 As Double = L1.Pt1.Y
        Dim x2 As Double = L1.Pt2.X
        Dim y2 As Double = L1.Pt2.Y

        Dim a1 As Double = L2.Pt1.X
        Dim b1 As Double = L2.Pt1.Y
        Dim a2 As Double = L2.Pt2.X
        Dim b2 As Double = L2.Pt2.Y

        Dim s As Double

        dx = x2 - x1
        dy = y2 - y1
        da = a2 - a1
        db = b2 - b1
        If (da * dy - db * dx) = 0 Then
            ' The segments are parallel.
            SegmentsIntersect = False
            Exit Function
        End If

        s = (dx * (b1 - y1) + dy * (x1 - a1)) / (da * dy - db * _
            dx)
        t = (da * (y1 - b1) + db * (a1 - x1)) / (db * dx - da * _
            dy)
        SegmentsIntersect = (s >= 0.0# And s <= 1.0# And _
                             t >= 0.0# And t <= 1.0#)

        ' If it exists, the point of intersection is:
        If SegmentsIntersect = True Then
            'Dim xx1, yy1, xx2, yy2, xx3, yy3 As Double
            'FindLineIntersection(x1, y1, x2, y2, a1, b1, a2, b2, xx1, yy1, xx2, yy2, xx3, yy3)
            pt.X = (x1 + t * dx)
            pt.Y = (y1 + t * dy)
        End If
    End Function



    Public Function DistToSegment(ByVal px As Double, ByVal py _
        As Double, ByVal X1 As Double, ByVal Y1 As Double, _
        ByVal X2 As Double, ByVal Y2 As Double, ByRef near_x As _
        Double, ByRef near_y As Double) As Double
        Dim dx As Double
        Dim dy As Double
        Dim t As Double

        dx = X2 - X1
        dy = Y2 - Y1
        If dx = 0 And dy = 0 Then
            ' It's a point not a line segment.
            dx = px - X1
            dy = py - Y1
            near_x = X1
            near_y = Y1
            DistToSegment = Math.Sqrt(dx * dx + dy * dy)
            Exit Function
        End If

        ' Calculate the t that minimizes the distance.
        t = ((px - X1) * dx + (py - Y1) * dy) / (dx * dx + dy * _
            dy)

        ' See if this represents one of the segment's
        ' end points or a point in the middle.
        If t < 0 Then
            dx = px - X1
            dy = py - Y1
            near_x = X1
            near_y = Y1
        ElseIf t > 1 Then
            dx = px - X2
            dy = py - Y2
            near_x = X2
            near_y = Y2
        Else
            near_x = X1 + t * dx
            near_y = Y1 + t * dy
            dx = px - near_x
            dy = py - near_y
        End If

        DistToSegment = Math.Sqrt(dx * dx + dy * dy)
    End Function

    Public Function DistToSegment(ByVal px As Double, ByVal py _
         As Double, ByVal LSegment As LineType, ByRef near_x As _
         Double, ByRef near_y As Double) As Double
        Dim X1, Y1, X2, Y2 As Double
        Dim dx As Double
        Dim dy As Double
        Dim t As Double
        X1 = LSegment.Pt1.X
        Y1 = LSegment.Pt1.Y
        X2 = LSegment.Pt2.X
        Y2 = LSegment.Pt2.Y
        dx = X2 - X1
        dy = Y2 - Y1
        If dx = 0 And dy = 0 Then
            ' It's a point not a line segment.
            dx = px - X1
            dy = py - Y1
            near_x = X1
            near_y = Y1
            DistToSegment = Math.Sqrt(dx * dx + dy * dy)
            Exit Function
        End If

        ' Calculate the t that minimizes the distance.
        t = ((px - X1) * dx + (py - Y1) * dy) / (dx * dx + dy * _
            dy)

        ' See if this represents one of the segment's
        ' end points or a point in the middle.
        If t < 0 Then
            dx = px - X1
            dy = py - Y1
            near_x = X1
            near_y = Y1
        ElseIf t > 1 Then
            dx = px - X2
            dy = py - Y2
            near_x = X2
            near_y = Y2
        Else
            near_x = X1 + t * dx
            near_y = Y1 + t * dy
            dx = px - near_x
            dy = py - near_y
        End If

        DistToSegment = Math.Sqrt(dx * dx + dy * dy)
    End Function
   '
   'Distance
   Function Distance(ByVal p1 As MapWinGIS.Point, ByVal p2 As MapWinGIS.Point) As Double
      Distance = ((p1.x - p2.x) ^ 2 + (p1.y - p2.y) ^ 2) ^ 0.5
   End Function
   Function Distance(ByVal p1 As POINTAPI, ByVal p2 As POINTAPI) As Double
      Distance = ((p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2) ^ 0.5
   End Function
   Public Function Distance(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
      Distance = ((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2) ^ 0.5
   End Function
   Public Function Distance(ByVal L As LineType) As Double
      Distance = ((L.Pt1.X - L.Pt2.X) ^ 2 + (L.Pt1.Y - L.Pt2.Y) ^ 2) ^ 0.5
   End Function
End Module
