Option Strict On
Option Infer On

Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface

' ==============================================================================
' SectionsModule 
' ==============================================================================
Module SectionsModule

    ' ---- Layer names (match your project) ----
    Private Const LYR_REBAR1 As String = "kr_donati1"
    Private Const LYR_REBAR2 As String = "kr_donati2"
    Private Const LYR_PHI As String = "kr_dnyazi"
    Private Const LYR_DIM As String = "kr_olcu"

    ' ---------------------- tiny drawing helpers ----------------------
    Private Sub AppendOnLayer(btr As BlockTableRecord, tr As Transaction, ent As Entity, layerName As String, Optional forceByLayer As Boolean = False)
        ent.Layer = layerName
        If forceByLayer Then ent.ColorIndex = 256
        btr.AppendEntity(ent)
        tr.AddNewlyCreatedDBObject(ent, True)
    End Sub

    Private Sub DrawRectOnLayer(btr As BlockTableRecord, tr As Transaction, x As Double, y As Double, w As Double, h As Double, layerName As String)
        Dim pl As New Autodesk.AutoCAD.DatabaseServices.Polyline()
        pl.AddVertexAt(0, New Point2d(x, y), 0, 0, 0)
        pl.AddVertexAt(1, New Point2d(x + w, y), 0, 0, 0)
        pl.AddVertexAt(2, New Point2d(x + w, y + h), 0, 0, 0)
        pl.AddVertexAt(3, New Point2d(x, y + h), 0, 0, 0)
        pl.Closed = True
        pl.ColorIndex = 256
        AppendOnLayer(btr, tr, pl, layerName, True)
    End Sub

    Private Sub DrawCircleOnLayer(btr As BlockTableRecord, tr As Transaction, cx As Double, cy As Double, r As Double, layerName As String)
        Dim c As New Circle(New Point3d(cx, cy, 0), Vector3d.ZAxis, Math.Max(0.1, r))
        c.ColorIndex = 256
        AppendOnLayer(btr, tr, c, layerName, True)
    End Sub

    Private Sub DrawSectionOutline(btr As BlockTableRecord, tr As Transaction, baseX As Double, baseY As Double, secW As Double, secH As Double)
        DrawRectOnLayer(btr, tr, baseX, baseY, secW, secH, LYR_DIM)
    End Sub

    Private Sub DrawPhiNote(btr As BlockTableRecord, tr As Transaction, x As Double, y As Double, txt As String)
        Dim mt As New MText() With {
            .Location = New Point3d(x, y, 0),
            .Attachment = AttachmentPoint.MiddleLeft,
            .TextHeight = 3.333,
            .Contents = txt
        }
        AppendOnLayer(btr, tr, mt, LYR_PHI, True)
    End Sub

    ' ---------------------- text helpers (parity) ----------------------
    Private Function QtyPhiText(qty As Integer, phiCm As Double) As String
        Return qty.ToString(Globalization.CultureInfo.InvariantCulture) &
               "Ø" & CInt(Math.Round(phiCm * 10.0R, MidpointRounding.AwayFromZero)).ToString()
    End Function

    ' If your project already has QtyPhiTextAdditional, you can delete this local copy.
    Private Function QtyPhiTextAdditional(qty As Integer, phiCm As Double) As String
        Return QtyPhiText(qty, phiCm) & " ila."
    End Function

    ' Row X’s including corners; center if n=1. Matches your usage.
    Private Function RowXsTouchingCorners(n As Integer, baseX As Double, secW As Double, cover As Double, r As Double) As List(Of Double)
        Dim xs As New List(Of Double)()
        Dim leftX = baseX + cover + r
        Dim rightX = baseX + secW - cover - r
        If n <= 0 Then Return xs
        If n = 1 Then
            xs.Add((leftX + rightX) / 2.0)
            Return xs
        End If
        If n = 2 Then
            xs.Add(leftX) : xs.Add(rightX)
            Return xs
        End If
        Dim stepX = (rightX - leftX) / (n - 1)
        For i = 0 To n - 1
            xs.Add(leftX + i * stepX)
        Next
        Return xs
    End Function

    ' Adjust (leg1 + leg2 + stirrupH) to be divisible by 5, alternating +1 starting from leg1; min 12.
    Private Sub AdjustLegPairToMultipleOf5(stirrupH As Double, ByRef leg1 As Double, ByRef leg2 As Double)
        Dim minLeg As Double = 12.0R
        If leg1 < minLeg Then leg1 = minLeg
        If leg2 < minLeg Then leg2 = minLeg
        Dim total As Double = stirrupH + leg1 + leg2
        Dim mod5 As Double = total Mod 5.0R
        If Math.Abs(mod5) < 0.0000001 Then Return
        Dim addNeeded As Integer = CInt(Math.Ceiling(5.0R - mod5))
        Dim flip As Boolean = False
        For i = 1 To addNeeded
            If Not flip Then leg1 += 1.0R Else leg2 += 1.0R
            flip = Not flip
        Next
    End Sub

    ' ================================================================
    ' PUBLIC: exact-behavior section renderer
    ' ================================================================
    Public Sub DrawRebarSectionForGroup(
        btr As BlockTableRecord, tr As Transaction,
        spanRight As Double, topY As Double, botY As Double,
        wCm As Double, hCm As Double,
        info As ReinforcementBarAutomation.BeamExcelRow,
        Optional addlBelowUpper As Double = 20.0,
        Optional baseXOverride As Nullable(Of Double) = Nothing,
        Optional side As ReinforcementBarAutomation.BeamSide = ReinforcementBarAutomation.BeamSide.LeftSide,
        Optional beamName As String = "",
        Optional secLetter As String = Nothing,
        Optional coverParam As Double = 3.0,
        Optional drawStandaloneTie As Boolean = False
    )

        ' ----- placement of the section to the right of the plan -----
        Dim gapRight As Double = 60.0
        Dim secW As Double = Math.Max(10.0, wCm)
        Dim secH As Double = Math.Max(10.0, hCm)
        Dim centerY As Double = 0.5 * (topY + botY)

        Dim baseX As Double = If(baseXOverride.HasValue, baseXOverride.Value, spanRight + gapRight)
        Dim baseY As Double = centerY - secH / 2.0

        ' ========= Outline (kr_olcu) =========
        DrawSectionOutline(btr, tr, baseX, baseY, secW, secH)

        ' ========= Covers & radii (cm model) =========
        Dim cover As Double = coverParam

        ' Top quantity (montaj)
        Dim nTop As Integer = Math.Max(1, info.QtyMontaj)

        ' Side-specific Additional (İlave) with fallbacks to side/opposite/upper
        Dim addlQtySide As Integer
        Dim addlPhiSide As Double
        If side = ReinforcementBarAutomation.BeamSide.LeftSide Then
            addlQtySide = If(info.QtyAddlLeft > 0, info.QtyAddlLeft, If(info.QtyAddlRight > 0, info.QtyAddlRight, info.QtyMontaj))
            addlPhiSide = If(info.PhiAddlLeft > 0, info.PhiAddlLeft, If(info.PhiAddlRight > 0, info.PhiAddlRight, info.PhiMontaj))
        Else
            addlQtySide = If(info.QtyAddlRight > 0, info.QtyAddlRight, If(info.QtyAddlLeft > 0, info.QtyAddlLeft, info.QtyMontaj))
            addlPhiSide = If(info.PhiAddlRight > 0, info.PhiAddlRight, If(info.PhiAddlLeft > 0, info.PhiAddlLeft, info.PhiMontaj))
        End If

        Dim rTop As Double = Math.Max(0.6, info.PhiMontaj / 2.0R)
        Dim rAddl As Double = Math.Max(0.6, addlPhiSide / 2.0R)
        Dim rBot As Double = Math.Max(0.6, info.PhiLower / 2.0R)
        Dim rWeb As Double = Math.Max(0.6, info.PhiWeb / 2.0R)

        ' Bottom quantity
        Dim nBot As Integer = Math.Max(1, info.QtyLower)

        ' ========= STIRRUP inner box at nominal cover =========
        Dim stirrupX As Double = baseX + cover
        Dim stirrupY As Double = baseY + cover
        Dim stirrupW As Double = secW - 2 * cover
        Dim stirrupH As Double = secH - 2 * cover
        Dim stirrupActual As Double = stirrupH + 2 * cover

        ' ========= BAR CENTERLINES (fully inside the tie) =========
        Dim yTopRow As Double = stirrupY + stirrupH - rTop
        Dim yBotRow As Double = stirrupY + rBot

        Dim yLower As Double = stirrupY                 ' stirrup bottom
        Dim yUpper As Double = stirrupY + stirrupH      ' stirrup top (full height)

        ' ========= WEB / STIRRUPS HORIZONTAL LEGS =========
        Dim webQty As Integer = If(info.WebQty > 0, info.WebQty, 2)
        Dim rows As Integer = Math.Max(1, webQty \ 2)
        Dim gapV As Double = stirrupH / (rows + 1)
        Dim cxLeft As Double = stirrupX + rWeb
        Dim cxRight As Double = stirrupX + stirrupW - rWeb
        For i As Integer = 1 To rows
            Dim cy As Double = stirrupY + i * gapV
            DrawCircleOnLayer(btr, tr, cxLeft, cy, rWeb, LYR_REBAR1)
            DrawCircleOnLayer(btr, tr, cxRight, cy, rWeb, LYR_REBAR1)
        Next

        ' ========= Longitudinal bars (Top) =========
        Dim topXs As List(Of Double) = RowXsTouchingCorners(nTop, baseX, secW, cover, rTop)
        For Each xx As Double In topXs
            DrawCircleOnLayer(btr, tr, xx, yTopRow, rTop, LYR_REBAR1)
        Next

        ' ========= Longitudinal bars (Bottom) =========
        Dim botXs As List(Of Double) = RowXsTouchingCorners(nBot, baseX, secW, cover, rBot)
        For Each xx As Double In botXs
            DrawCircleOnLayer(btr, tr, xx, yBotRow, rBot, LYR_REBAR1)
        Next

        ' ========= SECTION TIE BARS (from Excel only) =========
        Dim spec As String = TieBarsHelper.GetTieSpecFromRow(info)
        Dim nAll As Integer = 0, phiTie As Double = 0.0R, pitch As Double = 0.0R

        If Not String.IsNullOrWhiteSpace(spec) AndAlso TieBarsHelper.TryParseTieSpec(spec, nAll, phiTie, pitch) Then
            Dim tieCount As Integer = Math.Max(0, nAll - 2)    ' number1 includes the two stirrups
            If tieCount > 0 AndAlso topXs IsNot Nothing AndAlso botXs IsNot Nothing AndAlso
               topXs.Count > 0 AndAlso botXs.Count > 0 Then

                ' pick columns present on both rows (or fall back to one row)
                Dim tol As Double = 0.001
                Dim colXs As New List(Of Double)
                For Each tx In topXs
                    If botXs.Any(Function(bx) Math.Abs(bx - tx) < tol) Then colXs.Add(tx)
                Next
                If colXs.Count = 0 Then colXs = If(topXs.Count <= botXs.Count, New List(Of Double)(topXs), New List(Of Double)(botXs))
                If colXs.Count = 0 Then GoTo DoneTies

                ' order: middle → left → right → left → right, corners last
                colXs.Sort()
                Dim n As Integer = colXs.Count
                Dim cxMid As Double = baseX + secW / 2.0

                Dim ordered As New List(Of Double)()
                If n <= 2 Then
                    If n = 1 Then
                        ordered.Add(colXs(0))
                    Else
                        Dim dL = Math.Abs(colXs(0) - cxMid)
                        Dim dR = Math.Abs(colXs(1) - cxMid)
                        If dL <= dR Then ordered.Add(colXs(0)) : ordered.Add(colXs(1)) Else ordered.Add(colXs(1)) : ordered.Add(colXs(0))
                    End If
                Else
                    Dim used() As Boolean = New Boolean(n - 1) {}
                    Dim isCorner As Func(Of Integer, Boolean) = Function(i) (i = 0 OrElse i = n - 1)

                    Dim r0 As Integer = colXs.FindIndex(Function(x) x >= cxMid) : If r0 < 0 Then r0 = n
                    Dim l0 As Integer = r0 - 1

                    Dim midIdx As Integer = -1
                    If r0 < n AndAlso Math.Abs(colXs(r0) - cxMid) <= tol Then midIdx = r0
                    If midIdx = -1 AndAlso l0 >= 0 AndAlso Math.Abs(colXs(l0) - cxMid) <= tol Then midIdx = l0

                    Dim push As Action(Of Integer) =
                        Sub(i As Integer)
                            If i >= 0 AndAlso i < n AndAlso Not used(i) Then
                                ordered.Add(colXs(i)) : used(i) = True
                            End If
                        End Sub

                    If midIdx <> -1 Then push(midIdx)

                    Dim L As Integer, R As Integer, takeLeft As Boolean
                    If midIdx <> -1 Then
                        L = midIdx - 1 : R = midIdx + 1 : takeLeft = True
                    Else
                        L = l0 : R = r0 : takeLeft = True
                    End If

                    While True
                        Dim progressed As Boolean = False
                        If takeLeft Then
                            While L >= 0 AndAlso (used(L) OrElse isCorner(L)) : L -= 1 : End While
                            If L >= 0 Then push(L) : L -= 1 : progressed = True
                        Else
                            While R < n AndAlso (used(R) OrElse isCorner(R)) : R += 1 : End While
                            If R < n Then push(R) : R += 1 : progressed = True
                        End If
                        takeLeft = Not takeLeft
                        If Not progressed Then Exit While
                    End While

                    If Not used(0) Then push(0)
                    If Not used(n - 1) Then push(n - 1)
                End If

                Dim xMin As Double = stirrupX
                Dim xMax As Double = stirrupX + stirrupW
                Dim nearestX As Func(Of List(Of Double), Double, Double) =
                    Function(xs As List(Of Double), x0 As Double) As Double
                        Dim best = xs(0), d = Math.Abs(best - x0)
                        For Each v In xs
                            Dim dd = Math.Abs(v - x0)
                            If dd < d Then best = v : d = dd
                        Next
                        Return best
                    End Function

                Dim placed As Integer = 0
                For Each xCol In ordered
                    If placed >= tieCount Then Exit For

                    Dim xTopC = nearestX(topXs, xCol)
                    Dim xBotC = nearestX(botXs, xCol)

                    ' tangency at right sides of bars (outerwall)
                    Dim xTieTop As Double = Math.Max(xMin, Math.Min(xMax, xTopC + rTop))
                    Dim xTieBot As Double = Math.Max(xMin, Math.Min(xMax, xBotC + rBot))

                    Dim pTop As New Point3d(xTieTop, yTopRow, 0)
                    Dim pBot As New Point3d(xTieBot, yBotRow, 0)
                    Dim mainLn As New Line(pTop, pBot) : mainLn.ColorIndex = 256
                    AppendOnLayer(btr, tr, mainLn, LYR_REBAR2, True)

                    ' 60° leg from above-left into the TOP tangent
                    Const LEG_ANGLE_DEG As Double = 60.0
                    Dim ang As Double = Math.PI * LEG_ANGLE_DEG / 180.0
                    Dim c As Double = Math.Cos(ang)
                    Dim s As Double = Math.Sin(ang)

                    Dim nX As Double = -s
                    Dim nY As Double = c
                    Dim dOutX As Double = -c
                    Dim dOutY As Double = -s

                    Dim tX As Double = xTopC + rTop * nX
                    Dim tY As Double = yTopRow + rTop * nY
                    Dim pTan As New Point3d(tX, tY, 0)

                    Dim leg1 As Double = Math.Max(12.0R, rTop * 0.9)
                    Dim leg2 As Double = 12.0R
                    AdjustLegPairToMultipleOf5(stirrupActual, leg1, leg2)

                    Dim margin As Double = 0.05
                    Dim maxLenLeft As Double = (tX - (xMin + margin)) / Math.Max(0.000000001, c)
                    Dim useLen As Double = Math.Max(0.0R, Math.Min(leg1, maxLenLeft))

                    Dim pStart As New Point3d(tX + dOutX * useLen, tY + dOutY * useLen, 0)

                    Dim legLn As New Line(pStart, pTan) : legLn.ColorIndex = 256
                    AppendOnLayer(btr, tr, legLn, LYR_REBAR2, True)

                    ' ===

                    ' ===

                    placed += 1
                Next
            End If
        End If
DoneTies:

        ' ---------- ADDITIONAL (ila.) — ENFORCE MINIMUM CLEAR; SELECTIVE RELOCATION ----------
        Dim RequiredMinClear As Func(Of Double, Double, Double) = Function(a As Double, b As Double) Math.Max(2.7, Math.Max(a, b))
        Const SECOND_ROW_DROP As Double = 3.0
        Const EPSX As Double = 0.0001

        Dim yAddRow1 As Double = stirrupY + stirrupH - rAddl
        Dim yAddRow2 As Double = Math.Max(stirrupY + rAddl, yAddRow1 - SECOND_ROW_DROP)

        Dim drawAddl As Action(Of Double, Double) =
            Sub(cx As Double, cy As Double)
                Dim c As New Circle(New Point3d(cx, cy, 0), Vector3d.ZAxis, Math.Max(0.1, rAddl))
                c.Color = Color.FromRgb(255, 0, 0) ' red for ADDL row, as in your snippet
                AppendOnLayer(btr, tr, c, LYR_REBAR1, False)
            End Sub

        Dim addlPlacedRow1 As New List(Of Double)()
        Dim addlPlacedRow2 As New List(Of Double)()

        Dim needed As Integer = Math.Max(0, addlQtySide)
        Dim idxAddl As Integer = 0

        If topXs.Count >= 2 AndAlso needed > 0 Then
            Dim gapCount As Integer = topXs.Count - 1

            ' Visiting gaps: left, right, next-left, next-right, …
            Dim order As New List(Of Integer)()
            Dim L As Integer = 0
            Dim R As Integer = gapCount - 1
            While L <= R
                order.Add(L) : L += 1
                If L <= R Then order.Add(R) : R -= 1
            End While

            For Each g In order
                If idxAddl >= needed Then Exit For
                Dim xL As Double = topXs(g)
                Dim xR As Double = topXs(g + 1)

                Dim reqL As Double = RequiredMinClear(info.PhiMontaj, addlPhiSide)
                Dim reqR As Double = RequiredMinClear(info.PhiMontaj, addlPhiSide)

                Dim leftBound As Double = Math.Max(stirrupX + rAddl + EPSX, xL + rTop + rAddl + reqL + EPSX)
                Dim rightBound As Double = Math.Min(stirrupX + stirrupW - rAddl - EPSX, xR - rTop - rAddl - reqR - EPSX)
                If rightBound <= leftBound Then Continue For

                Dim reqAA As Double = RequiredMinClear(addlPhiSide, addlPhiSide)
                Dim stepCC As Double = 2 * rAddl + reqAA

                Dim xNext As Double = leftBound
                While idxAddl < needed AndAlso xNext <= rightBound + 0.000000001
                    If Math.Abs(xNext - xL) < EPSX Then xNext += EPSX
                    If Math.Abs(xNext - xR) < EPSX Then Exit While

                    addlPlacedRow1.Add(xNext)
                    idxAddl += 1
                    xNext += stepCC
                End While
            Next
        End If

        ' Remaining ADDL → Row-2 (preserve order); keep inside stirrup & avoid TOP X
        If idxAddl < needed Then
            Dim reqAA2 As Double = RequiredMinClear(addlPhiSide, addlPhiSide)
            Dim stepCC2 As Double = 2 * rAddl + reqAA2
            Dim left2 As Double = stirrupX + rAddl + EPSX
            Dim right2 As Double = stirrupX + stirrupW - rAddl - EPSX
            If right2 > left2 Then
                Dim lp As Double = left2
                Dim rp As Double = right2
                Dim pickLeft As Boolean = True
                While idxAddl < needed AndAlso lp <= rp + 0.000000001
                    If pickLeft Then
                        If lp <= rp + 0.000000001 Then
                            addlPlacedRow2.Add(lp)
                            idxAddl += 1
                            lp += stepCC2
                        End If
                    Else
                        If rp >= lp - 0.000000001 Then
                            addlPlacedRow2.Add(rp)
                            idxAddl += 1
                            rp -= stepCC2
                        End If
                    End If
                    pickLeft = Not pickLeft
                End While
            End If
        End If

        ' ---------- Labels on the right ----------
        For Each xx In addlPlacedRow1 : drawAddl(xx, yAddRow1) : Next
        For Each xx In addlPlacedRow2 : drawAddl(xx, yAddRow2) : Next

        Dim labelX As Double = baseX + secW + 8.0
        DrawPhiNote(btr, tr, labelX, stirrupY + stirrupH / 2.0, QtyPhiText(webQty, info.PhiWeb))
        DrawPhiNote(btr, tr, labelX, yTopRow, QtyPhiText(nTop, info.PhiMontaj))
        DrawPhiNote(btr, tr, labelX, yTopRow - 8.0, QtyPhiTextAdditional(addlQtySide, addlPhiSide))
        DrawPhiNote(btr, tr, labelX, yBotRow, QtyPhiText(nBot, info.PhiLower))

        ' ========= Draw the stirrup (order no longer matters—bars are inside) =========
        DrawRectOnLayer(btr, tr, stirrupX, stirrupY, stirrupW, stirrupH, LYR_REBAR1)

        ' ---- Beam/Section title under the section ----
        Dim title As String = ""
        If Not String.IsNullOrEmpty(secLetter) Then
            title = secLetter & "-" & secLetter & " KESITI"
        ElseIf Not String.IsNullOrWhiteSpace(beamName) Then
            title = beamName & If(side = ReinforcementBarAutomation.BeamSide.LeftSide, " - Sol Kesit", " - Sağ Kesit")
        End If

        If title <> "" Then
            Dim tName As New MText() With {
                .Location = New Point3d(baseX + secW / 2.0, baseY - 6.0, 0),
                .Attachment = AttachmentPoint.TopCenter,
                .TextHeight = 3.333,
                .Contents = title
            }
            AppendOnLayer(btr, tr, tName, LYR_PHI, True)
        End If
    End Sub

End Module