Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Module PosTiebar

    ' Dedicated layer for the separate symbol. We mirror the color of the sections' rebar layer.
    Public Const LYR_POS_TIEBAR As String = "POS_TIEBAR"

    ' -------------------- Public API --------------------

    ' (A) Easiest to call: pass the same inputs you already have for each section (including Φs).
    Public Sub DrawSeparateTieUnderSection(
        btr As BlockTableRecord, tr As Transaction,
        baseX As Double, secW As Double,
        topY As Double, bottomY As Double,
        phiMontaj As Double, phiLower As Double, phiWeb As Double,
        tieBarsSpec As String,
        Optional gapDown As Double = 200.0,
        Optional slantDeg As Double = 315.0,
        Optional textStyleName As String = "1.25_3.75_ULK",
        Optional textHeight As Double = 3.75)

        ' Covers computed exactly like SectionsModule (your screenshot).
        Dim rTop As Double = Math.Max(0.6, phiMontaj / 2.0R)
        Dim rBot As Double = Math.Max(0.6, phiLower / 2.0R)
        Dim rWeb As Double = Math.Max(0.6, phiWeb / 2.0R)

        Dim tc = ComputeFromSectionInputs(baseX, secW, topY, bottomY, rTop, rBot, rWeb, tieBarsSpec)
        DrawSeparateTie(btr, tr, tc, gapDown, slantDeg, textStyleName, textHeight)
    End Sub

    ' (B) If you already have the covers, use this overload.
    Public Sub DrawSeparateTieWithCovers(
        btr As BlockTableRecord, tr As Transaction,
        baseX As Double, secW As Double,
        topY As Double, bottomY As Double,
        rTop As Double, rBot As Double, rWeb As Double,
        tieBarsSpec As String,
        Optional gapDown As Double = 200.0,
        Optional slantDeg As Double = 315.0,
        Optional textStyleName As String = "1.25_3.75_ULK",
        Optional textHeight As Double = 3.75)

        Dim tc = ComputeFromSectionInputs(baseX, secW, topY, bottomY, rTop, rBot, rWeb, tieBarsSpec)
        DrawSeparateTie(btr, tr, tc, gapDown, slantDeg, textStyleName, textHeight)
    End Sub

    ' -------------------- Compute (pure) --------------------

    Public Structure TieCalc
        Public XLeft As Double, XRight As Double
        Public YLower As Double, YUpper As Double
        Public Leg1 As Double, Leg2 As Double, StirrupH As Double
    End Structure

    Public Function ComputeFromSectionInputs(
    baseX As Double, secW As Double,
    topY As Double, bottomY As Double,
    secHcm As Double, coverParam As Double,
    rWeb As Double,
    tieBarsSpec As String) As TieCalc

        Dim res As New TieCalc()

        ' --- section box & inner stirrup box (exactly like the section drawer) ---
        Dim secH As Double = Math.Max(10.0, secHcm)
        Dim cover As Double = coverParam

        Dim centerY As Double = 0.5 * (topY + bottomY)
        Dim baseY As Double = centerY - secH / 2.0

        Dim stirrupX As Double = baseX + cover
        Dim stirrupY As Double = baseY + cover
        Dim stirrupW As Double = secW - 2 * cover
        Dim innerBoxH As Double = secH - 2 * cover     ' <<< THIS is the actual stirrup height

        ' Horizontal placement for the separate symbol (use inner faces like bars do)
        res.XLeft = stirrupX + rWeb
        res.XRight = stirrupX + stirrupW - rWeb

        ' Vertical extents of the inner box (used to place the symbol)
        res.YLower = stirrupY
        res.YUpper = stirrupY + innerBoxH

        ' <<< Set the PosTiebar stirrupH to the height of THE stirrup >>>
        res.StirrupH = Math.Max(0.0, innerBoxH)

        ' Legs as before (min 12 + /5 rule)
        ParseLegs(tieBarsSpec, res.Leg1, res.Leg2, 12)
        Dim total As Integer = CInt(Math.Round(res.StirrupH)) + CInt(Math.Round(res.Leg1)) + CInt(Math.Round(res.Leg2))
        Dim rem5 As Integer = total Mod 5
        Dim toggle As Boolean = True
        While rem5 <> 0
            If toggle Then res.Leg2 += 1 Else res.Leg1 += 1
            If res.Leg2 < 12 Then res.Leg2 = 12
            If res.Leg1 < 12 Then res.Leg1 = 12
            total = CInt(Math.Round(res.StirrupH)) + CInt(Math.Round(res.Leg1)) + CInt(Math.Round(res.Leg2))
            rem5 = total Mod 5
            toggle = Not toggle
        End While

        Return res
    End Function

    ' -------------------- Draw (entities + plain text) --------------------

    Public Sub DrawSeparateTie(
        btr As BlockTableRecord, tr As Transaction,
        tc As TieCalc,
        Optional gapDown As Double = 200.0,
        Optional slantDeg As Double = 315.0,
        Optional textStyleName As String = "1.25_3.75_ULK",
        Optional textHeight As Double = 3.75)

        ' Create/refresh POS_TIEBAR and make its color match the main rebar layer.
        EnsurePosLayerMatchesRebar(btr.Database, tr)

        Dim xMid As Double = 0.5 * (tc.XLeft + tc.XRight)
        Dim baseY As Double = tc.YLower - gapDown

        Dim pBot As New Point3d(xMid, baseY, 0)
        Dim pBotLeft As New Point3d(xMid - tc.Leg1, baseY, 0)
        Dim pTop As New Point3d(xMid, baseY + tc.StirrupH, 0)

        Dim ang As Double = slantDeg * Math.PI / 180.0
        Dim dx As Double = Math.Cos(ang) * tc.Leg2
        Dim dy As Double = Math.Sin(ang) * tc.Leg2
        Dim pTopSlant As New Point3d(pTop.X - dx, pTop.Y - dy, 0)


        ' Lines — ByLayer color on POS_TIEBAR layer (which mirrors LYR_REBAR1 color).
        AppendLine(btr, tr, pBotLeft, pBot)
        AppendLine(btr, tr, pBot, pTop)
        AppendLine(btr, tr, pTop, pTopSlant)

        ' Optional text style lookup.
        Dim txtStyleId As ObjectId = ObjectId.Null
        Using tst As TextStyleTable = CType(tr.GetObject(btr.Database.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
            If tst.Has(textStyleName) Then txtStyleId = tst(textStyleName)
        End Using

        ' Label offsets (in drawing units).
        Dim offY As Double = textHeight * 1.2
        Dim offX As Double = textHeight * 1.6

        ' --- leg1 text
        Dim leg1Pos As New Point3d(pBot.X - 0.5 * offX, pBot.Y - offY, 0)
        CreateText(btr, tr, leg1Pos, FormatNum(tc.Leg1), textHeight, txtStyleId, colorIndex:=7)

        ' --- stirrupH text: to the right of the vertical stem 
        Dim midVert As New Point3d(pBot.X + offX, (pBot.Y + pTop.Y) * 0.5, 0)
        CreateText(btr, tr, midVert, FormatNum(tc.StirrupH), textHeight, txtStyleId, colorIndex:=7)

        ' --- leg2 text: at the TIP of leg2 (slanted leg end point)
        Dim leg2Pos As New Point3d((pTopSlant.X - 1.0), (pTopSlant.Y - 2.0), 0)
        CreateText(btr, tr, leg2Pos, FormatNum(tc.Leg2), textHeight, txtStyleId, colorIndex:=7)
    End Sub

    ' -------------------- Helpers --------------------

    Private Sub ParseLegs(spec As String, ByRef leg1 As Double, ByRef leg2 As Double, Optional minLeg As Double = 12)
        leg1 = minLeg : leg2 = minLeg
        If String.IsNullOrWhiteSpace(spec) Then Exit Sub
        Dim s = spec.Trim().Replace("×", "x").Replace("X", "x")
        Dim parts = s.Split(New Char() {"-"c, "x"c, "/"c, " "c}, StringSplitOptions.RemoveEmptyEntries)
        If parts.Length >= 1 Then Double.TryParse(parts(0), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, leg1)
        If parts.Length >= 2 Then Double.TryParse(parts(1), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, leg2)
        If leg1 < minLeg Then leg1 = minLeg
        If leg2 < minLeg Then leg2 = minLeg
    End Sub

    ' Make POS_TIEBAR exist and mirror the color of LayerManager.LYR_REBAR1.
    Private Sub EnsurePosLayerMatchesRebar(db As Database, tr As Transaction)
        Dim srcColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256)

        Using lt As LayerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            If lt.Has(LayerManager.LYR_REBAR1) Then
                Using src As LayerTableRecord = CType(tr.GetObject(lt(LayerManager.LYR_REBAR1), OpenMode.ForRead), LayerTableRecord)
                    srcColor = src.Color
                End Using
            End If

            If lt.Has(LYR_POS_TIEBAR) Then
                Using tgt As LayerTableRecord = CType(tr.GetObject(lt(LYR_POS_TIEBAR), OpenMode.ForWrite), LayerTableRecord)
                    tgt.Color = srcColor
                End Using
            Else
                lt.UpgradeOpen()
                Dim rec As New LayerTableRecord() With {
                    .Name = LYR_POS_TIEBAR,
                    .Color = srcColor
                }
                lt.Add(rec) : tr.AddNewlyCreatedDBObject(rec, True)
            End If
        End Using
    End Sub

    Private Sub AppendLine(btr As BlockTableRecord, tr As Transaction, p0 As Point3d, p1 As Point3d)
        Dim ln As New Line(p0, p1) With {
            .ColorIndex = 256,         ' ByLayer
            .Layer = LYR_POS_TIEBAR
        }
        btr.AppendEntity(ln) : tr.AddNewlyCreatedDBObject(ln, True)
    End Sub

    Private Sub CreateText(
    btr As BlockTableRecord, tr As Transaction,
    pos As Point3d, txt As String, h As Double,
    txtStyleId As ObjectId,
    Optional colorIndex As Short = 7)   ' default white

        Dim t As New DBText()
        t.SetDatabaseDefaults()
        t.Position = pos
        t.Height = h
        t.TextString = txt
        If Not txtStyleId.IsNull Then t.TextStyleId = txtStyleId
        t.Layer = LYR_POS_TIEBAR
        t.ColorIndex = colorIndex          ' << hard color (white = 7)
        t.HorizontalMode = TextHorizontalMode.TextMid
        t.VerticalMode = TextVerticalMode.TextVerticalMid
        t.AlignmentPoint = pos
        btr.AppendEntity(t) : tr.AddNewlyCreatedDBObject(t, True)
    End Sub

    Private Function FormatNum(v As Double) As String
        Return Math.Round(v, 0).ToString(Globalization.CultureInfo.InvariantCulture)
    End Function

End Module