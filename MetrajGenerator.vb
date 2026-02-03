' ===========================================
' MetrajGenerator.vb
' Generates Quantity Takeoff / Bill of Materials Table
' based on POZ_MARKER blocks found in the drawing.
' ===========================================
Option Strict On
Option Infer On

Imports System.Globalization
Imports System.Collections.Generic
Imports System.Linq
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.Runtime

Public Class MetrajGenerator

    ' Standard Unit Weights (Kg/m) for Rebar
    Private Shared ReadOnly UnitWeights As New Dictionary(Of Integer, Double) From {
       {6, 0.222}, {8, 0.395}, {10, 0.617}, {12, 0.888}, {14, 1.209},
        {16, 1.58}, {18, 2}, {20, 2.469}, {22, 2.987},
        {24, 3.55}, {25, 3.858}, {26, 4.17}, {28, 4.839},
        {30, 5.55}, {32, 6.32}
    }

    ' Table Layout Constants (Increased for "Bigger Chambers")
    Private Const ROW_HEIGHT As Double = 35.0
    Private Const COL_WIDTH_STD As Double = 80.0
    Private Const COL_WIDTH_SHAPE As Double = 200.0 ' Widened for better shape visibility
    Private Const COL_WIDTH_DIA As Double = 80.0 ' Changed from 70.0 to 80.0
    Private Const TEXT_HEIGHT_STD As Double = 12.0
    Private Const TEXT_HEIGHT_TITLE As Double = 15.0

    <CommandMethod("METRAJ_HESAPLA")>
    Public Sub DrawMetrajTableCommand()
        Dim doc = Application.DocumentManager.MdiActiveDocument
        Dim db = doc.Database
        Dim ed = doc.Editor

        ' 1. Select Objects
        Dim tvs() As TypedValue = {
            New TypedValue(CInt(DxfCode.Start), "INSERT"),
            New TypedValue(CInt(DxfCode.LayerName), "poz")
        }
        Dim filter As New SelectionFilter(tvs)
        Dim selOpts As New PromptSelectionOptions()
        selOpts.MessageForAdding = vbLf & "Select Rebar Positions (POZ_MARKER) to include in table: "

        Dim selRes = ed.GetSelection(selOpts, filter)

        If selRes.Status <> PromptStatus.OK Then
            ed.WriteMessage(vbLf & "Selection canceled or no objects selected.")
            Return
        End If

        Dim selectedIds As ObjectId() = selRes.Value.GetObjectIds()

        ' 2. Select Insertion Point
        Dim ppo As New PromptPointOptions(vbLf & "Select insertion point for Quantity Table: ")
        Dim ppr = ed.GetPoint(ppo)
        If ppr.Status <> PromptStatus.OK Then Return
        Dim insertPt As Point3d = ppr.Value

        ' 3. Collect Data from Selection
        Dim rebarList As List(Of RebarEntry) = CollectRebarData(db, selectedIds)
        If rebarList.Count = 0 Then
            ed.WriteMessage(vbLf & "No valid 'POZ_MARKER' blocks found in the selection.")
            Return
        End If

        ' --- Determine Dynamic Diameter Columns ---
        ' Find all unique diameters in the selection and sort them
        Dim activeDiameters As List(Of Integer) = rebarList _
                                                  .Select(Function(r) r.Diameter) _
                                                  .Where(Function(d) d > 0) _
                                                  .Distinct() _
                                                  .OrderBy(Function(d) d) _
                                                  .ToList()

        If activeDiameters.Count = 0 Then
            ' Fallback if no diameters found (unlikely if blocks exist)
            activeDiameters.Add(10)
        End If

        ' 4. Draw Table
        Using tr = db.TransactionManager.StartTransaction()
            Dim btr = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

            ' Sort rows by Pos No
            rebarList = rebarList.OrderBy(Function(x) x.PosNo).ToList()

            ' Pass the dynamic list of diameters
            DrawTableGeometry(db, tr, btr, insertPt, rebarList, activeDiameters)

            tr.Commit()
        End Using
    End Sub

    ' Structure to hold extracted data
    Private Class RebarEntry
        Public PosNo As Integer
        Public Diameter As Integer
        Public Count As Integer ' ADET
        Public Multiplier As Integer ' CARPAN
        Public LengthCm As Integer ' BOY
        Public ShapeCode As String ' DEMIR_TIPI
        Public ShapeDims As String ' BUKUM_BOY (Space separated)
        Public BarType As String ' "Etr", "Crz" or Standard

        Public ReadOnly Property TotalCount As Integer
            Get
                Dim m = If(Multiplier <= 0, 1, Multiplier)
                Return Count * m
            End Get
        End Property

        Public ReadOnly Property TotalLengthM As Double
            Get
                Return (TotalCount * LengthCm) / 100.0R
            End Get
        End Property
    End Class

    Private Function CollectRebarData(db As Database, idsToProcess As ObjectId()) As List(Of RebarEntry)
        Dim list As New List(Of RebarEntry)()

        Using tr = db.TransactionManager.StartTransaction()
            ' Iterate over the selected IDs only
            For Each id As ObjectId In idsToProcess
                Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                If ent Is Nothing Then Continue For

                If TypeOf ent Is BlockReference Then
                    Dim br = CType(ent, BlockReference)
                    Dim blkName As String = GetBlockName(br, tr)

                    If blkName.Equals("POZ_MARKER", StringComparison.OrdinalIgnoreCase) AndAlso
                       br.Layer.Equals("poz", StringComparison.OrdinalIgnoreCase) Then

                        Dim entry As New RebarEntry()
                        Dim foundNo As Boolean = False

                        For Each attId As ObjectId In br.AttributeCollection
                            Dim att = CType(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
                            Select Case att.Tag.ToUpperInvariant()
                                Case "NO"
                                    Integer.TryParse(att.TextString, entry.PosNo)
                                    foundNo = True
                                Case "CAP"
                                    Dim cln = System.Text.RegularExpressions.Regex.Replace(att.TextString, "[^0-9]", "")
                                    Integer.TryParse(cln, entry.Diameter)
                                Case "ADET"
                                    Integer.TryParse(att.TextString, entry.Count)
                                Case "CARPAN"
                                    Integer.TryParse(att.TextString, entry.Multiplier)
                                Case "BOY"
                                    Dim cln = System.Text.RegularExpressions.Regex.Replace(att.TextString, "[^0-9]", "")
                                    Integer.TryParse(cln, entry.LengthCm)
                                Case "BUKUM_BOY"
                                    entry.ShapeDims = att.TextString
                                Case "DEMIR_TIPI"
                                    entry.ShapeCode = att.TextString
                                Case "YER"
                                    If att.TextString.Contains("Etr") Then entry.BarType = "Etr"
                                    If att.TextString.Contains("Crz") Then entry.BarType = "Crz"
                            End Select
                        Next

                        If foundNo AndAlso entry.PosNo > 0 Then
                            list.Add(entry)
                        End If
                    End If
                End If
            Next
            tr.Commit()
        End Using

        Dim groupedList = list.GroupBy(Function(x) x.PosNo).Select(Function(g)
                                                                       Dim first = g.First()
                                                                       Dim totalCnt = g.Sum(Function(x) x.TotalCount)
                                                                       Dim combined As New RebarEntry With {
                                                                           .PosNo = first.PosNo,
                                                                           .Diameter = first.Diameter,
                                                                           .LengthCm = first.LengthCm,
                                                                           .ShapeCode = first.ShapeCode,
                                                                           .ShapeDims = first.ShapeDims,
                                                                           .BarType = first.BarType,
                                                                           .Count = totalCnt,
                                                                           .Multiplier = 1
                                                                       }
                                                                       Return combined
                                                                   End Function).ToList()

        Return groupedList
    End Function

    Private Function GetBlockName(br As BlockReference, tr As Transaction) As String
        Dim btr = CType(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
        If btr.IsAnonymous Then
            Dim dynBtrId = br.DynamicBlockTableRecord
            Dim dynBtr = CType(tr.GetObject(dynBtrId, OpenMode.ForRead), BlockTableRecord)
            Return dynBtr.Name
        End If
        Return btr.Name
    End Function

    Private Sub DrawTableGeometry(db As Database, tr As Transaction, btr As BlockTableRecord,
                                  insPt As Point3d, data As List(Of RebarEntry),
                                  diamColumns As List(Of Integer))

        ' --- 1. Define Grid Coordinates ---
        Dim headers As String() = {"POZ NO", "CAP", "ADET", "BOY(cm)", "SEKIL"}
        Dim xOffs As New List(Of Double)()
        Dim curX As Double = 0
        xOffs.Add(curX)
        ' Pos
        curX += COL_WIDTH_STD : xOffs.Add(curX)
        ' Cap
        curX += COL_WIDTH_STD : xOffs.Add(curX)
        ' Count
        curX += COL_WIDTH_STD : xOffs.Add(curX)
        ' Length
        curX += COL_WIDTH_STD : xOffs.Add(curX)
        ' Shape
        curX += COL_WIDTH_SHAPE : xOffs.Add(curX)

        Dim diamStartIdx As Integer = xOffs.Count - 1

        ' Generate columns dynamically based on the passed list
        For Each d In diamColumns
            curX += COL_WIDTH_DIA
            xOffs.Add(curX)
        Next

        Dim totalWidth As Double = curX
        Dim curY As Double = 0

        ' --- 2. Draw Main Header ---
        DrawCell(btr, tr, insPt, 0, curY, totalWidth, ROW_HEIGHT, "METRAJ TABLOSU / QUANTITY TABLE", 2, True)
        curY -= ROW_HEIGHT

        ' --- 3. Draw Column Headers ---
        Dim diaHeaderY As Double = curY
        Dim rowH2 As Double = ROW_HEIGHT * 1.5

        For i = 0 To 4
            Dim w = xOffs(i + 1) - xOffs(i)
            DrawCell(btr, tr, insPt, xOffs(i), curY, w, rowH2, headers(i), 7, False)
        Next

        Dim diaTotalW = totalWidth - xOffs(diamStartIdx)
        DrawCell(btr, tr, insPt, xOffs(diamStartIdx), curY, diaTotalW, ROW_HEIGHT * 0.75, "BOY / LENGTH (m)", 7, True)

        Dim diaSubY = curY - (ROW_HEIGHT * 0.75)
        ' Draw Sub Headers using dynamic diameters
        For i = 0 To diamColumns.Count - 1
            Dim cx_col = xOffs(diamStartIdx + i)
            Dim w = COL_WIDTH_DIA
            DrawCell(btr, tr, insPt, cx_col, diaSubY, w, ROW_HEIGHT * 0.75, "Ø" & diamColumns(i), 7, True)
        Next

        curY -= rowH2

        ' --- 4. Draw Data Rows ---
        Dim colTotals(diamColumns.Count - 1) As Double

        For Each row In data
            Dim yTop = curY
            Dim yBot = curY - ROW_HEIGHT

            DrawLine(btr, tr, insPt, 0, yBot, totalWidth, yBot, 1)

            ' 1. Pos
            DrawText(btr, tr, insPt, xOffs(0) + COL_WIDTH_STD / 2, yTop - ROW_HEIGHT / 2, row.PosNo.ToString(), 7)
            ' 2. Cap
            DrawText(btr, tr, insPt, xOffs(1) + COL_WIDTH_STD / 2, yTop - ROW_HEIGHT / 2, row.Diameter.ToString(), 7)
            ' 3. Adet
            DrawText(btr, tr, insPt, xOffs(2) + COL_WIDTH_STD / 2, yTop - ROW_HEIGHT / 2, row.TotalCount.ToString(), 7)
            ' 4. Boy
            DrawText(btr, tr, insPt, xOffs(3) + COL_WIDTH_STD / 2, yTop - ROW_HEIGHT / 2, row.LengthCm.ToString(), 7)

            ' 5. Shape Diagram
            Dim shapeBoxCenter As New Point3d(insPt.X + xOffs(4) + COL_WIDTH_SHAPE / 2, insPt.Y + yTop - ROW_HEIGHT / 2, 0)
            DrawShapeDiagram(btr, tr, shapeBoxCenter, row)

            ' 6. Diameter Columns
            Dim totalLenM As Double = row.TotalLengthM
            ' Find index in the dynamic list
            Dim dIdx As Integer = diamColumns.IndexOf(row.Diameter)

            If dIdx >= 0 Then
                Dim colX = xOffs(diamStartIdx + dIdx) + COL_WIDTH_DIA / 2
                DrawText(btr, tr, insPt, colX, yTop - ROW_HEIGHT / 2, totalLenM.ToString("F2", CultureInfo.InvariantCulture), 7)
                colTotals(dIdx) += totalLenM
            End If

            curY -= ROW_HEIGHT
        Next

        ' Vertical Grid Lines
        Dim dataHeight = data.Count * ROW_HEIGHT
        For Each x In xOffs
            DrawLine(btr, tr, insPt, x, -rowH2, x, curY, 1)
        Next
        DrawLine(btr, tr, insPt, totalWidth, -rowH2, totalWidth, curY, 1)


        ' --- 5. Footer Rows ---
        Dim fLabels As String() = {"TOPLAM BOY / TOTAL LENGTH (m)", "BIRIM AGIRLIK / UNIT WEIGHT (kg/m)", "AGIRLIK / WEIGHT (kg)", "TOPLAM AGIRLIK / TOTAL WEIGHT (kg)"}
        Dim grandTotalWeight As Double = 0

        For r = 0 To 3
            Dim yTop = curY
            Dim yBot = curY - ROW_HEIGHT

            Dim labelW = xOffs(diamStartIdx)
            DrawCell(btr, tr, insPt, 0, yTop, labelW, ROW_HEIGHT, fLabels(r), 7, False, True)

            For c = 0 To diamColumns.Count - 1
                Dim dia = diamColumns(c)
                Dim valStr As String = ""
                Dim val As Double = 0

                If r = 0 Then ' Total Length
                    val = colTotals(c)
                    valStr = If(val > 0, val.ToString("F2", CultureInfo.InvariantCulture), "")
                ElseIf r = 1 Then ' Unit Weight
                    If UnitWeights.ContainsKey(dia) Then
                        val = UnitWeights(dia)
                    Else
                        ' Fallback approx calculation if diameter not in standard table: d^2 / 162
                        val = (dia * dia) / 162.0
                    End If
                    valStr = val.ToString("F3", CultureInfo.InvariantCulture)
                ElseIf r = 2 Then ' Weight
                    Dim uWt As Double = 0
                    If UnitWeights.ContainsKey(dia) Then
                        uWt = UnitWeights(dia)
                    Else
                        uWt = (dia * dia) / 162.0
                    End If

                    val = colTotals(c) * uWt
                    valStr = If(val > 0, val.ToString("F2", CultureInfo.InvariantCulture), "")
                    grandTotalWeight += val
                End If

                If r < 3 Then
                    Dim cx = xOffs(diamStartIdx + c)
                    DrawCell(btr, tr, insPt, cx, yTop, COL_WIDTH_DIA, ROW_HEIGHT, valStr, 7, True, False)
                End If
            Next

            If r = 3 Then
                Dim restW = totalWidth - labelW
                DrawCell(btr, tr, insPt, labelW, yTop, restW, ROW_HEIGHT, grandTotalWeight.ToString("F2", CultureInfo.InvariantCulture) & " Kg.", 2, True, True)
            End If

            curY -= ROW_HEIGHT
        Next

        DrawLine(btr, tr, insPt, 0, curY + ROW_HEIGHT, totalWidth, curY + ROW_HEIGHT, 1)

    End Sub

    Private Sub DrawCell(btr As BlockTableRecord, tr As Transaction, origin As Point3d,
                         localX As Double, localY As Double, w As Double, h As Double,
                         txt As String, colorIdx As Short, centerText As Boolean, Optional drawBorder As Boolean = True)
        If drawBorder Then
            DrawLine(btr, tr, origin, localX, localY, localX + w, localY, 1)
            DrawLine(btr, tr, origin, localX, localY - h, localX + w, localY - h, 1)
            DrawLine(btr, tr, origin, localX, localY, localX, localY - h, 1)
            DrawLine(btr, tr, origin, localX + w, localY, localX + w, localY - h, 1)
        End If

        If String.IsNullOrEmpty(txt) Then Return

        Dim mt As New MText()
        mt.Contents = txt
        mt.TextHeight = TEXT_HEIGHT_STD
        mt.ColorIndex = colorIdx
        mt.Attachment = If(centerText, AttachmentPoint.MiddleCenter, AttachmentPoint.MiddleLeft)

        Dim tx = If(centerText, localX + w / 2, localX + 2.0)
        Dim ty = localY - h / 2
        mt.Location = New Point3d(origin.X + tx, origin.Y + ty, 0)

        btr.AppendEntity(mt)
        tr.AddNewlyCreatedDBObject(mt, True)
    End Sub

    Private Sub DrawLine(btr As BlockTableRecord, tr As Transaction, origin As Point3d,
                         x1 As Double, y1 As Double, x2 As Double, y2 As Double, col As Short)
        Dim ln As New Line(New Point3d(origin.X + x1, origin.Y + y1, 0),
                           New Point3d(origin.X + x2, origin.Y + y2, 0))
        ln.ColorIndex = col
        btr.AppendEntity(ln)
        tr.AddNewlyCreatedDBObject(ln, True)
    End Sub

    Private Sub DrawText(btr As BlockTableRecord, tr As Transaction, origin As Point3d,
                         x As Double, y As Double, txt As String, col As Short)
        If String.IsNullOrEmpty(txt) Then Return
        Dim mt As New MText()
        mt.Contents = txt
        mt.TextHeight = TEXT_HEIGHT_STD
        mt.ColorIndex = col
        mt.Attachment = AttachmentPoint.MiddleCenter
        mt.Location = New Point3d(origin.X + x, origin.Y + y, 0)
        btr.AppendEntity(mt)
        tr.AddNewlyCreatedDBObject(mt, True)
    End Sub

    ' =========================================================
    ' SHAPE DIAGRAM LOGIC (IMPROVED FOR NON-OVERLAPPING TEXT)
    ' =========================================================
    Private Sub DrawShapeDiagram(btr As BlockTableRecord, tr As Transaction, center As Point3d, row As RebarEntry)
        ' Adjusted for bigger cell and padding
        Const CELL_W As Double = 200.0 ' Match COL_WIDTH_SHAPE
        Const CELL_H As Double = 35.0  ' Match ROW_HEIGHT
        Const MARGIN As Double = 10.0

        Dim MAX_W As Double = CELL_W - (MARGIN * 2)
        Dim MAX_H As Double = CELL_H - (MARGIN * 2)

        Dim dims As List(Of Integer) = ParseDimensions(row.ShapeDims)
        Dim shapeType As Integer = 0
        Integer.TryParse(row.ShapeCode, shapeType)

        Dim pl As New Polyline()
        pl.ColorIndex = 2 ' Yellow
        pl.ConstantWidth = 0.0

        Dim pts As New List(Of Point2d)()
        Dim texts As New List(Of Tuple(Of Point2d, String))()

        ' Mapping:
        ' 1 = Straight
        ' 2 = L-shape (m1 s)
        ' 3 = U-shape (m1 s m2)
        ' 13 = Tie (Crz)
        ' 15 = Stirrup (Etr)

        Select Case shapeType
            Case 15 ' Stirrup (Rect)
                Dim w = MAX_W * 0.4
                Dim h = MAX_H * 0.6
                pts.Add(New Point2d(-w / 2, -h / 2))
                pts.Add(New Point2d(w / 2, -h / 2))
                pts.Add(New Point2d(w / 2, h / 2))
                pts.Add(New Point2d(-w / 2, h / 2))
                pl.Closed = True

                ' Text: Center
                texts.Add(Tuple.Create(New Point2d(0, 0), "L=" & row.LengthCm.ToString()))

            Case 13 ' Tie (Hooked Line - Left Vertical, Right 45 deg)
                Dim w = MAX_W * 0.5
                pts.Add(New Point2d(-w / 2, 5)) ' Left Hook Up
                pts.Add(New Point2d(-w / 2, 0)) ' Left Corner
                pts.Add(New Point2d(w / 2, 0))  ' Right Corner
                pts.Add(New Point2d(w / 2 - 5, 5)) ' Right Hook Up (Slanted 45 Left)

                ' Text above center
                texts.Add(Tuple.Create(New Point2d(0, 8), "L=" & row.LengthCm.ToString()))

            Case 1 ' Straight
                Dim w = MAX_W * 0.6
                pts.Add(New Point2d(-w / 2, 0))
                pts.Add(New Point2d(w / 2, 0))

                Dim val As String = If(dims.Count > 0, dims(0).ToString(), row.LengthCm.ToString())
                texts.Add(Tuple.Create(New Point2d(0, 5), val))

            Case 2 ' L-Shape
                Dim m1 As Double = If(dims.Count > 0, dims(0), 20)
                Dim s As Double = If(dims.Count > 1, dims(1), 100)

                Dim drawW = MAX_W * 0.5
                Dim drawH = MAX_H * 0.5

                pts.Add(New Point2d(-drawW / 2, drawH / 2)) ' Top Left
                pts.Add(New Point2d(-drawW / 2, -drawH / 2)) ' Bottom Left
                pts.Add(New Point2d(drawW / 2, -drawH / 2)) ' Bottom Right

                ' Text offsets: Move away from lines
                ' Left dimension shifted from -8 to -15 as requested
                If dims.Count > 0 Then texts.Add(Tuple.Create(New Point2d(-drawW / 2 - 15, 0), dims(0).ToString())) ' Left of vertical leg
                If dims.Count > 1 Then texts.Add(Tuple.Create(New Point2d(0, -drawH / 2 + 5), dims(1).ToString())) ' Above horizontal leg

            Case 3 ' U-Shape
                Dim drawW = MAX_W * 0.5
                Dim drawH = MAX_H * 0.5

                pts.Add(New Point2d(-drawW / 2, drawH / 2))
                pts.Add(New Point2d(-drawW / 2, -drawH / 2))
                pts.Add(New Point2d(drawW / 2, -drawH / 2))
                pts.Add(New Point2d(drawW / 2, drawH / 2))

                ' Text offsets
                If dims.Count > 0 Then texts.Add(Tuple.Create(New Point2d(-drawW / 2 - 8, 0), dims(0).ToString())) ' Left leg
                If dims.Count > 1 Then texts.Add(Tuple.Create(New Point2d(0, -drawH / 2 + 5), dims(1).ToString())) ' Bottom leg
                If dims.Count > 2 Then texts.Add(Tuple.Create(New Point2d(drawW / 2 + 8, 0), dims(2).ToString())) ' Right leg

            Case Else
                Dim w = MAX_W * 0.6
                pts.Add(New Point2d(-w / 2, 0))
                pts.Add(New Point2d(w / 2, 0))
                texts.Add(Tuple.Create(New Point2d(0, 5), "L=" & row.LengthCm.ToString()))
        End Select

        For i = 0 To pts.Count - 1
            pl.AddVertexAt(i, pts(i), 0, 0, 0)
        Next

        Dim mt As Matrix3d = Matrix3d.Displacement(center.GetAsVector())
        pl.TransformBy(mt)
        btr.AppendEntity(pl)
        tr.AddNewlyCreatedDBObject(pl, True)

        For Each t In texts
            Dim txtEnt As New DBText()
            txtEnt.Height = 8.0
            txtEnt.TextString = t.Item2
            txtEnt.HorizontalMode = TextHorizontalMode.TextMid
            txtEnt.VerticalMode = TextVerticalMode.TextVerticalMid
            txtEnt.AlignmentPoint = New Point3d(center.X + t.Item1.X, center.Y + t.Item1.Y, 0)
            txtEnt.ColorIndex = 7
            btr.AppendEntity(txtEnt)
            tr.AddNewlyCreatedDBObject(txtEnt, True)
        Next

    End Sub

    Private Function ParseDimensions(s As String) As List(Of Integer)
        Dim res As New List(Of Integer)
        If String.IsNullOrWhiteSpace(s) Then Return res
        Dim parts = s.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
        For Each p In parts
            Dim v As Integer
            If Integer.TryParse(p, v) Then res.Add(v)
        Next
        Return res
    End Function

End Class