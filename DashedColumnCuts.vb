Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Module DashedColumnCuts

    ' ----------------------------- low-level helpers -----------------------------

    ''' <summary>Ensure a linetype is present; returns its ObjectId or Null if not found.</summary>
    Private Function EnsureOrLoadLinetype(db As Database, tr As Transaction, name As String) As ObjectId
        Dim ltTbl = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)
        If ltTbl.Has(name) Then Return ltTbl(name)

        Try : db.LoadLineTypeFile(name, "acad.lin") : Catch : End Try
        Try
            ' Reopen & recheck after attempting load
            ltTbl = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead, True), LinetypeTable)
            If ltTbl.Has(name) Then Return ltTbl(name)
        Catch
        End Try

        ' (Optional) try acadiso.lin as a fallback
        Try : db.LoadLineTypeFile(name, "acadiso.lin") : Catch : End Try
        Try
            ltTbl = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead, True), LinetypeTable)
            If ltTbl.Has(name) Then Return ltTbl(name)
        Catch
        End Try

        Return ObjectId.Null
    End Function

    ''' <summary>Create layer if missing and set color + linetype.</summary>
    Private Sub EnsureLayerWithColorAndLt(
        db As Database, tr As Transaction, layerName As String,
        ByVal colorAci As Short, ByVal ltName As String)

        Dim ltId As ObjectId = EnsureOrLoadLinetype(db, tr, ltName)

        Dim lyrTbl = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
        If Not lyrTbl.Has(layerName) Then
            lyrTbl.UpgradeOpen()
            Dim ltr As New LayerTableRecord() With {.Name = layerName}
            ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorAci)
            If ltId.IsValid Then ltr.LinetypeObjectId = ltId
            lyrTbl.Add(ltr) : tr.AddNewlyCreatedDBObject(ltr, True)
        Else
            Dim lyrId = lyrTbl(layerName)
            Dim ltr = CType(tr.GetObject(lyrId, OpenMode.ForWrite), LayerTableRecord)
            ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorAci)
            If ltId.IsValid Then ltr.LinetypeObjectId = ltId
        End If
    End Sub

    ''' <summary>Add entity to BTR on a layer; force ByLayer color/linetype if requested.</summary>
    Private Sub AppendEntityOnLayer(
        btr As BlockTableRecord, tr As Transaction, ent As Entity,
        layerName As String, Optional forceByLayer As Boolean = True)

        If forceByLayer Then
            ent.ColorIndex = 256          ' ByLayer color
            ent.Linetype = "ByLayer"      ' ByLayer linetype
        End If
        ent.Layer = layerName
        btr.AppendEntity(ent) : tr.AddNewlyCreatedDBObject(ent, True)
    End Sub

    ' --------------------------------- drawing helpers ---------------------------------

    Private Sub DrawOneDashed(
        btr As BlockTableRecord, tr As Transaction,
        x As Double, yTop As Double, yBottom As Double,
        Optional layer As String = "kr_mesnet-iz", Optional ltscale As Double = 1.0)

        Dim ln As New Line(New Point3d(x, yTop, 0), New Point3d(x, yBottom, 0)) With {
            .LinetypeScale = ltscale
        }
        AppendEntityOnLayer(btr, tr, ln, layer, True)   ' ByLayer → uses layer ACAD_ISO07W100 + orange
    End Sub

    ' ----------------------------------- public APIs -----------------------------------

    ' A) Two columns (left/right), column bottoms share same Y
    Public Sub DrawColumnBottomDashedCuts(
        btr As BlockTableRecord, tr As Transaction,
        xA As Double, xB As Double,
        colBottomY As Double, dupLowerY As Double, extraDownCm As Double,
        Optional layerName As String = "kr_mesnet-iz")

        Dim db = btr.Database
        ' Make the layer ORANGE and ACAD_ISO07W100
        EnsureLayerWithColorAndLt(db, tr, layerName, colorAci:=30, ltName:="ACAD_ISO07W100")

        Dim yTo As Double = dupLowerY - Math.Abs(extraDownCm)
        DrawOneDashed(btr, tr, xA, colBottomY, yTo, layerName)
        DrawOneDashed(btr, tr, xB, colBottomY, yTo, layerName)
    End Sub

    ' B) Multiple columns with same bottom Y
    Public Sub DrawColumnBottomDashedCuts(
        btr As BlockTableRecord, tr As Transaction,
        columnXs As IList(Of Double),
        yColumnBottom As Double, yLowerRebarRow As Double,
        Optional dropBelowLowerCm As Double = 10.0,
        Optional layer As String = "kr_mesnet-iz",
        Optional dashScale As Double = 1.0)

        If columnXs Is Nothing OrElse columnXs.Count = 0 Then Return

        Dim db = btr.Database
        EnsureLayerWithColorAndLt(db, tr, layer, colorAci:=30, ltName:="ACAD_ISO07W100")

        Dim yEnd As Double = yLowerRebarRow - Math.Abs(dropBelowLowerCm)
        For Each x In columnXs
            DrawOneDashed(btr, tr, x, yColumnBottom, yEnd, layer, dashScale)
        Next
    End Sub

    ' C) Per-column bottom Y
    Public Sub DrawColumnBottomDashedCuts(
        btr As BlockTableRecord, tr As Transaction,
        columnXtoBottomY As IList(Of Tuple(Of Double, Double)),
        yLowerRebarRow As Double,
        Optional dropBelowLowerCm As Double = 10.0,
        Optional layer As String = "kr_mesnet-iz",
        Optional dashScale As Double = 1.0)

        If columnXtoBottomY Is Nothing OrElse columnXtoBottomY.Count = 0 Then Return

        Dim db = btr.Database
        EnsureLayerWithColorAndLt(db, tr, layer, colorAci:=30, ltName:="ACAD_ISO07W100")

        Dim yEnd As Double = yLowerRebarRow - Math.Abs(dropBelowLowerCm)
        For Each t In columnXtoBottomY
            DrawOneDashed(btr, tr, t.Item1, t.Item2, yEnd, layer, dashScale)
        Next
    End Sub

End Module