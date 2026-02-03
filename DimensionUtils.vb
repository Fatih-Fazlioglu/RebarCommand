Option Strict On
Option Infer On

' ========= AutoCAD =========
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors


' =====================================================================
' DIMENSION UTILS — overlay RotatedDimensions for rebar views
' =====================================================================
Module DimensionUtils

    ''' <summary>
    ''' Adds a rotated linear dimension between p1 and p2, offset along the
    ''' normal to the segment by `offset`. Returns the ObjectId of the dim.
    ''' </summary>
    Public Function AddOverlayDimLinear(
      btr As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord,
    tr As Autodesk.AutoCAD.DatabaseServices.Transaction,
    p1 As Autodesk.AutoCAD.Geometry.Point3d,
    p2 As Autodesk.AutoCAD.Geometry.Point3d,
    offset As Double,
    layerName As String,
    Optional textOverride As String = Nothing,
    Optional dimStyleName As String = "1.25_MERD_OLCU_ULK",
    Optional textStyleName As String = "1.25_3.75_ULK",
    Optional precision As Integer = 1,             ' << force 1 decimal by default
    Optional textHeight As Double? = Nothing,
    Optional arrowSize As Double? = Nothing,
    Optional textOffset As Double? = Nothing,
    Optional colorAci As Short = 94                 ' dark green by default
) As Autodesk.AutoCAD.DatabaseServices.ObjectId

        If p1.DistanceTo(p2) < 0.000001 Then Return Autodesk.AutoCAD.DatabaseServices.ObjectId.Null

        Dim db = Autodesk.AutoCAD.ApplicationServices.Application.
             DocumentManager.MdiActiveDocument.Database

        ' ---------- geometry ----------
        Dim ang As Double = Math.Atan2(p2.Y - p1.Y, p2.X - p1.X)
        Dim nx As Double = -Math.Sin(ang), ny As Double = Math.Cos(ang)
        Dim mx As Double = 0.5 * (p1.X + p2.X), my As Double = 0.5 * (p1.Y + p2.Y)
        Dim eps As Double = 0.0001
        Dim dimPt As New Autodesk.AutoCAD.Geometry.Point3d(
        mx + nx * (offset + eps), my + ny * (offset + eps), p1.Z)

        ' ---------- ensure target layer & linetype ----------
        If Not String.IsNullOrEmpty(layerName) Then
            LayerManager.EnsureLayer(db, tr, layerName)
        End If
        Dim dotx2Id As Autodesk.AutoCAD.DatabaseServices.ObjectId =
        LayerManager.EnsureLinetypeLoaded(db, tr, "DOTX2")

        ' ---------- resolve target dimstyle (fallback to current) ----------
        Dim dimStyleId As Autodesk.AutoCAD.DatabaseServices.ObjectId = db.Dimstyle
        Dim dst = CType(tr.GetObject(db.DimStyleTableId, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.DimStyleTable)
        If Not String.IsNullOrEmpty(dimStyleName) AndAlso dst.Has(dimStyleName) Then
            dimStyleId = dst(dimStyleName)
        End If

        ' ---------- arrow block: OBLIQUE -> ARCHTICK fallbacks ----------
        Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.BlockTable)
        Dim obliqueId As Autodesk.AutoCAD.DatabaseServices.ObjectId = Autodesk.AutoCAD.DatabaseServices.ObjectId.Null
        For Each n In New String() {"_OBLIQUE", "OBLIQUE", "Oblique", "_ARCHTICK", "ARCHTICK", "Architectural tick", "_ArchTick"}
            If bt.Has(n) Then obliqueId = bt(n) : Exit For
        Next

        ' ---------- text style (optional) ----------
        Dim textStyleId As Autodesk.AutoCAD.DatabaseServices.ObjectId = Autodesk.AutoCAD.DatabaseServices.ObjectId.Null
        If Not String.IsNullOrEmpty(textStyleName) Then
            Dim tst = CType(tr.GetObject(db.TextStyleTableId, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.TextStyleTable)
            If tst.Has(textStyleName) Then textStyleId = tst(textStyleName)
        End If

        ' ---------- tweak style record (1 decimal, colors, linetypes, arrows) ----------
        Dim dsr = CType(tr.GetObject(dimStyleId, OpenMode.ForWrite), Autodesk.AutoCAD.DatabaseServices.DimStyleTableRecord)
        dsr.Dimtad = 1            ' text above dimension line

        dsr.Dimdec = precision      ' << ensure 1 decimal on the style
        dsr.Dimzin = 0              ' show trailing zeros (e.g., 80.0)

        If Not obliqueId.IsNull Then
            dsr.Dimblk = obliqueId
            dsr.Dimblk1 = obliqueId
            dsr.Dimblk2 = obliqueId
        End If

        If Not dotx2Id.IsNull Then
            dsr.Dimltex1 = dotx2Id          ' extension line 1 linetype
            dsr.Dimltex2 = dotx2Id          ' extension line 2 linetype
            dsr.Dimse1 = False : dsr.Dimse2 = False
        End If

        Dim styleClr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorAci)
        dsr.Dimclrd = styleClr              ' dim line color
        dsr.Dimclre = styleClr              ' extension line color
        dsr.Dimclrt = styleClr              ' text color

        If Not textStyleId.IsNull Then dsr.Dimtxsty = textStyleId

        ' ---------- build the dimension entity ----------
        Dim dimText As String = If(String.IsNullOrEmpty(textOverride), "<>", textOverride)

        Dim d As New Autodesk.AutoCAD.DatabaseServices.RotatedDimension()
        d.SetDatabaseDefaults()
        d.Rotation = ang
        d.XLine1Point = p1
        d.XLine2Point = p2
        d.DimLinePoint = dimPt
        d.DimensionText = dimText
        d.DimensionStyle = dimStyleId
        d.Layer = layerName
        d.Normal = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis
        d.TextPosition = dimPt
        d.Color = styleClr
        If Not textStyleId.IsNull Then d.TextStyleId = textStyleId

        ' --- entity-level numeric overrides (lock to 1 decimal) ---
        d.Dimdec = precision          ' << force 1 decimal on the entity as well
        d.Dimzin = 0                  ' show trailing zeros (80.0)
        d.Dimtad = 1

        If textHeight.HasValue Then d.Dimtxt = textHeight.Value
        If arrowSize.HasValue Then d.Dimasz = arrowSize.Value
        If textOffset.HasValue Then d.Dimgap = textOffset.Value
        ' If desired later: d.Dimrnd = 5.0 ' round to nearest 5 units

        btr.AppendEntity(d)
        tr.AddNewlyCreatedDBObject(d, True)
        d.RecomputeDimensionBlock(True)

        Return d.ObjectId
    End Function
    ''' <summary>
    ''' Convenience wrapper: adds dim only if `show=True`.
    ''' Places above/below by flipping the offset sign.
    ''' </summary>
    Public Sub DrawDimIf(
        btr As BlockTableRecord, tr As Transaction, show As Boolean,
        x0 As Double, x1 As Double, yBar As Double,
        label As String, aci As Short,
        Optional placeAbove As Boolean = True
    )

        If Not show Then Return

        Dim p1 As New Point3d(x0, yBar, 0.0)
        Dim p2 As New Point3d(x1, yBar, 0.0)

        ' Choose a small default offset (you were passing 0.0; keep that if you prefer)
        Dim baseOffset As Double = 0.0
        Dim signedOffset As Double = If(placeAbove, baseOffset, -baseOffset)

        AddOverlayDimLinear(
            btr, tr, p1, p2, signedOffset, LYR_DIM,
            textOverride:=If(String.IsNullOrWhiteSpace(label), Nothing, label),
            dimStyleName:="1.25_MERD_OLCU_ULK",
            textStyleName:="1.25_7.50_ULK",
            precision:=0, textHeight:=7.5, arrowSize:=0.15, textOffset:=0.75,
            colorAci:=aci
        )
    End Sub

End Module