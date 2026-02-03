
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors

Public Module PlanTriangleFlag

    ' dir: +1 → triangle points RIGHT (use on LeftSide spans)
    '      -1 → triangle points LEFT  (use on RightSide spans)
    Public Function DrawTriangleFlag(
        btr As BlockTableRecord, tr As Transaction,
        poleTopX As Double, poleTopY As Double, dir As Integer,
        Optional poleLen As Double = 28.0,
        Optional triLen As Double = 18.0,
        Optional triHt As Double = 10.0,
        Optional rotate180 As Boolean = False,
        Optional layerName As String = Nothing
    ) As ObjectId

        Dim d As Integer = If(dir >= 0, +1, -1)

        If Not String.IsNullOrEmpty(layerName) Then
            Dim lt = CType(tr.GetObject(btr.Database.LayerTableId, OpenMode.ForRead), LayerTable)
            If Not lt.Has(layerName) Then
                lt.UpgradeOpen()
                Dim rec As New LayerTableRecord() With {.Name = layerName}
                lt.Add(rec) : tr.AddNewlyCreatedDBObject(rec, True)
            End If
        End If

        ' Build points: pole bottom → pole top → apex (into span) → lower base on pole
        Dim pts As New List(Of Point2d) From {
            New Point2d(poleTopX, poleTopY - poleLen),                               ' pole bottom
            New Point2d(poleTopX, poleTopY),                                         ' pole top (triangle base top)
            New Point2d(poleTopX + d * triLen, poleTopY - triHt / 2.0),              ' triangle apex (into span)
            New Point2d(poleTopX, poleTopY - triHt)                                  ' triangle base bottom on pole
        }

        ' Rotate entire shape 180° about the pole top for the bottom row
        If rotate180 Then
            For i = 0 To pts.Count - 1
                Dim px = pts(i).X : Dim py = pts(i).Y
                pts(i) = New Point2d(2 * poleTopX - px, 2 * poleTopY - py)
            Next
        End If

        ' One open GREEN polyline
        Dim pl As New Polyline()
        pl.SetDatabaseDefaults()
        If Not String.IsNullOrEmpty(layerName) Then pl.Layer = layerName
        pl.Color = Color.FromColorIndex(ColorMethod.ByAci, 3) ' green
        For i = 0 To pts.Count - 1
            pl.AddVertexAt(i, pts(i), 0, 0, 0)
        Next

        Dim id = btr.AppendEntity(pl)
        tr.AddNewlyCreatedDBObject(pl, True)
        Return id
    End Function

End Module
