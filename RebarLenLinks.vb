Imports System.Text.RegularExpressions
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry


' Everything for responsive L= labels lives here
Module RebarLenLinks
    ' 0) Public things the rest of your code needs
    Friend Enum LenKind
        Total = 0
        NoMiter = 1
        MiterLeft = 2
        MiterRight = 3
        OverLay = 4
    End Enum

    Friend Const LENLINK_NOD As String = "ULK_LENLINKS"   ' Named Object Dictionary bucket
    Private _lenUpdaterInstalled As Boolean = False

    ' 2) Save/Read link helpers  <--- put them here
    Friend Sub SaveLenLink(
    tr As Transaction, db As Database,
    labelId As ObjectId, targetId As ObjectId,
    kind As LenKind, segIndex As Integer,
    Optional dx As Double = 0.0, Optional dy As Double = 0.0)


        Dim nod = CType(tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead), DBDictionary)

        Dim linksId As ObjectId
        If Not nod.Contains(LENLINK_NOD) Then
            nod.UpgradeOpen()
            Dim d As New DBDictionary()
            linksId = nod.SetAt(LENLINK_NOD, d)
            tr.AddNewlyCreatedDBObject(d, True)
        Else
            linksId = nod.GetAt(LENLINK_NOD)
        End If

        Dim links = CType(tr.GetObject(linksId, OpenMode.ForWrite), DBDictionary)

        Dim xr As New Xrecord()
        xr.Data = New ResultBuffer(
        New TypedValue(DxfCode.Text, targetId.Handle.ToString()),
        New TypedValue(DxfCode.Int16, CInt(kind)),
        New TypedValue(DxfCode.Int16, segIndex),
        New TypedValue(DxfCode.Real, dx),
        New TypedValue(DxfCode.Real, dy))

        links.SetAt(labelId.Handle.ToString(), xr)
        tr.AddNewlyCreatedDBObject(xr, True)
    End Sub





    ' --- Public helper you call from drawing code (now accepts dx, dy) ---
    Friend Sub LinkLabelTo(
    tr As Transaction, db As Database,
    labelId As ObjectId, targetId As ObjectId,
    kind As LenKind, segIndex As Integer,
    Optional dx As Double = 0.0, Optional dy As Double = 0.0)

        SaveLenLink(tr, db, labelId, targetId, kind, segIndex, dx, dy)
    End Sub
    Friend Function GetLabelsLinkedTo(db As Database, tr As Transaction, targetHandle As String) _
    As List(Of Tuple(Of ObjectId, LenKind, Integer, Double, Double))

        Dim result As New List(Of Tuple(Of ObjectId, LenKind, Integer, Double, Double))()
        Dim nod = CType(tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead), DBDictionary)
        If Not nod.Contains(LENLINK_NOD) Then Return result

        Dim links = CType(tr.GetObject(nod.GetAt(LENLINK_NOD), OpenMode.ForRead), DBDictionary)
        For Each kv In links
            Dim xr = CType(tr.GetObject(kv.Value, OpenMode.ForRead), Xrecord)
            Dim data = xr.Data.AsArray()
            Dim tgt = CStr(data(0).Value)
            If tgt <> targetHandle Then Continue For

            Dim k As LenKind = CType(CInt(data(1).Value), LenKind)
            Dim seg As Integer = CInt(data(2).Value)
            Dim dx As Double = If(data.Length >= 4, CDbl(data(3).Value), 0.0)
            Dim dy As Double = If(data.Length >= 5, CDbl(data(4).Value), 0.0)

            Dim lblHandle As New Handle(Convert.ToInt64(kv.Key, 16))
            Dim lblId As ObjectId
            If db.TryGetObjectId(lblHandle, lblId) Then
                result.Add(Tuple.Create(lblId, k, seg, dx, dy))
            End If
        Next
        Return result
    End Function
    Friend Function PolylineSegmentLength(pl As Polyline, segIdx As Integer) As Double
        If segIdx < 0 Then Return 0.0

        Dim segCount As Integer = pl.NumberOfVertices - If(pl.Closed, 0, 1)
        If segIdx >= segCount Then Return 0.0

        Dim i0 As Integer = segIdx
        Dim i1 As Integer = (segIdx + 1) Mod pl.NumberOfVertices

        Dim p0 As Point2d = pl.GetPoint2dAt(i0)
        Dim p1 As Point2d = pl.GetPoint2dAt(i1)
        Dim c As Double = p0.GetDistanceTo(p1)               ' chord length
        Dim bulge As Double = pl.GetBulgeAt(i0)

        If Math.Abs(bulge) < 0.000000000001 Then
            ' straight segment
            Return c
        Else
            ' arc segment: bulge = tan(theta/4)  ->  theta = 4*atan(bulge)
            Dim theta As Double = 4.0 * Math.Atan(bulge)     ' signed central angle (radians)
            Dim R As Double = (c / 2.0) / Math.Sin(Math.Abs(theta) / 2.0)
            Return Math.Abs(R * theta)                       ' arc length
        End If
    End Function

    Friend Function ComputeLen(ent As Entity, kind As LenKind, segIdx As Integer) As Double
        If TypeOf ent Is Line Then
            Return CType(ent, Line).Length
        ElseIf TypeOf ent Is Polyline Then
            Dim pl = CType(ent, Polyline)
            If kind = LenKind.Total Then Return pl.Length
            Return PolylineSegmentLength(pl, segIdx)
        End If
        Return 0.0
    End Function

    ' ===== Text replacement (keeps your "L=" prefix; uses your formatter) =====
    ' If your global FormatLen is not visible here, use this local one or call the global explicitly.
    Friend Function FormatLenLabel(v As Double) As String
        Return Math.Round(v, 0).ToString()
    End Function

    Friend Function ReplaceFirstNumberKeepingPrefix(existing As String, newVal As Double) As String
        Dim newTxt As String = FormatLenLabel(newVal)  ' or call your public FormatLen(newVal)
        Dim re As New Regex("([0-9]+(\.[0-9]+)?)")
        Return re.Replace(existing, newTxt, 1)
    End Function


    Friend Sub InstallLenUpdater()
        If _lenUpdaterInstalled Then Exit Sub

        Dim dm = Application.DocumentManager
        Dim cur = dm.MdiActiveDocument
        If cur IsNot Nothing Then
            AddHandler cur.CommandEnded, AddressOf OnCommandEnded_Refresh
        End If

        AddHandler dm.DocumentCreated, AddressOf OnDocumentCreated_Hook
        _lenUpdaterInstalled = True
    End Sub

    Private Sub OnDocumentCreated_Hook(sender As Object, e As DocumentCollectionEventArgs)
        If e Is Nothing OrElse e.Document Is Nothing Then Exit Sub
        AddHandler e.Document.CommandEnded, AddressOf OnCommandEnded_Refresh
    End Sub

    Private Sub OnCommandEnded_Refresh(sender As Object, e As CommandEventArgs)
        Dim doc = TryCast(sender, Document)
        If doc Is Nothing Then Exit Sub

        Dim name As String = (If(e?.GlobalCommandName, "")).ToUpperInvariant()
        Select Case name
            Case "STRETCH", "MOVE", "PEDIT", "LENGTHEN", "GRIP_STRETCH", "GRIP_MOVE", "GRIP_ROTATE", "GRIP_SCALE", "ROTATE", "SCALE"
                ' Lock the doc here, then refresh
                Using doc.LockDocument()
                    RefreshAllLenLabelsForDb(doc.Database)  ' <-- lock-free helper below
                End Using
        End Select
    End Sub

    Friend Sub OnObjectModified_UpdateLenLabels(sender As Object, e As ObjectEventArgs)
        Dim db = TryCast(sender, Database)
        If db Is Nothing Then Exit Sub

        Using tr = db.TransactionManager.StartTransaction()
            Dim ent = TryCast(tr.GetObject(e.DBObject.ObjectId, OpenMode.ForRead, False), Entity)
            If ent Is Nothing Then tr.Commit() : Exit Sub
            UpdateLabelsForTarget(db, tr, ent)
            tr.Commit()
        End Using
    End Sub


    Friend Function CreateLengthLabel(
    tr As Transaction, btr As BlockTableRecord,
    ins As Point3d, text As String, textHeight As Double, colorIndex As Short) As ObjectId

        Dim mt As New MText()
        mt.Location = ins
        mt.Attachment = AttachmentPoint.MiddleCenter
        mt.TextHeight = textHeight
        mt.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex)
        mt.Contents = text   ' e.g., "L= 350" from your existing FormatLen
        btr.AppendEntity(mt) : tr.AddNewlyCreatedDBObject(mt, True)
        Return mt.ObjectId
    End Function



    Friend Sub UpdateLabelsForTarget(db As Database, tr As Transaction, ent As Entity)
        Dim linked = GetLabelsLinkedTo(db, tr, ent.Handle.ToString())
        If linked.Count = 0 Then Exit Sub

        For Each tpl In linked
            Dim lblId = tpl.Item1
            Dim kind = tpl.Item2
            Dim segIdx = tpl.Item3
            Dim dx = tpl.Item4
            Dim dy = tpl.Item5

            Dim mt = TryCast(tr.GetObject(lblId, OpenMode.ForWrite, False), MText)
            If mt Is Nothing Then Continue For

            ' 1) value
            Dim newLen As Double = ComputeLen(ent, kind, segIdx)
            mt.Contents = ReplaceFirstNumberKeepingPrefix(mt.Contents, newLen)

            ' 2) position = (segment midpoint) + (dx,dy)
            Dim mp As Point3d = SegmentMidPoint(ent, kind, segIdx)
            mt.Location = New Point3d(mp.X + dx, mp.Y + dy, 0)
            mt.Attachment = AttachmentPoint.MiddleCenter
        Next
    End Sub

    Friend Sub RefreshAllLenLabelsForDb(db As Database)
        If db Is Nothing Then Exit Sub
        Using tr = db.TransactionManager.StartTransaction()
            Dim nod = CType(tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead), DBDictionary)
            If Not nod.Contains(LENLINK_NOD) Then tr.Commit() : Exit Sub
            Dim links = CType(tr.GetObject(nod.GetAt(LENLINK_NOD), OpenMode.ForRead), DBDictionary)

            For Each kv In links
                Dim xr = CType(tr.GetObject(kv.Value, OpenMode.ForRead), Xrecord)
                Dim data = xr.Data.AsArray()
                Dim tgtHandle As String = CStr(data(0).Value)

                Dim h As New Handle(Convert.ToInt64(tgtHandle, 16))
                Dim targetId As ObjectId
                If Not db.TryGetObjectId(h, targetId) Then Continue For

                Dim ent = TryCast(tr.GetObject(targetId, OpenMode.ForRead, False), Entity)
                If ent Is Nothing Then Continue For

                UpdateLabelsForTarget(db, tr, ent) ' opens labels ForWrite only
            Next
            tr.Commit()
        End Using
    End Sub

    Friend Sub RefreshOnIdle(db As Database)
        If db Is Nothing Then Exit Sub
        Dim h As EventHandler = Nothing
        h = Sub(sender As Object, e As EventArgs)
                RemoveHandler Application.Idle, h
                RefreshAllLenLabelsForDb(db)
            End Sub
        AddHandler Application.Idle, h
    End Sub

    Friend Function SegmentMidPoint(ent As Entity, kind As LenKind, segIdx As Integer) As Point3d
        If TypeOf ent Is Line Then
            Dim ln = CType(ent, Line)
            Return New Point3d((ln.StartPoint.X + ln.EndPoint.X) / 2.0,
                           (ln.StartPoint.Y + ln.EndPoint.Y) / 2.0, 0)
        ElseIf TypeOf ent Is Polyline Then
            Dim pl = CType(ent, Polyline)

            ' choose segment:
            ' - if we were given a seg index (>=0) → use it
            ' - else (Total on 2-vertex polyline) → use [0]
            Dim seg As Integer = Math.Max(0, segIdx)

            Dim n As Integer = pl.NumberOfVertices
            Dim i0 As Integer = seg
            Dim i1 As Integer = (seg + 1) Mod n

            Dim p0 = pl.GetPoint2dAt(i0)
            Dim p1 = pl.GetPoint2dAt(i1)

            ' If arc (bulge != 0), midpoint ≈ chord midpoint (good enough for label anchor)
            Return New Point3d((p0.X + p1.X) / 2.0, (p0.Y + p1.Y) / 2.0, 0)
        End If
        Return Point3d.Origin
    End Function
End Module