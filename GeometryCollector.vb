Option Strict On
Option Infer On

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports System.Text.RegularExpressions

Module GeometryCollector

    ' === EXACT types from your main code ===
    Public Sub CollectGeometry(ids As ObjectId(),
                               ByVal columnEdges As List(Of Double),
                               ByVal beamLines As List(Of ReinforcementBarAutomation.HLine))

        If ids Is Nothing Then Exit Sub
        Dim tm = HostApplicationServices.WorkingDatabase.TransactionManager
        Using tr = tm.StartTransaction()
            For Each id In ids
                Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                If ent Is Nothing Then Continue For

                If TypeOf ent Is Line Then
                    Dim ln = CType(ent, Line)
                    Dim dx = Math.Abs(ln.Delta.X)
                    Dim dy = Math.Abs(ln.Delta.Y)

                    If dy > dx Then
                        columnEdges.Add(ln.StartPoint.X)
                        columnEdges.Add(ln.EndPoint.X)
                    Else
                        If String.Equals(ln.Layer, "S-BEAM", StringComparison.OrdinalIgnoreCase) Then
                            beamLines.Add(New ReinforcementBarAutomation.HLine(ln.StartPoint.Y, ln.StartPoint.X, ln.EndPoint.X))
                        End If
                    End If

                ElseIf TypeOf ent Is Polyline Then
                    Dim pl = CType(ent, Polyline)
                    Dim horiz As Double = 0.0, vert As Double = 0.0
                    For i = 0 To pl.NumberOfVertices - 2
                        Dim p0 = pl.GetPoint2dAt(i) : Dim p1 = pl.GetPoint2dAt(i + 1)
                        horiz += Math.Abs(p1.X - p0.X)
                        vert += Math.Abs(p1.Y - p0.Y)
                    Next

                    If String.Equals(pl.Layer, "S-BEAM", StringComparison.OrdinalIgnoreCase) AndAlso horiz >= vert Then
                        Dim minx = Double.PositiveInfinity, maxx = Double.NegativeInfinity
                        For i = 0 To pl.NumberOfVertices - 1
                            Dim p = pl.GetPoint2dAt(i)
                            If p.X < minx Then minx = p.X
                            If p.X > maxx Then maxx = p.X
                        Next
                        Dim yref = pl.GetPoint2dAt(0).Y
                        beamLines.Add(New ReinforcementBarAutomation.HLine(yref, minx, maxx))
                    ElseIf vert > horiz Then
                        For i = 0 To pl.NumberOfVertices - 1
                            columnEdges.Add(pl.GetPoint2dAt(i).X)
                        Next
                    End If
                End If
            Next
            tr.Commit()
        End Using
    End Sub

    Public Sub CollectLabels(ids As ObjectId(),
                             labels As List(Of ReinforcementBarAutomation.LabelPt))

        If ids Is Nothing Then Exit Sub
        Dim tm = HostApplicationServices.WorkingDatabase.TransactionManager
        Using tr = tm.StartTransaction()
            For Each id In ids
                Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                If ent Is Nothing Then Continue For

                If TypeOf ent Is DBText Then
                    Dim t = CType(ent, DBText)
                    Dim txt = (If(t.TextString, "")).Trim()
                    If txt.Length > 0 Then labels.Add(New ReinforcementBarAutomation.LabelPt With {.Text = txt, .Pt = t.Position})

                ElseIf TypeOf ent Is MText Then
                    Dim mt = CType(ent, MText)
                    Dim raw = (If(mt.Contents, "")).Trim()
                    If raw.Length > 0 Then
                        Dim clean = Regex.Replace(raw, "\{.*?\}", "").Trim()
                        labels.Add(New ReinforcementBarAutomation.LabelPt With {.Text = clean, .Pt = mt.Location})
                    End If
                End If
            Next
            tr.Commit()
        End Using
    End Sub

End Module