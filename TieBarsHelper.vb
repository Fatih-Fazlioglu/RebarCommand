Imports System.Text.RegularExpressions
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Module TieBarsHelper

    ' Parses "number1xnumber2/number3"   (e.g., "4x8/20")
    Public Function TryParseTieSpec(spec As String, ByRef qty As Integer, ByRef phi As Double, ByRef pitch As Double) As Boolean
        qty = 0 : phi = 0 : pitch = 0
        If String.IsNullOrWhiteSpace(spec) Then Return False
        Dim m = Regex.Match(spec.Trim(), "^\s*(\d+)\s*[xX]\s*(\d+(?:[.,]\d+)?)\s*/\s*(\d+(?:[.,]\d+)?)\s*$")
        If Not m.Success Then Return False
        qty = Integer.Parse(m.Groups(1).Value)
        phi = Double.Parse(m.Groups(2).Value.Replace(","c, "."c), Globalization.CultureInfo.InvariantCulture)
        pitch = Double.Parse(m.Groups(3).Value.Replace(","c, "."c), Globalization.CultureInfo.InvariantCulture)
        Return (qty > 0 AndAlso phi > 0 AndAlso pitch > 0)
    End Function

    ' Try to read the tie spec from the BeamExcelRow (tolerant to property names)
    Public Function GetTieSpecFromRow(row As Object) As String
        If row Is Nothing Then Return ""

        ' target key in normalized form
        Const TargetNorm As String = "duseyciroz"

        ' pattern like 3x10/10 (commas or dots allowed for decimals)
        Dim rx As New Regex("^\s*\d+\s*[xX]\s*\d+(?:[.,]\d+)?\s*/\s*\d+(?:[.,]\d+)?\s*$")

        ' 1) Try strongly-typed PROPERTIES (reflection)
        Dim t = row.GetType()
        For Each p In t.GetProperties()
            If Not p.CanRead Then Continue For
            Try
                Dim normName = NormalizeKey(p.Name)
                If normName = TargetNorm Then
                    Dim v = p.GetValue(row, Nothing)
                    Dim s = If(v Is Nothing, "", v.ToString())
                    If rx.IsMatch(s.Trim()) Then Return s.Trim()
                End If
            Catch
            End Try
        Next

        ' 2) Try FIELDS (sometimes codegen uses fields)
        For Each f In t.GetFields()
            Try
                Dim normName = NormalizeKey(f.Name)
                If normName = TargetNorm Then
                    Dim v = f.GetValue(row)
                    Dim s = If(v Is Nothing, "", v.ToString())
                    If rx.IsMatch(s.Trim()) Then Return s.Trim()
                End If
            Catch
            End Try
        Next

        ' 3) Try IDictionary(Of String, Object) or IDictionary
        Dim dictObj = TryCast(row, System.Collections.IDictionary)
        If dictObj IsNot Nothing Then
            For Each key In dictObj.Keys
                Try
                    Dim kn = NormalizeKey(key.ToString())
                    If kn = TargetNorm Then
                        Dim v = dictObj(key)
                        Dim s = If(v Is Nothing, "", v.ToString())
                        If rx.IsMatch(s.Trim()) Then Return s.Trim()
                    End If
                Catch
                End Try
            Next
        End If

        ' 4) Try DataRow (common when rows originate from Excel)
        Dim dr = TryCast(row, System.Data.DataRow)
        If dr IsNot Nothing Then
            For Each col As System.Data.DataColumn In dr.Table.Columns
                Try
                    Dim kn = NormalizeKey(col.ColumnName)
                    If kn = TargetNorm Then
                        Dim v = dr(col)
                        Dim s = If(v Is Nothing, "", v.ToString())
                        If rx.IsMatch(s.Trim()) Then Return s.Trim()
                    End If
                Catch
                End Try
            Next
        End If

        ' 5) Last resort: scan all string-like members for a value that looks like "3x10/10"
        '    (your original robust scan idea, but applied to properties+fields)
        For Each p In t.GetProperties()
            If Not p.CanRead Then Continue For
            Try
                Dim v = p.GetValue(row, Nothing)
                If v Is Nothing Then Continue For
                Dim s = v.ToString()
                If rx.IsMatch(s.Trim()) Then Return s.Trim()
            Catch
            End Try
        Next
        For Each f In t.GetFields()
            Try
                Dim v = f.GetValue(row)
                If v Is Nothing Then Continue For
                Dim s = v.ToString()
                If rx.IsMatch(s.Trim()) Then Return s.Trim()
            Catch
            End Try
        Next

        Return ""
    End Function

    ' NEW: Resolve tie spec — if not in row, prompt once per KNumber and cache
    Public Function ResolveTieSpec(ed As Editor, row As Object, kKey As String, cache As Dictionary(Of String, String)) As String
        Dim s As String = GetTieSpecFromRow(row)
        If Not String.IsNullOrWhiteSpace(s) Then Return s

        If cache IsNot Nothing AndAlso cache.TryGetValue(kKey, s) Then
            Return s
        End If

        Dim po As New PromptStringOptions(vbLf & $"Enter Düşey Çiroz for {kKey} (qty x phi / pitch, e.g. 4x8/20). Enter '-' to skip:")
        po.AllowSpaces = False
        Dim pr = ed.GetString(po)
        If pr.Status = PromptStatus.OK Then
            s = pr.StringResult.Trim()
            If s = "-" Then Return ""
            If cache IsNot Nothing Then
                If cache.ContainsKey(kKey) Then cache(kKey) = s Else cache.Add(kKey, s)
            End If
            Return s
        End If

        Return ""
    End Function

    Private Sub DrawFourTiesFromFace(
    btr As BlockTableRecord, tr As Transaction,
    faceX As Double, dir As Integer,
    baseY_Upper As Double, baseY_Lower As Double,
    pitch As Double, layerName As String)

        Const OFFSET_FROM_COLUMN As Double = 5.0 ' cm from face


        Dim xFirst As Double = Double.NaN        ' remember 1st tie X
        Dim yMid As Double = 0.5 * (baseY_Upper + baseY_Lower)

        For n As Integer = 0 To 3
            Dim xTie As Double = faceX + dir * (OFFSET_FROM_COLUMN + n * pitch)

            ' draw the vertical tie
            Dim ln As New Line(New Point3d(xTie, baseY_Upper, 0), New Point3d(xTie, baseY_Lower, 0)) With {
            .ColorIndex = 256,
            .Layer = layerName
        }
            btr.AppendEntity(ln) : tr.AddNewlyCreatedDBObject(ln, True)

            ' remember first tie X
            If n = 0 Then
                xFirst = xTie
            End If

            ' when second tie is drawn, dimension the spacing between first & second
            If n = 1 AndAlso Not Double.IsNaN(xFirst) Then
                Dim xLeft As Double = Math.Min(xFirst, xTie)
                Dim xRight As Double = Math.Max(xFirst, xTie)

                Dim p1 As New Point3d(xLeft, yMid, 0)
                Dim p2 As New Point3d(xRight, yMid, 0)

                ' horizontal measurement; +DIM_OFFSET places the dim 5 cm ABOVE
                DimensionUtils.AddOverlayDimLinear(
    btr, tr, p1, p2, 0.0, LayerManager.LYR_DIM,
    dimStyleName:="1.25_MERD_OLCU_ULK",
    textStyleName:="1.25_3.75_ULK",
    precision:=0, textHeight:=3.75, arrowSize:=0.15, textOffset:=0.1
)
            End If
        Next
    End Sub

    ' Main: draw tie bars for this assembly and beam group
    Public Sub DrawTieBarsForGroup(btr As BlockTableRecord, tr As Transaction,
                               uniqueEdges As List(Of Double),
                               c0 As Integer, c1 As Integer,
                               baseY_Upper As Double, baseY_Lower As Double,
                               getSpanSpec As Func(Of Integer, String),
                               layerName As String)
        If getSpanSpec Is Nothing Then Exit Sub
        Dim getPitch As Func(Of Integer, Double) =
        Function(spanIdx As Integer) As Double
            If spanIdx < c0 OrElse spanIdx > c1 - 1 Then Return 0.0
            Dim s As String = getSpanSpec(spanIdx)
            Dim q As Integer, ph As Double, p As Double
            If TryParseTieSpec(s, q, ph, p) Then Return p
            Return 0.0
        End Function

        For colIdx As Integer = c0 To c1
            Dim xL As Double = uniqueEdges(2 * colIdx)
            Dim xR As Double = uniqueEdges(2 * colIdx + 1)

            If colIdx = c0 Then
                Dim pitchR = getPitch(c0)
                If pitchR > 0 Then DrawFourTiesFromFace(btr, tr, xR, +1, baseY_Upper, baseY_Lower, pitchR, layerName)
            ElseIf colIdx = c1 Then
                Dim pitchL = getPitch(c1 - 1)
                If pitchL > 0 Then DrawFourTiesFromFace(btr, tr, xL, -1, baseY_Upper, baseY_Lower, pitchL, layerName)
            Else
                Dim pitchR = getPitch(colIdx)
                If pitchR > 0 Then DrawFourTiesFromFace(btr, tr, xR, +1, baseY_Upper, baseY_Lower, pitchR, layerName)
                Dim pitchL = getPitch(colIdx - 1)
                If pitchL > 0 Then DrawFourTiesFromFace(btr, tr, xL, -1, baseY_Upper, baseY_Lower, pitchL, layerName)
            End If
        Next
    End Sub




    ' === Section === 
    Private Function MakeCenteredPositions(xLeft As Double, xRight As Double,
                                       count As Integer, pitch As Double,
                                       bothUpperLowerEven As Boolean) As List(Of Double)
        Dim xs As New List(Of Double)
        If count <= 0 Then Return xs

        Dim mid As Double = 0.5 * (xLeft + xRight)
        Dim dir As Integer = -1  ' next step toggles L/R: -1 = left, +1 = right
        Dim stepIdx As Integer = 0

        ' If both groups even -> start at exact middle; otherwise start offset half a pitch to the left.
        If bothUpperLowerEven Then
            xs.Add(mid)
            stepIdx = 1
            dir = -1
        Else
            xs.Add(Math.Max(xLeft, Math.Min(xRight, mid - 0.5 * pitch)))
            stepIdx = 1
            dir = +1
        End If

        While xs.Count < count
            Dim off As Double = stepIdx * pitch
            Dim xTry As Double = mid + dir * off
            ' keep inside faces
            xTry = Math.Max(xLeft, Math.Min(xRight, xTry))
            xs.Add(xTry)
            ' toggle L/R and advance step
            dir = -dir
            If dir < 0 Then stepIdx += 1
        End While

        Return xs
    End Function

    ' Draw section tie bars (vertical lines) using a Düşey Çiroz spec (e.g. "6x8/20").
    ' number1 = total (incl. 2 stirrups) => draw (number1 - 2) ties.
    Public Sub DrawSectionTieBars(btr As BlockTableRecord, tr As Transaction,
                              xLeftFace As Double, xRightFace As Double,
                              yUpper As Double, yLower As Double,
                              tieSpec As String,
                              upperQty As Integer, lowerQty As Integer,
                              layerName As String)

        Dim qty As Integer, phi As Double, pitch As Double
        If Not TryParseTieSpec(tieSpec, qty, phi, pitch) Then Exit Sub
        Dim countToDraw As Integer = Math.Max(0, qty - 2)
        If countToDraw <= 0 OrElse pitch <= 0 Then Exit Sub

        Dim bothEven As Boolean = ((upperQty Mod 2) = 0 AndAlso (lowerQty Mod 2) = 0)

        Dim xs = MakeCenteredPositions(xLeftFace, xRightFace, countToDraw, pitch, bothEven)
        For Each x In xs
            Dim ln As New Line(New Point3d(x, yUpper, 0), New Point3d(x, yLower, 0)) With {
            .Layer = layerName,
            .ColorIndex = 256   ' ByLayer
        }
            btr.AppendEntity(ln) : tr.AddNewlyCreatedDBObject(ln, True)
        Next
    End Sub
    ' Add this helper inside TieBarsHelper
    Private Function NormalizeKey(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Dim t As String = s.Trim()

        ' Turkish diacritics → ASCII
        t = t.Replace("ş"c, "s"c).Replace("Ş"c, "S"c) _
             .Replace("ü"c, "u"c).Replace("Ü"c, "U"c) _
             .Replace("ç"c, "c"c).Replace("Ç"c, "C"c) _
             .Replace("ö"c, "o"c).Replace("Ö"c, "O"c) _
             .Replace("ğ"c, "g"c).Replace("Ğ"c, "G"c) _
             .Replace("ı"c, "i"c).Replace("İ"c, "I"c)

        ' remove spaces, underscores and dots
        t = t.Replace(" ", "").Replace("_", "").Replace(".", "")
        Return t.ToLowerInvariant()
    End Function

    Public Function TryGetKNumber(row As Object) As String
        If row Is Nothing Then Return ""
        Try
            Dim t = row.GetType()
            For Each p In t.GetProperties()
                If p.CanRead Then
                    Dim n = p.Name.ToLowerInvariant()
                    If n = "knumber" OrElse n = "k" OrElse n = "kno" OrElse n = "knum" Then
                        Dim v = p.GetValue(row, Nothing)
                        Return If(v Is Nothing, "", v.ToString())
                    End If
                End If
            Next
            For Each f In t.GetFields()
                Dim n = f.Name.ToLowerInvariant()
                If n = "knumber" OrElse n = "k" OrElse n = "kno" OrElse n = "knum" Then
                    Dim v = f.GetValue(row)
                    Return If(v Is Nothing, "", v.ToString())
                End If
            Next
        Catch
        End Try
        Return ""
    End Function


End Module
