
Option Strict On
Option Infer On

Imports System.Globalization
' ========= BCL =========
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.PlottingServices
' ========= AutoCAD =========
Imports Autodesk.AutoCAD.Runtime
' ========= ExcelDataReader (used only if available) =========
Imports ExcelDataReader
' Alias to avoid clash with WinForms.Application
Imports AcApp = Autodesk.AutoCAD.ApplicationServices.Application
Imports WinReg = Microsoft.Win32   ' <<— use alias to avoid ambiguity with AutoCAD.Runtime.Registry

' Ensure AutoCAD calls Loader.Initialize() on NETLOAD (adds AssemblyResolve + CodePages)
<Assembly: ExtensionApplication(GetType(Loader))>



' =========================================================================================
' LOADER — registers CodePagesEncodingProvider and resolves assemblies from the plug-in folder
' =========================================================================================
Public Class Loader
    Implements IExtensionApplication

    Public Sub Initialize() Implements IExtensionApplication.Initialize
        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf ResolveFromPluginFolder
        Try
            Dim t = Type.GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages")
            If t IsNot Nothing Then
                Dim inst = Activator.CreateInstance(t)
                Dim reg = GetType(System.Text.Encoding).GetMethod("RegisterProvider", BindingFlags.Public Or BindingFlags.Static)
                reg?.Invoke(Nothing, New Object() {inst})
            End If
        Catch
        End Try

    End Sub

    Public Sub Terminate() Implements IExtensionApplication.Terminate
    End Sub

    Private Function ResolveFromPluginFolder(sender As Object, args As ResolveEventArgs) As Assembly
        Try
            Dim baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            Dim dllName = New AssemblyName(args.Name).Name & ".dll"
            Dim candidate = Path.Combine(baseDir, dllName)
            If File.Exists(candidate) Then Return Assembly.LoadFrom(candidate)
        Catch
        End Try
        Return Nothing
    End Function
End Class


' =========================================================================================
' MAIN COMMAND
' =========================================================================================
Public Class ReinforcementBarAutomation
    Public Enum BeamSide
        LeftSide = 0
        RightSide = 1
    End Enum



    Private Class SectionJob
        Public TopY As Double
        Public BottomY As Double
        Public WCm As Double
        Public HCm As Double
        Public Info As BeamExcelRow
        Public SpanLeft As Double
        Public SpanRight As Double        ' for placement reference
        Public Side As BeamSide           ' left/right cut
        Public BeamName As String         ' e.g., "K12"
        Public TagX As Double
    End Class

    Private Shared _planBeamDimlineY As Double = Double.NaN

    ' ---------- Settings memory (HKCU) ----------
    Private NotInheritable Class Settings
        Private Sub New()
        End Sub
        Private Const ROOT As String = "Software\RebarCommand"

        Public Shared Function GetDouble(name As String, defaultVal As Double) As Double
            Try
                Using k = WinReg.Registry.CurrentUser.CreateSubKey(ROOT)
                    Dim s = TryCast(k.GetValue(name), String)
                    Dim v As Double
                    If s IsNot Nothing AndAlso Double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, v) Then
                        Return v
                    End If
                End Using
            Catch
            End Try
            Return defaultVal
        End Function

        Public Shared Sub SetDouble(name As String, value As Double)
            Try
                Using k = WinReg.Registry.CurrentUser.CreateSubKey(ROOT)
                    k.SetValue(name, value.ToString(CultureInfo.InvariantCulture), WinReg.RegistryValueKind.String)
                End Using
            Catch
            End Try
        End Sub

        Public Shared Function GetString(name As String, defaultVal As String) As String
            Try
                Using k = WinReg.Registry.CurrentUser.CreateSubKey(ROOT)
                    Dim s = TryCast(k.GetValue(name), String)
                    If Not String.IsNullOrEmpty(s) Then Return s
                End Using
            Catch
            End Try
            Return defaultVal
        End Function

        Public Shared Sub SetString(name As String, value As String)
            Try
                Using k = WinReg.Registry.CurrentUser.CreateSubKey(ROOT)
                    k.SetValue(name, value, WinReg.RegistryValueKind.String)
                End Using
            Catch
            End Try
        End Sub
    End Class

    ' Safety net: ensure AssemblyResolve/CodePages are active even if Loader didn’t run yet
    Shared Sub New()
        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf ResolveFromPluginFolder
        Try
            Dim t = Type.GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages")
            If t IsNot Nothing Then
                Dim inst = Activator.CreateInstance(t)
                Dim reg = GetType(System.Text.Encoding).GetMethod("RegisterProvider", BindingFlags.Public Or BindingFlags.Static)
                reg?.Invoke(Nothing, New Object() {inst})
            End If
        Catch
        End Try
    End Sub

    Private Shared Function ResolveFromPluginFolder(sender As Object, args As ResolveEventArgs) As Assembly
        Try
            Dim baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            Dim dllName = New AssemblyName(args.Name).Name & ".dll"
            Dim candidate = Path.Combine(baseDir, dllName)
            If File.Exists(candidate) Then Return Assembly.LoadFrom(candidate)
        Catch
        End Try
        Return Nothing
    End Function

#Region "Layers + helpers"



    Private Function GetConnectedBeamSpanX(beamLines As List(Of HLine),
                                       topY As Double,
                                       bottomY As Double,
                                       winL As Double,
                                       winR As Double,
                                       Optional tolY As Double = 0.001,
                                       Optional tolX As Double = 0.5) As Tuple(Of Double, Double)

        Dim spans As New List(Of Tuple(Of Double, Double))()

        For Each h As HLine In beamLines
            If Math.Abs(h.Y - topY) < tolY OrElse Math.Abs(h.Y - bottomY) < tolY Then
                Dim s As Double = Math.Min(h.X0, h.X1)
                Dim e As Double = Math.Max(h.X0, h.X1)
                If e > s Then spans.Add(Tuple.Create(s, e))
            End If
        Next

        If spans.Count = 0 Then
            ' Fallback: the window itself
            Return Tuple.Create(winL, winR)
        End If

        ' Sort by start, then merge overlaps/touching spans
        spans.Sort(Function(a, b) a.Item1.CompareTo(b.Item1))
        Dim merged As New List(Of Tuple(Of Double, Double))()
        Dim cur As Tuple(Of Double, Double) = spans(0)

        For i As Integer = 1 To spans.Count - 1
            Dim nxt = spans(i)
            If nxt.Item1 <= cur.Item2 + tolX Then
                cur = Tuple.Create(cur.Item1, Math.Max(cur.Item2, nxt.Item2))
            Else
                merged.Add(cur)
                cur = nxt
            End If
        Next
        merged.Add(cur)

        ' Choose the merged span(s) that intersect the current window
        Dim chosen As Tuple(Of Double, Double) = Nothing
        For Each m In merged
            Dim overlapL As Double = Math.Max(m.Item1, winL)
            Dim overlapR As Double = Math.Min(m.Item2, winR)
            If overlapR >= overlapL - tolX Then
                If chosen Is Nothing Then
                    chosen = m
                Else
                    chosen = Tuple.Create(Math.Min(chosen.Item1, m.Item1), Math.Max(chosen.Item2, m.Item2))
                End If
            End If
        Next

        If chosen Is Nothing Then
            ' If nothing intersects (rare), pick the closest merged span (first by sort)
            chosen = merged(0)
        End If

        Return chosen
    End Function


    ' =====================================================================
    ' PLAN TITLES — place "K####   W x H" above beam, just right of column
    ' =====================================================================
    Private Sub PlacePlanBeamNameSizeLabels(
    db As Database,
    ed As Editor,
    columnEdges As List(Of Double),
    beamLines As List(Of HLine),
    labels As List(Of LabelPt),
    Optional xGap As Double = 12.0,     ' right of column face (for text)
    Optional yGap As Double = 50.0,      ' above top beam line (for text)
    Optional textH As Double = 3.0      ' slightly smaller than Φ text (~3.33)
)
        If beamLines Is Nothing OrElse beamLines.Count = 0 Then Exit Sub

        Dim tm = db.TransactionManager
        Using tr = tm.StartTransaction()
            Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

            EnsureAllLayers(db, tr)

            ' Unique column edges (left,right,left,right,...)
            Dim uniqueEdges = UniqueSorted(New List(Of Double)(columnEdges), 0.000001)
            Dim numColumns As Integer = uniqueEdges.Count \ 2
            If numColumns < 1 Then
                tr.Commit()
                Return
            End If

            Dim lt = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)

            Dim asmRanges = SplitAssemblies(uniqueEdges)

            ' *** COMMON DIM BASELINE ***
            Dim globalBottom As Double = beamLines.Min(Function(h) h.Y)
            Dim dimOffset As Double = -2 * (Math.Abs(yGap) + textH + 18.0R)

            _planBeamDimlineY = globalBottom + dimOffset

            Dim beamSpanKeys As New HashSet(Of String)()

            ' -----------------------------
            ' MAIN LOOP: titles + BEAM dims
            '-------------------------------
            For Each r In asmRanges
                Dim c0 = r.Item1
                Dim c1 = r.Item2
                Dim asmLeft = uniqueEdges(2 * c0)
                Dim asmRight = uniqueEdges(2 * c1 + 1)

                ' all beam Y-levels intersecting this assembly
                Dim yValsAsm As New List(Of Double)()
                For Each h In beamLines
                    If h.X1 > asmLeft + 0.01 AndAlso h.X0 < asmRight - 0.01 Then
                        yValsAsm.Add(h.Y)
                    End If
                Next

                Dim uniqueYAsm = UniqueSorted(yValsAsm, 0.001)
                Dim groups = PairBeamLevels(uniqueYAsm)
                Dim nSpaces As Integer = (c1 - c0)
                If groups Is Nothing OrElse groups.Count = 0 OrElse nSpaces <= 0 Then Continue For

                For Each grp In groups
                    Dim topY As Double = grp.Item1
                    Dim bottomY As Double = grp.Item2

                    For iSpace As Integer = 0 To nSpaces - 1
                        Dim winL As Double = uniqueEdges(2 * (c0 + iSpace) + 1)
                        Dim winR As Double = uniqueEdges(2 * (c0 + iSpace + 1))

                        Dim span = GetConnectedBeamSpanX(beamLines, topY, bottomY, winL, winR)
                        If span.Item2 <= span.Item1 + 0.1R Then Continue For

                        ' ----------------- TITLE TEXT -----------------
                        Dim kName As String = FindClosestKNumber(labels, winL, winR, topY, bottomY)
                        If String.IsNullOrWhiteSpace(kName) Then kName = "K?"

                        Dim wCm As Double, hCm As Double
                        If Not FindClosestBeamSize(labels, winL, winR, topY, bottomY, wCm, hCm) Then
                            hCm = Math.Abs(topY - bottomY)
                            wCm = Math.Max(10.0R, Math.Round(hCm * 0.55R))
                        End If

                        ' Format: K101 (25/50)
                        Dim labelText As String =
                        $"{kName} ({CInt(Math.Round(wCm))}/{CInt(Math.Round(hCm))})"

                        Dim px As Double = winL + xGap
                        If px < span.Item1 + 2.0R Then px = span.Item1 + 2.0R
                        Dim py As Double = topY + yGap

                        ' Regex cleanup: catches "25x50", "25X50", and "25/50"
                        Dim rxTitle As New Regex("K\s*\d+.*?(x|×|X|/)\s*\d+", RegexOptions.IgnoreCase)

                        Dim xL As Double = winL - 5.0R
                        Dim xR As Double = winR + 5.0R
                        Dim yMin As Double = topY + 0.5R
                        Dim yMax As Double = topY + 600.0R

                        ' ---- Cleanup Old Labels ----
                        For Each id As ObjectId In btr
                            Dim ent = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                            If ent Is Nothing Then Continue For
                            If Not (TypeOf ent Is DBText OrElse TypeOf ent Is MText) Then Continue For

                            Dim p As Point3d
                            Dim raw As String = ""
                            If TypeOf ent Is MText Then
                                Dim mt0 = CType(ent, MText)
                                p = mt0.Location
                                raw = mt0.Contents.Replace("\P", " ")
                                raw = Regex.Replace(raw, "\{[^}]*\}", "")
                            Else
                                Dim t0 = CType(ent, DBText)
                                p = t0.Position
                                raw = t0.TextString
                            End If

                            If p.X < xL OrElse p.X > xR OrElse p.Y < yMin OrElse p.Y > yMax Then Continue For

                            Dim isYellow As Boolean = False
                            Dim c = ent.Color
                            If c IsNot Nothing AndAlso c.ColorMethod = ColorMethod.ByAci AndAlso c.ColorIndex = 2 Then
                                isYellow = True
                            ElseIf ent.ColorIndex = 256 OrElse (c IsNot Nothing AndAlso c.ColorMethod = ColorMethod.ByLayer) Then
                                If lt.Has(ent.Layer) Then
                                    Dim lr = CType(tr.GetObject(lt(ent.Layer), OpenMode.ForRead), LayerTableRecord)
                                    Dim lc = lr.Color
                                    If lc IsNot Nothing AndAlso lc.ColorMethod = ColorMethod.ByAci AndAlso lc.ColorIndex = 2 Then
                                        isYellow = True
                                    End If
                                End If
                            End If

                            Dim looksLike As Boolean =
                            (Not String.IsNullOrWhiteSpace(kName) AndAlso raw.IndexOf(kName, StringComparison.OrdinalIgnoreCase) >= 0) _
                            OrElse rxTitle.IsMatch(raw)

                            If isYellow AndAlso looksLike Then
                                ent.UpgradeOpen()
                                ent.Erase()
                            End If
                        Next

                        ' ---- CREATE NEW TEXT ----
                        Dim mt As New MText()
                        mt.Location = New Point3d(px, py, 0)
                        mt.Attachment = AttachmentPoint.BottomLeft
                        mt.TextHeight = textH
                        mt.Contents = labelText
                        mt.ColorIndex = 256

                        ' *** CALLING EXISTING ROMANS FUNCTION ***
                        mt.TextStyleId = GetRomansTextStyleId(tr, db)

                        AppendOnLayer(btr, tr, mt, LYR_BEAMTITLE, forceByLayer:=True)

                        ' ------------- BEAM DIM (INNER LENGTH) -------------
                        Dim innerL As Double = Math.Max(winL, span.Item1)
                        Dim innerR As Double = Math.Min(winR, span.Item2)
                        If innerR <= innerL + 0.1R Then Continue For

                        Dim key As String =
                        $"{Math.Round(innerL, 3).ToString(CultureInfo.InvariantCulture)}|" &
                        $"{Math.Round(innerR, 3).ToString(CultureInfo.InvariantCulture)}"
                        If beamSpanKeys.Contains(key) Then Continue For
                        beamSpanKeys.Add(key)

                        Dim pBeamL As New Point3d(innerL, globalBottom, 0)
                        Dim pBeamR As New Point3d(innerR, globalBottom, 0)
                        AddOverlayDimLinear(btr, tr, pBeamL, pBeamR, dimOffset + 30.0, LYR_DIM)
                    Next
                Next
            Next

            ' ----------------------------------------
            ' COLUMN DIMS
            ' ----------------------------------------
            For iCol As Integer = 0 To numColumns - 1
                Dim colL As Double = uniqueEdges(2 * iCol)
                Dim colR As Double = uniqueEdges(2 * iCol + 1)
                If colR <= colL + 0.1R Then Continue For

                Dim pColL As New Point3d(colL, globalBottom, 0)
                Dim pColR As New Point3d(colR, globalBottom, 0)
                AddOverlayDimLinear(btr, tr, pColL, pColR, dimOffset + 25.0, LYR_DIM)
            Next

            tr.Commit()
        End Using
    End Sub

    ' ==== END TITLE FIXER ======
#End Region

#Region "Simple types & helpers"


    ' ======== WEB DEBUG HELPERS ========
    Private Function ParseQtyFromRawGovde(raw As String) As Integer
        If String.IsNullOrWhiteSpace(raw) Then Return 0
        Dim s = raw.Replace("×", "x").Replace("Ø", "x").Replace("φ", "x").Replace("Φ", "x")
        s = Regex.Replace(s, "\s+", "")
        Dim m = Regex.Match(s, "^\s*(\d+)\s*x", RegexOptions.IgnoreCase)
        If m.Success Then Return Math.Max(0, Integer.Parse(m.Groups(1).Value, CultureInfo.InvariantCulture))
        Return 0
    End Function

    ' Use this everywhere instead of calling GetHalvedWebQty directly while debugging.
    ' It logs what info.WebQty is, what RawGovde implies, and the final halved value we will use.
    Private Function VerifyAndGetHalvedWebQty(ed As Editor, kKey As String, info As BeamExcelRow, context As String) As Integer
        If info Is Nothing Then
            LogWeb(ed, $"{context}: info is Nothing -> use 0")
            Return 0
        End If

        Dim qtyInfo As Integer = Math.Max(0, info.WebQty)
        Dim qtyParse As Integer = ParseQtyFromRawGovde(info.RawGovde)
        Dim halfInfo As Integer = qtyInfo \ 2
        Dim halfParse As Integer = qtyParse \ 2

        ' Choose info.WebQty as the source of truth; fall back only if it is 0 and parse>0
        Dim target As Integer = If(qtyInfo > 0, halfInfo, halfParse)

        ed.WriteMessage(vbLf & $"[WEBCHK] {context} | K='{kKey}' | Raw='{info.RawGovde}' | info.WebQty={qtyInfo} | parseQty={qtyParse} | half(info)={halfInfo} | half(parse)={halfParse} | target={target}")

        ' Optional: warn if the two paths disagree (helps catch any silent mismatch)
        If qtyInfo > 0 AndAlso qtyParse > 0 AndAlso halfInfo <> halfParse Then
            ed.WriteMessage(vbLf & $"[WEBCHK][WARN] {context} | info.WebQty and RawGovde parse disagree (using info.WebQty).")
        End If

        Return target
    End Function
    Private Function GetHalvedWebQty(info As BeamExcelRow) As Integer
        If info Is Nothing Then Return 0
        Dim q As Integer = Math.Max(0, info.WebQty)
        If q <= 0 AndAlso Not String.IsNullOrWhiteSpace(info.RawGovde) Then
            Dim qp = ParseQtyPhi(info.RawGovde, 0, ParsePhi(info.RawGovde)) ' qp.Item1 is qty
            q = Math.Max(0, qp.Item1)
        End If
        Return q \ 2
    End Function

    Private Const DEBUG_WEB As Boolean = True   ' <- flip to False to silence

    Private Sub LogWeb(ed As Editor, msg As String)
        If DEBUG_WEB Then ed.WriteMessage(vbLf & "[WEB] " & msg)
    End Sub

    Private Sub LogWebBeam(ed As Editor,
                       context As String,
                       kKey As String,
                       info As BeamExcelRow,
                       Optional extra As String = "")
        If Not DEBUG_WEB Then Return
        Dim rawGovde As String = If(info IsNot Nothing, info.RawGovde, "")
        Dim parsedQty As Integer = If(info IsNot Nothing, info.WebQty, -1)
        Dim halfQty As Integer = Math.Max(0, parsedQty) \ 2
        Dim phiWeb As Double = If(info IsNot Nothing, info.PhiWeb, 0.0R)
        ed.WriteMessage(
        vbLf & "[WEB] " & context &
        $" | K='{kKey}' | RawGovde='{rawGovde}' | ParsedQty={parsedQty} | HalfQty={halfQty} | PhiWeb={phiWeb} {extra}"
    )
    End Sub
    ' ======== END DEBUG HELPERS ========


    Public Structure HLine
        Public Y As Double
        Public X0 As Double
        Public X1 As Double
        Public Sub New(y As Double, x0 As Double, x1 As Double)
            Me.Y = y : Me.X0 = Math.Min(x0, x1) : Me.X1 = Math.Max(x0, x1)
        End Sub
    End Structure

    Public Class LabelPt
        Public Text As String = ""
        Public Pt As Point3d
    End Class

    Public Class BeamExcelRow
        Public PhiLower As Double = 1.6
        Public PhiMontaj As Double = 1.6
        Public WebQty As Integer = 0
        Public QtyLower As Integer = 2
        Public QtyMontaj As Integer = 2
        Public PhiWeb As Double = 1.2
        Public RawAlt As String = ""
        Public RawMontaj As String = ""
        Public RawGovde As String = ""
        Public TieBarsSpec As String = ""   ' e.g., "3x10/10"


        ' ---  Additional bars (ila.) ---
        Public QtyAddlLeft As Integer = 0
        Public PhiAddlLeft As Double = 0.0
        Public QtyAddlRight As Integer = 0
        Public PhiAddlRight As Double = 0.0
        Public RawAddlLeft As String = ""
        Public RawAddlRight As String = ""
    End Class


    ' ---- Φ yazıları (kr_dnyazi) ----
    Private Sub DrawPhiNote(btr As BlockTableRecord, tr As Transaction, x As Double, y As Double, txt As String, aci As Short)
        Const PhiTextHeight As Double = 15.0 / 3.0
        Const LabelYOffset As Double = 2.0
        Dim mt As New MText()
        mt.Location = New Point3d(x, y + LabelYOffset, 0)
        mt.Attachment = AttachmentPoint.MiddleCenter
        mt.TextHeight = PhiTextHeight
        mt.Contents = txt
        mt.ColorIndex = 256
        mt.TextStyleId = GetRomansTextStyleId(tr, btr.Database)
        AppendOnLayer(btr, tr, mt, LYR_PHI, forceByLayer:=True)
    End Sub

    ' ===== New helpers for QtyΦ labels =====
    Private Function ParseQtyPhi(s As String, defaultQty As Integer, defaultPhiCm As Double) As Tuple(Of Integer, Double)
        If String.IsNullOrWhiteSpace(s) Then Return Tuple.Create(defaultQty, defaultPhiCm)
        Dim t As String = s.Replace("×", "x").Replace("Ø", "x").Replace("φ", "x").Replace("Φ", "x")
        t = Regex.Replace(t, "\s+", "")
        Dim m = Regex.Match(t, "^(?:(\d+)\s*[x])?\s*(\d+)$", RegexOptions.IgnoreCase)
        If m.Success Then
            Dim q As Integer = If(m.Groups(1).Success, Math.Max(1, CInt(m.Groups(1).Value)), defaultQty)
            Dim dmm As Integer = CInt(m.Groups(2).Value)
            Return Tuple.Create(q, dmm / 10.0R)
        End If
        Dim md = Regex.Match(t, "(\d+)")
        If md.Success Then
            Dim dmm As Integer = CInt(md.Value)
            Return Tuple.Create(defaultQty, dmm / 10.0R)
        End If
        Return Tuple.Create(defaultQty, defaultPhiCm)
    End Function

    Private Function QtyPhiText(qty As Integer, phiCm As Double) As String
        Return qty.ToString(CultureInfo.InvariantCulture) & "Ø" &
               CInt(Math.Round(phiCm * 10.0R, MidpointRounding.AwayFromZero)).ToString()
    End Function

    Private Function QtyPhiTextAdditional(qty As Integer, phiCm As Double) As String
        Return QtyPhiText(qty, phiCm) & " ila."
    End Function

    Private Function ParsePhi(textVal As String) As Double
        If String.IsNullOrWhiteSpace(textVal) Then Return 1.6
        Dim s = textVal.Replace("Ø", "x").Replace("φ", "x").Replace("Φ", "x")
        s = Regex.Replace(s, "\s+", "")
        Dim m = Regex.Match(s, "(?:[0-9]+x)?([0-9]+)", RegexOptions.IgnoreCase)
        If m.Success Then
            Dim d As Integer
            If Integer.TryParse(m.Groups(1).Value, d) AndAlso d > 0 Then Return d / 10.0R
        End If
        Return 1.6
    End Function

    Private Function ParseWebQty(textVal As String) As Integer
        If String.IsNullOrWhiteSpace(textVal) Then Return 0
        Dim s As String = textVal
        s = s.Replace("×", "x")
        s = Regex.Replace(s, "\s+", "")
        Dim m As Match = Regex.Match(s, "(\d+)\s*[x]", RegexOptions.IgnoreCase)
        If m.Success Then Return Math.Max(0, Convert.ToInt32(m.Groups(1).Value, CultureInfo.InvariantCulture))
        Dim m2 As Match = Regex.Match(s, "^\s*(\d+)\s*[ØφΦ]", RegexOptions.IgnoreCase)
        If m2.Success Then Return Math.Max(0, Convert.ToInt32(m2.Groups(1).Value, CultureInfo.InvariantCulture))
        Return 0
    End Function

    Private Function TryParseBeamSizeFromText(ByVal txt As String, ByRef wCm As Double, ByRef hCm As Double) As Boolean
        wCm = 0 : hCm = 0
        If String.IsNullOrWhiteSpace(txt) Then Return False
        Dim s As String = txt
        s = s.Replace("×", "x").Replace("X", "x")
        s = Regex.Replace(s, "(\d+),(\d+)", "$1.$2")
        Dim m As Match = Regex.Match(s, "(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase)
        If Not m.Success Then Return False
        Dim a As Double, b As Double
        If Not Double.TryParse(m.Groups(1).Value, NumberStyles.Float, CultureInfo.InvariantCulture, a) Then Return False
        If Not Double.TryParse(m.Groups(2).Value, NumberStyles.Float, CultureInfo.InvariantCulture, b) Then Return False
        If a > 300 Then a /= 10.0R
        If b > 300 Then b /= 10.0R
        Dim widthCandidateCm As Double = Math.Min(a, b)
        Dim heightCandidateCm As Double = Math.Max(a, b)
        wCm = widthCandidateCm : hCm = heightCandidateCm
        Return True
    End Function

    Private Function FindClosestBeamSize(labels As List(Of LabelPt),
                                         xLeft As Double, xRight As Double,
                                         yTop As Double, yBot As Double,
                                         ByRef wCm As Double, ByRef hCm As Double) As Boolean
        wCm = 0 : hCm = 0
        If labels Is Nothing OrElse labels.Count = 0 Then Return False
        Dim bestDist As Double = Double.PositiveInfinity
        Dim found As Boolean = False
        Dim xMid As Double = 0.5 * (xLeft + xRight)
        Dim yMid As Double = 0.5 * (yTop + yBot)
        For Each l In labels
            Dim w As Double, h As Double
            If Not TryParseBeamSizeFromText(l.Text, w, h) Then Continue For
            Dim dx = Math.Max(0, Math.Max(xLeft - l.Pt.X, l.Pt.X - xRight))
            Dim dy = Math.Max(0, Math.Max(yBot - l.Pt.Y, l.Pt.Y - yTop))
            Dim d = Math.Sqrt(dx * dx + dy * dy) + 0.2 * Math.Abs(l.Pt.X - xMid) + 0.2 * Math.Abs(l.Pt.Y - yMid)
            If d < bestDist Then
                bestDist = d
                wCm = w : hCm = h
                found = True
            End If
        Next
        Return found
    End Function

#End Region

#Region "Geometry + labels"

    Private Function FindClosestKNumber(labels As List(Of LabelPt),
                                        xLeft As Double, xRight As Double,
                                        yTop As Double, yBot As Double) As String
        If labels.Count = 0 Then Return ""
        Dim rx As New Regex("^\s*K\s*\d+", RegexOptions.IgnoreCase)
        Dim best As String = "" : Dim bestDist As Double = Double.PositiveInfinity
        Dim xMid = 0.5 * (xLeft + xRight), yMid = 0.5 * (yTop + yBot)
        For Each l In labels
            Dim m = rx.Match(l.Text) : If Not m.Success Then Continue For
            Dim dx = Math.Max(0, Math.Max(xLeft - l.Pt.X, l.Pt.X - xRight))
            Dim dy = Math.Max(0, Math.Max(yBot - l.Pt.Y, l.Pt.Y - yTop))
            Dim d = Math.Sqrt(dx * dx + dy * dy) + 0.2 * Math.Abs(l.Pt.X - xMid) + 0.2 * Math.Abs(l.Pt.Y - yMid)
            If d < bestDist Then
                bestDist = d
                best = "K" & Regex.Match(m.Value, "\d+").Value
            End If
        Next
        Return best
    End Function

#End Region

#Region "Excel / CSV readers"

    Private Function TryLoadExcelMap(filePath As String,
                                 ByRef map As Dictionary(Of String, BeamExcelRow),
                                 ed As Editor) As Boolean
        map = New Dictionary(Of String, BeamExcelRow)(StringComparer.OrdinalIgnoreCase)
        If Not File.Exists(filePath) Then Return False

        Try
            Using fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using reader = ExcelReaderFactory.CreateReader(fs)

                    Do
                        Dim headerFound As Boolean = False

                        ' Required
                        Dim cK As Integer = -1
                        Dim cAlt As Integer = -1
                        Dim cMontaj As Integer = -1
                        Dim cGovde As Integer = -1
                        Dim cTie As Integer = -1

                        ' NEW: optional İlave columns
                        Dim cSolIlave As Integer = -1
                        Dim cSagIlave As Integer = -1

                        While reader.Read()
                            Dim cols = reader.FieldCount

                            ' -------- find headers once per sheet --------
                            If Not headerFound Then
                                Dim seen As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                                For i = 0 To cols - 1
                                    Dim v = reader.GetValue(i)
                                    If v Is Nothing OrElse v Is DBNull.Value Then Continue For
                                    Dim s = Convert.ToString(v, CultureInfo.InvariantCulture).Trim()
                                    If s.Length > 0 AndAlso Not seen.ContainsKey(s) Then seen(s) = i
                                Next

                                If seen.ContainsKey("KNumber") Then cK = seen("KNumber")
                                If seen.ContainsKey("Alt Donatı") Then cAlt = seen("Alt Donatı")
                                If seen.ContainsKey("Montaj") Then cMontaj = seen("Montaj")

                                If seen.ContainsKey("Gövde") Then cGovde = seen("Gövde")
                                If cGovde = -1 AndAlso seen.ContainsKey("Govde") Then cGovde = seen("Govde")

                                ' NEW: recognize Sol/Sağ İlave (with ASCII fallbacks)
                                If seen.ContainsKey("Sol İlave") Then cSolIlave = seen("Sol İlave")
                                If cSolIlave = -1 AndAlso seen.ContainsKey("Sol Ilave") Then cSolIlave = seen("Sol Ilave")

                                If seen.ContainsKey("Sağ İlave") Then cSagIlave = seen("Sağ İlave")
                                If cSagIlave = -1 AndAlso seen.ContainsKey("Sag Ilave") Then cSagIlave = seen("Sag Ilave")

                                If seen.ContainsKey("Düşey Çiroz") Then cTie = seen("Düşey Çiroz")
                                If cTie = -1 AndAlso seen.ContainsKey("Dusey Ciroz") Then cTie = seen("Dusey Ciroz")
                                If cTie = -1 AndAlso seen.ContainsKey("DUSEY CIROZ") Then cTie = seen("DUSEY CIROZ")
                                If cTie = -1 AndAlso seen.ContainsKey("DUSEY_CIROZ") Then cTie = seen("DUSEY_CIROZ")

                                headerFound = (cK >= 0 AndAlso cAlt >= 0 AndAlso cMontaj >= 0)
                                Continue While
                            End If
                            ' -------- after headers: read data rows --------

                            Dim kStr As String = SafeCellStr(reader, cK)
                            If String.IsNullOrWhiteSpace(kStr) Then Continue While

                            Dim altStr As String = SafeCellStr(reader, cAlt)
                            Dim monStr As String = SafeCellStr(reader, cMontaj)
                            Dim govStr As String = If(cGovde >= 0, SafeCellStr(reader, cGovde), "")

                            ' NEW: left/right additional strings (can be blank)
                            Dim solStr As String = If(cSolIlave >= 0, SafeCellStr(reader, cSolIlave), "")
                            Dim sagStr As String = If(cSagIlave >= 0, SafeCellStr(reader, cSagIlave), "")

                            Dim row As New BeamExcelRow()
                            row.RawAlt = altStr
                            row.RawMontaj = monStr
                            row.RawGovde = govStr
                            row.RawAddlLeft = solStr
                            row.RawAddlRight = sagStr
                            row.TieBarsSpec = If(cTie >= 0, SafeCellStr(reader, cTie).Trim(), "")

                            ' ---- parse main bars ----
                            Dim qpAlt = ParseQtyPhi(altStr, 2, ParsePhi(altStr))
                            row.QtyLower = qpAlt.Item1 : row.PhiLower = qpAlt.Item2

                            Dim qpMon = ParseQtyPhi(monStr, 2, ParsePhi(monStr))
                            row.QtyMontaj = qpMon.Item1 : row.PhiMontaj = qpMon.Item2

                            row.WebQty = ParseWebQty(govStr)
                            Dim qpGov = ParseQtyPhi(govStr, Math.Max(2, row.WebQty), ParsePhi(govStr))
                            row.PhiWeb = qpGov.Item2

                            ' ---- parse İlave (left/right); if blank they remain 0 ----
                            If Not String.IsNullOrWhiteSpace(solStr) Then
                                Dim qpL = ParseQtyPhi(solStr, 0, ParsePhi(solStr))
                                row.QtyAddlLeft = qpL.Item1
                                row.PhiAddlLeft = qpL.Item2
                            End If
                            If Not String.IsNullOrWhiteSpace(sagStr) Then
                                Dim qpR = ParseQtyPhi(sagStr, 0, ParsePhi(sagStr))
                                row.QtyAddlRight = qpR.Item1
                                row.PhiAddlRight = qpR.Item2
                            End If
                            If row.WebQty <= 0 Then row.WebQty = qpGov.Item1
                            LogWeb(AcApp.DocumentManager.MdiActiveDocument.Editor,
       $"LoadedRow: K will be normalized later; rawGovde='{row.RawGovde}' qty={row.WebQty} phi={row.PhiWeb}")


                            ' key normalization (allow "12" or "K12")
                            Dim key As String = kStr.Trim().ToUpperInvariant()
                            If Not key.StartsWith("K"c) Then key = "K" & key
                            map(key) = row
                        End While
                    Loop While reader.NextResult()
                End Using
            End Using

            Return (map.Count > 0)

        Catch ex As FileNotFoundException
            ed.WriteMessage(vbLf & "ExcelDataReader not available (" & ex.Message & ").")
            Return False
        Catch ex As Exception
            ed.WriteMessage(vbLf & "Excel read error: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function SafeCellStr(reader As IExcelDataReader, col As Integer) As String
        If col < 0 OrElse col >= reader.FieldCount Then Return ""
        Dim v = reader.GetValue(col)
        If v Is Nothing OrElse v Is DBNull.Value Then Return ""
        Return Convert.ToString(v, CultureInfo.InvariantCulture)
    End Function

    Private Function TryLoadCsvMap(filePath As String,
                               ByRef map As Dictionary(Of String, BeamExcelRow),
                               ed As Editor) As Boolean
        map = New Dictionary(Of String, BeamExcelRow)(StringComparer.OrdinalIgnoreCase)
        If Not File.Exists(filePath) Then Return False

        Try
            Dim enc As System.Text.Encoding = System.Text.Encoding.UTF8
            Using sr As New StreamReader(filePath, enc, detectEncodingFromByteOrderMarks:=True)

                Dim header As String = sr.ReadLine()
                If header Is Nothing Then Return False

                Dim sep As Char = DetectSeparator(header)
                Dim headers = header.Split(sep).Select(Function(h) h.Trim()).ToArray()

                ' required
                Dim idxK = Array.FindIndex(headers, Function(s) s.Equals("KNumber", StringComparison.OrdinalIgnoreCase))
                Dim idxAlt = Array.FindIndex(headers, Function(s) s.Equals("Alt Donatı", StringComparison.OrdinalIgnoreCase))
                Dim idxMon = Array.FindIndex(headers, Function(s) s.Equals("Montaj", StringComparison.OrdinalIgnoreCase))

                ' optional
                Dim idxGov = Array.FindIndex(headers, Function(s) s.Equals("Gövde", StringComparison.OrdinalIgnoreCase))
                If idxGov < 0 Then idxGov = Array.FindIndex(headers, Function(s) s.Equals("Govde", StringComparison.OrdinalIgnoreCase))

                ' NEW: optional İlave columns (with ASCII fallbacks)
                Dim idxSolIlave = Array.FindIndex(headers, Function(s) s.Equals("Sol İlave", StringComparison.OrdinalIgnoreCase))
                If idxSolIlave < 0 Then idxSolIlave = Array.FindIndex(headers, Function(s) s.Equals("Sol Ilave", StringComparison.OrdinalIgnoreCase))

                Dim idxSagIlave = Array.FindIndex(headers, Function(s) s.Equals("Sağ İlave", StringComparison.OrdinalIgnoreCase))
                If idxSagIlave < 0 Then idxSagIlave = Array.FindIndex(headers, Function(s) s.Equals("Sag Ilave", StringComparison.OrdinalIgnoreCase))
                Dim idxTie = Array.FindIndex(headers, Function(s) s.Equals("Düşey Çiroz", StringComparison.OrdinalIgnoreCase))
                If idxTie < 0 Then idxTie = Array.FindIndex(headers, Function(s) s.Equals("Dusey Ciroz", StringComparison.OrdinalIgnoreCase))
                If idxTie < 0 Then idxTie = Array.FindIndex(headers, Function(s) s.Equals("DUSEY CIROZ", StringComparison.OrdinalIgnoreCase))
                If idxTie < 0 Then idxTie = Array.FindIndex(headers, Function(s) s.Equals("DUSEY_CIROZ", StringComparison.OrdinalIgnoreCase))

                If idxK < 0 OrElse idxAlt < 0 OrElse idxMon < 0 Then
                    ed.WriteMessage(vbLf & "CSV missing required headers (KNumber, Alt Donatı, Montaj).")
                    Return False
                End If

                While Not sr.EndOfStream
                    Dim line = sr.ReadLine()
                    If String.IsNullOrWhiteSpace(line) Then Continue While

                    Dim cols = line.Split(sep)

                    Dim kStr As String = If(idxK < cols.Length, cols(idxK).Trim(), "")
                    If kStr.Length = 0 Then Continue While

                    Dim altStr As String = If(idxAlt < cols.Length, cols(idxAlt).Trim(), "")
                    Dim monStr As String = If(idxMon < cols.Length, cols(idxMon).Trim(), "")
                    Dim govStr As String = If(idxGov >= 0 AndAlso idxGov < cols.Length, cols(idxGov).Trim(), "")

                    ' NEW: İlave cells (can be empty)
                    Dim solStr As String = If(idxSolIlave >= 0 AndAlso idxSolIlave < cols.Length, cols(idxSolIlave).Trim(), "")
                    Dim sagStr As String = If(idxSagIlave >= 0 AndAlso idxSagIlave < cols.Length, cols(idxSagIlave).Trim(), "")

                    Dim row As New BeamExcelRow()
                    row.RawAlt = altStr
                    row.RawMontaj = monStr
                    row.RawGovde = govStr
                    row.RawAddlLeft = solStr
                    row.RawAddlRight = sagStr
                    row.TieBarsSpec = If(idxTie >= 0 AndAlso idxTie < cols.Length, cols(idxTie).Trim(), "")

                    ' ---- parse main bars ----
                    Dim qpAlt = ParseQtyPhi(altStr, 2, ParsePhi(altStr))
                    row.QtyLower = qpAlt.Item1 : row.PhiLower = qpAlt.Item2

                    Dim qpMon = ParseQtyPhi(monStr, 2, ParsePhi(monStr))
                    row.QtyMontaj = qpMon.Item1 : row.PhiMontaj = qpMon.Item2

                    row.WebQty = ParseWebQty(govStr)
                    Dim qpGov = ParseQtyPhi(govStr, Math.Max(2, row.WebQty), ParsePhi(govStr))
                    row.PhiWeb = qpGov.Item2

                    ' ---- parse İlave (left/right); if blank they remain 0 ----
                    If Not String.IsNullOrWhiteSpace(solStr) Then
                        Dim qpL = ParseQtyPhi(solStr, 0, ParsePhi(solStr))
                        row.QtyAddlLeft = qpL.Item1
                        row.PhiAddlLeft = qpL.Item2
                    End If
                    If Not String.IsNullOrWhiteSpace(sagStr) Then
                        Dim qpR = ParseQtyPhi(sagStr, 0, ParsePhi(sagStr))
                        row.QtyAddlRight = qpR.Item1
                        row.PhiAddlRight = qpR.Item2
                    End If

                    ' key normalization (allow "12" or "K12")
                    Dim key As String = kStr.ToUpperInvariant().Trim()
                    If Not key.StartsWith("K"c) Then key = "K" & key
                    map(key) = row
                End While
            End Using

            Return (map.Count > 0)

        Catch ex As Exception
            ed.WriteMessage(vbLf & "CSV read error: " & ex.Message)
            Return False
        End Try
    End Function


    Private Function DetectSeparator(header As String) As Char
        If header.Contains(vbTab) Then Return vbTab(0)
        If header.Contains(";") Then Return ";"c
        Return ","c
    End Function

#End Region

#Region "Command"

    <CommandMethod("DONATI_CIZ")>
    Public Sub DrawAllRebar()
        Dim doc = AcApp.DocumentManager.MdiActiveDocument
        Dim db = doc.Database
        Dim ed = doc.Editor
        Dim tieSpecCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)


        ' 1) Select columns + beams
        Dim selOpts As New PromptSelectionOptions() With {.MessageForAdding = vbLf & "Select columns (vertical) and beams (S-BEAM lines): "}
        Dim selRes = ed.GetSelection(selOpts)
        If selRes.Status <> PromptStatus.OK Then Return
        Dim ids = selRes.Value.GetObjectIds()

        ' 2) Ask for data file (remember last file)
        Dim ofo As New PromptOpenFileOptions(vbLf & "Select Excel/CSV with beam info (KNumber, Alt Donatı, Montaj, [Gövde]):")
        ofo.Filter = "Excel (*.xlsx;*.xls)|*.xlsx;*.xls|CSV (*.csv)|*.csv|All files (*.*)|*.*"
        Try
            Dim lastPath = Settings.GetString("LastDataPath", "")
            If Not String.IsNullOrEmpty(lastPath) Then
                Dim ip = Path.GetDirectoryName(lastPath)
                If Not String.IsNullOrEmpty(ip) AndAlso Directory.Exists(ip) Then
                    ofo.InitialDirectory = ip
                End If
                ofo.InitialFileName = Path.GetFileName(lastPath)
            End If
        Catch
        End Try
        Dim ofr = ed.GetFileNameForOpen(ofo)
        If ofr.Status <> PromptStatus.OK Then Return
        Dim dataPath As String = ofr.StringResult
        Settings.SetString("LastDataPath", dataPath)

        Dim ext = Path.GetExtension(dataPath).ToLowerInvariant()

        ' 3) Load map
        Dim excelMap As Dictionary(Of String, BeamExcelRow) = Nothing
        Dim ok As Boolean = False
        If ext = ".csv" Then
            ok = TryLoadCsvMap(dataPath, excelMap, ed)
        Else
            ok = TryLoadExcelMap(dataPath, excelMap, ed)
            If Not ok Then
                ed.WriteMessage(vbLf & "Falling back to CSV. Please choose a .csv file with KNumber, Alt Donatı, Montaj, [Gövde]...")
                Dim ofoCsv As New PromptOpenFileOptions(vbLf & "Select CSV:")
                ofoCsv.Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*"
                Dim ofrCsv = ed.GetFileNameForOpen(ofoCsv)
                If ofrCsv.Status = PromptStatus.OK Then
                    ok = TryLoadCsvMap(ofrCsv.StringResult, excelMap, ed)
                    If ofrCsv.Status = PromptStatus.OK Then Settings.SetString("LastDataPath", ofrCsv.StringResult)
                End If
            End If
        End If
        If Not ok Then
            ed.WriteMessage(vbLf & "No usable beam info loaded; using defaults.")
            excelMap = New Dictionary(Of String, BeamExcelRow)(StringComparer.OrdinalIgnoreCase)
        End If

        ' 4) UI parameters (remember last values)
        Dim dlg = New RebarParamsForm()
        Dim dr = AcApp.ShowModalDialog(dlg)
        If dr <> Windows.Forms.DialogResult.OK Then Return

        PosManager.Manager.Reset(1)
        PosManager.Manager.CarpanDonati = CInt(Math.Round(dlg.CarpanDonati))

        Dim overlayUser As Double = dlg.OverlayUpper
        Dim offsetUL As Double = dlg.OffsetUL
        Dim offsetWeb As Double = offsetUL
        Dim offsetLower As Double = offsetUL


        ' constants
        Const overlayHeight As Double = 3.0
        Const addlBelowUpper As Double = 2.0
        Const EPS As Double = 0.01
        Const segLen As Double = 1200.0
        Const DUP_SPACING As Double = 40.0
        Const DUP_OFFSET As Double = 180.0

        ' ---- dimension switches ----
        Dim showDimsOriginal As Boolean = False
        Dim showDimsDuplicate As Boolean = True
        '  Const SHOW_OVERLAYS_ON_ACTUAL As Boolean = False
        Const PASS_THROUGH_COLUMNS As Boolean = True


        ' 5) Collect geometry + labels
        Dim columnEdges As New List(Of Double)()
        Dim beamLines As New List(Of HLine)()          ' <-- your struct
        GeometryCollector.CollectGeometry(ids, columnEdges, beamLines)
        If beamLines.Count = 0 Then
            ed.WriteMessage(vbLf & "No S-BEAM lines detected.") : Return
        End If

        Dim labels As New List(Of LabelPt)()           ' <-- your struct
        GeometryCollector.CollectLabels(ids, labels)

        Dim uniqueEdges = UniqueSorted(columnEdges, 0.000001)
        Dim numColumns As Integer = uniqueEdges.Count \ 2

        Using tr = db.TransactionManager.StartTransaction()
            Dim btr = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
            EnsureAllLayers(db, tr)
            Dim colTxt As Short = 2



            ' ===================== BEAM-ONLY fallback =====================
            If numColumns < 2 Then
                Dim yVals As New List(Of Double)()
                For Each h In beamLines : yVals.Add(h.Y) : Next
                Dim uniqueY = UniqueSorted(yVals, 0.001)
                Dim groups = PairBeamLevels(uniqueY)
                If groups.Count = 0 Then tr.Commit() : Return

                For Each grp In groups
                    Dim secCounter As Integer = 0
                    Dim topY = grp.Item1
                    Dim bottomY = grp.Item2

                    Dim tolY = 0.001
                    Dim groupXs As New List(Of Double)()
                    For Each h In beamLines
                        If Math.Abs(h.Y - topY) < tolY OrElse Math.Abs(h.Y - bottomY) < tolY Then
                            groupXs.Add(h.X0) : groupXs.Add(h.X1)
                        End If
                    Next
                    If groupXs.Count = 0 Then Continue For
                    Dim minX = groupXs.Min()
                    Dim maxX = groupXs.Max()

                    Dim kKey = FindClosestKNumber(labels, minX, maxX, topY, bottomY)
                    Dim info As BeamExcelRow = Nothing
                    If Not excelMap.TryGetValue(kKey, info) Then info = New BeamExcelRow()

                    Dim wCm As Double, hCm As Double, beamHeightCm As Double
                    If FindClosestBeamSize(labels, minX, maxX, topY, bottomY, wCm, hCm) Then
                        beamHeightCm = hCm
                    Else
                        beamHeightCm = Math.Abs(topY - bottomY)
                    End If
                    Dim miterLen As Double = Round5Up(Math.Max(5.0R, beamHeightCm - 2.0R * offsetUL))

                    Dim extPhiLower = 50.0R * info.PhiLower
                    Dim overlayZoneUpper = overlayUser * info.PhiMontaj

                    Dim startX_UL = minX + offsetUL
                    Dim endX_UL = maxX - offsetUL
                    Dim startX_Web = minX + offsetWeb
                    Dim endX_Web = maxX - offsetWeb

                    ' UPPER (original)
                    Dim baseY_U = topY - offsetUL
                    If endX_UL - startX_UL > EPS Then
                        Dim plU As New Polyline()
                        ' Miters at 45° to bottom-left (225°)
                        plU.AddVertexAt(0, New Point2d(startX_UL - miterLen, baseY_U - miterLen), 0, 0, 0)
                        plU.AddVertexAt(1, New Point2d(startX_UL, baseY_U), 0, 0, 0)
                        plU.AddVertexAt(2, New Point2d(endX_UL, baseY_U), 0, 0, 0)
                        plU.AddVertexAt(3, New Point2d(endX_UL - miterLen, baseY_U - miterLen), 0, 0, 0)
                        plU.ColorIndex = 256
                        AppendOnLayer(btr, tr, plU, LYR_REBAR1, forceByLayer:=True)

                        Dim midXU = (startX_UL + endX_UL) / 2.0
                        DrawPhiNote(btr, tr, midXU, baseY_U, QtyPhiText(info.QtyMontaj, info.PhiMontaj), colTxt)
                    End If

                    ' LOWER (original)
                    Dim baseY_L = bottomY + offsetLower
                    If endX_UL - startX_UL > EPS Then
                        Dim plL As New Polyline()
                        ' Miters at 45° to top-left (135°) - symmetric with upper
                        plL.AddVertexAt(0, New Point2d(startX_UL - miterLen, baseY_L + miterLen), 0, 0, 0)
                        plL.AddVertexAt(1, New Point2d(startX_UL, baseY_L), 0, 0, 0)
                        plL.AddVertexAt(2, New Point2d(endX_UL, baseY_L), 0, 0, 0)
                        plL.AddVertexAt(3, New Point2d(endX_UL - miterLen, baseY_L + miterLen), 0, 0, 0)
                        plL.ColorIndex = 256
                        AppendOnLayer(btr, tr, plL, LYR_REBAR1, forceByLayer:=True)

                        Dim midXL = (startX_UL + endX_UL) / 2.0
                        DrawPhiNote(btr, tr, midXL, baseY_L, QtyPhiText(info.QtyLower, info.PhiLower), colTxt)
                    End If

                    ' ADDITIONAL (original; 10 cm below upper)
                    If endX_UL - startX_UL > EPS Then
                        Dim addlY = (topY - offsetUL) - addlBelowUpper

                        ' --- texts from Excel/CSV ---
                        Dim solQty As Integer = info.QtyAddlLeft
                        Dim solPhi As Double = info.PhiAddlLeft
                        Dim sagQty As Integer = info.QtyAddlRight
                        Dim sagPhi As Double = info.PhiAddlRight

                        Dim txtSol As String = If(solQty > 0 AndAlso solPhi > 0, QtyPhiTextAdditional(solQty, solPhi) & " (ila.)", "")
                        Dim txtSag As String = If(sagQty > 0 AndAlso sagPhi > 0, QtyPhiTextAdditional(sagQty, sagPhi) & " (ila.)", "")

                        If txtSol = "" AndAlso txtSag <> "" Then txtSol = txtSag
                        If txtSag = "" AndAlso txtSol <> "" Then txtSag = txtSol

                        Dim fallbackTxt As String = QtyPhiText(info.QtyMontaj, info.PhiMontaj) & " (ila.)"
                        If txtSol = "" Then txtSol = fallbackTxt
                        If txtSag = "" Then txtSag = fallbackTxt

                        Dim midX As Double = 0.5 * (startX_UL + endX_UL)

                        Dim plLeft As New Polyline()
                        ' Miter at 45° to bottom-left (225°)
                        plLeft.AddVertexAt(0, New Point2d(startX_UL - miterLen, addlY - miterLen), 0, 0, 0)
                        plLeft.AddVertexAt(1, New Point2d(startX_UL, addlY), 0, 0, 0)
                        plLeft.AddVertexAt(2, New Point2d(midX, addlY), 0, 0, 0)
                        plLeft.ColorIndex = 256
                        AppendOnLayer(btr, tr, plLeft, LYR_REBAR1, forceByLayer:=True)

                        Dim plRight As New Polyline()
                        plRight.AddVertexAt(0, New Point2d(midX, addlY), 0, 0, 0)
                        plRight.AddVertexAt(1, New Point2d(endX_UL, addlY), 0, 0, 0)
                        ' Miter at 45° to bottom-left (225°)
                        plRight.AddVertexAt(2, New Point2d(endX_UL - miterLen, addlY - miterLen), 0, 0, 0)
                        plRight.ColorIndex = 256
                        AppendOnLayer(btr, tr, plRight, LYR_REBAR1, forceByLayer:=True)

                        Dim midLeft As Double = 0.5 * (startX_UL + midX)
                        Dim midRight As Double = 0.5 * (midX + endX_UL)

                        DrawPhiNote(btr, tr, midLeft, addlY, txtSol, colTxt)
                        DrawPhiNote(btr, tr, midRight, addlY, txtSag, colTxt)
                    End If
                    ' WEB (original)
                    If endX_Web - startX_Web > EPS Then
                        Dim nWeb As Integer = VerifyAndGetHalvedWebQty(ed, kKey, info, "Normal-Orig-Target")
                        Dim topBound As Double = baseY_U
                        Dim botBound As Double = baseY_L
                        Dim usableH As Double = Math.Max(0.0R, topBound - botBound)
                        Dim webGap As Double = If(nWeb > 0, usableH / (nWeb + 1), 0.0R)
                        Dim lapWeb As Double = 50.0R * info.PhiMontaj
                        Dim stepLen As Double = Math.Max(5.0R, segLen - lapWeb)

                        For k As Integer = 1 To nWeb
                            ' UPDATED: Subtracting DUP_SPACING to shift this row down as requested
                            Dim yBase As Double = (topBound - k * webGap) - DUP_SPACING

                            Dim currentX As Double = startX_Web
                            Dim segIndex As Integer = 0
                            Do While currentX < (endX_Web - EPS)
                                Dim remaining = endX_Web - currentX
                                Dim length = If(remaining >= segLen, segLen, remaining)
                                Dim nextX = currentX + length

                                Dim ySeg As Double = yBase
                                Dim plW As New Polyline()
                                plW.AddVertexAt(0, New Point2d(currentX, ySeg), 0, 0, 0)
                                plW.AddVertexAt(1, New Point2d(nextX, ySeg), 0, 0, 0)
                                plW.ColorIndex = 256
                                AppendOnLayer(btr, tr, plW, LYR_REBAR1, forceByLayer:=True)

                                currentX += stepLen
                                segIndex += 1
                            Loop
                        Next
                        Dim midXW = (startX_Web + endX_Web) / 2.0
                        Dim midYW = (topBound + botBound) / 2.0
                        DrawPhiNote(btr, tr, midXW, midYW, QtyPhiText(2 * GetHalvedWebQty(info), info.PhiWeb), colTxt)
                    End If

                    ' ===== DUPLICATE SET =====
                    Dim qtyDup As Integer = If(info.WebQty > 0, info.WebQty, 2)
                    Dim nWebDup As Integer = Math.Max(1, qtyDup \ 2)
                    Dim dupAddlY As Double = (topY - offsetUL) - DUP_OFFSET
                    Dim dupUpperY As Double = dupAddlY - DUP_SPACING
                    ' Draw only ONE web line, so use 1 instead of nWebDup for Y position calculation
                    Dim dupLowerY As Double = dupUpperY - 1 * DUP_SPACING - DUP_SPACING

                    ' Upper (duplicate)
                    If endX_UL - startX_UL > EPS Then
                        Dim plU2 As New Polyline()
                        ' Miters at 45° to bottom-left (225°)
                        plU2.AddVertexAt(0, New Point2d(startX_UL - miterLen, dupUpperY - miterLen), 0, 0, 0)
                        plU2.AddVertexAt(1, New Point2d(startX_UL, dupUpperY), 0, 0, 0)
                        plU2.AddVertexAt(2, New Point2d(endX_UL, dupUpperY), 0, 0, 0)
                        plU2.AddVertexAt(3, New Point2d(endX_UL - miterLen, dupUpperY - miterLen), 0, 0, 0)
                        plU2.ColorIndex = 256
                        AppendOnLayer(btr, tr, plU2, LYR_REBAR2, forceByLayer:=True)
                        Dim span = endX_UL - startX_UL

                    End If

                    ' Additional (duplicate)
                    If endX_UL - startX_UL > EPS Then
                        Dim plA2 As New Polyline()
                        ' Miters at 45° to bottom-left (225°)
                        plA2.AddVertexAt(0, New Point2d(startX_UL - miterLen, dupAddlY - miterLen), 0, 0, 0)
                        plA2.AddVertexAt(1, New Point2d(startX_UL, dupAddlY), 0, 0, 0)
                        plA2.AddVertexAt(2, New Point2d(endX_UL, dupAddlY), 0, 0, 0)
                        plA2.AddVertexAt(3, New Point2d(endX_UL - miterLen, dupAddlY - miterLen), 0, 0, 0)
                        plA2.ColorIndex = 256
                        AppendOnLayer(btr, tr, plA2, LYR_REBAR2, forceByLayer:=True)
                    End If

                    ' Webs (duplicate) - draw only ONE line
                    If endX_Web - startX_Web > EPS AndAlso nWebDup > 0 Then
                        Dim lapWeb As Double = 50.0R * info.PhiMontaj
                        Dim stepLen As Double = Math.Max(5.0R, segLen - lapWeb)

                        ' Draw only one web rebar line (k=1)
                        Dim yDup As Double = (dupUpperY - 1 * DUP_SPACING)
                        Dim currentX As Double = startX_Web
                        Dim segIndex As Integer = 0

                        Do While currentX < (endX_Web - EPS)
                            Dim remaining = endX_Web - currentX
                            Dim length = If(remaining >= segLen, segLen, remaining)
                            Dim nextX = currentX + length

                            Dim ySeg As Double = If((segIndex And 1) = 0, yDup, yDup - overlayHeight)

                            ' draw the segment
                            Dim plW2 As New Polyline()
                            plW2.AddVertexAt(0, New Point2d(currentX, ySeg), 0, 0, 0)
                            plW2.AddVertexAt(1, New Point2d(nextX, ySeg), 0, 0, 0)
                            plW2.ColorIndex = 256
                            AppendOnLayer(btr, tr, plW2, LYR_REBAR2, forceByLayer:=True)

                            currentX += stepLen
                            segIndex += 1
                        Loop
                    End If

                    ' Lower (duplicate)
                    If endX_UL - startX_UL > EPS Then
                        Dim plL2 As New Polyline()
                        ' Miters at 45° to top-left (135°) - symmetric with upper
                        plL2.AddVertexAt(0, New Point2d(startX_UL - miterLen, dupLowerY + miterLen), 0, 0, 0)
                        plL2.AddVertexAt(1, New Point2d(startX_UL, dupLowerY), 0, 0, 0)
                        plL2.AddVertexAt(2, New Point2d(endX_UL, dupLowerY), 0, 0, 0)
                        plL2.AddVertexAt(3, New Point2d(endX_UL - miterLen, dupLowerY + miterLen), 0, 0, 0)
                        plL2.ColorIndex = 256
                        AppendOnLayer(btr, tr, plL2, LYR_REBAR2, forceByLayer:=True)
                        DrawDimIf(btr, tr, True, endX_UL - extPhiLower, endX_UL, dupLowerY, FormatLen(extPhiLower), 3, placeAbove:=False)
                        DrawDimIf(btr, tr, True, startX_UL, startX_UL + extPhiLower, dupLowerY, FormatLen(extPhiLower), 3, placeAbove:=False)
                    End If
                Next

                tr.Commit() : Return
            End If

            ' ===================== NORMAL path (with columns) =====================
            Dim asmRanges = SplitAssemblies(uniqueEdges)
            If asmRanges.Count = 0 Then tr.Commit() : Return

            For Each r In asmRanges
                Dim c0 As Integer = r.Item1
                Dim c1 As Integer = r.Item2
                Dim asmLeft As Double = uniqueEdges(2 * c0)
                Dim asmRight As Double = uniqueEdges(2 * c1 + 1)

                Dim hasLeftInside As Boolean = (2 * c0 + 1) < uniqueEdges.Count
                Dim leftInsideFace As Double = If(hasLeftInside, uniqueEdges(2 * c0 + 1), asmLeft)

                Dim hasRightInside As Boolean = (2 * c1) < uniqueEdges.Count
                Dim rightInsideFace As Double = If(hasRightInside, uniqueEdges(2 * c1), asmRight)

                Dim yValsAsm As New List(Of Double)()
                For Each h As HLine In beamLines
                    If h.X1 > asmLeft + EPS AndAlso h.X0 < asmRight - EPS Then yValsAsm.Add(h.Y)
                Next
                Dim uniqueYAsm As List(Of Double) = UniqueSorted(yValsAsm, 0.001)
                Dim groups As List(Of Tuple(Of Double, Double)) = PairBeamLevels(uniqueYAsm)
                If groups.Count = 0 Then Continue For

                ' Midpoints between columns in this assembly window (used by "original, segmented" mode)
                Dim mids As New List(Of Double)()
                For i As Integer = c0 To c1 - 1
                    Dim rightFace As Double = uniqueEdges(2 * i + 1)
                    Dim leftNext As Double = uniqueEdges(2 * (i + 1))
                    mids.Add((rightFace + leftNext) / 2.0R)
                Next
                Dim nMid As Integer = mids.Count
                Dim startX_U As Double = asmLeft + offsetUL
                Dim endX_U As Double = asmRight - offsetUL

                For Each grp In groups
                    Dim secCounter As Integer = 0
                    Dim topY As Double = grp.Item1
                    Dim bottomY As Double = grp.Item2

                    ' union range for pass-through
                    Dim ux As Tuple(Of Double, Double) = GetConnectedBeamSpanX(beamLines, topY, bottomY, asmLeft, asmRight)
                    Dim uMin As Double = ux.Item1
                    Dim uMax As Double = ux.Item2

                    Dim kKey As String = FindClosestKNumber(labels, asmLeft, asmRight, topY, bottomY)
                    Dim info As BeamExcelRow = Nothing
                    If Not excelMap.TryGetValue(kKey, info) Then info = New BeamExcelRow()

                    Dim wCm As Double, hCm As Double, beamHeightCm As Double
                    If FindClosestBeamSize(labels, asmLeft, asmRight, topY, bottomY, wCm, hCm) Then
                        beamHeightCm = hCm
                    Else
                        beamHeightCm = Math.Abs(topY - bottomY)
                    End If
                    Dim miterLen As Double = Round5Up(Math.Max(5.0R, beamHeightCm - 2.0R * offsetUL))



                    Dim extPhiLower As Double = 50.0R * info.PhiLower
                    Dim overlayZoneUpper As Double = overlayUser * info.PhiMontaj

                    ' ----- UPPER (original, segmented)
                    Dim baseY_Upper As Double = topY - offsetUL

                    If PASS_THROUGH_COLUMNS Then
                        Dim sU As Double = asmLeft + offsetUL
                        Dim eU As Double = Math.Max(asmRight, uMax) - offsetUL
                        If eU - sU > EPS Then
                            Dim pl As New Polyline()
                            ' Miters at 45° to bottom-left (225°)
                            pl.AddVertexAt(0, New Point2d(sU - miterLen, baseY_Upper - miterLen), 0, 0, 0)
                            pl.AddVertexAt(1, New Point2d(sU, baseY_Upper), 0, 0, 0)
                            pl.AddVertexAt(2, New Point2d(eU, baseY_Upper), 0, 0, 0)
                            pl.AddVertexAt(3, New Point2d(eU - miterLen, baseY_Upper - miterLen), 0, 0, 0)
                            pl.ColorIndex = 256
                            AppendOnLayer(btr, tr, pl, LYR_REBAR1, forceByLayer:=True)
                            DrawPhiNote(btr, tr, 0.5R * (sU + eU), baseY_Upper, QtyPhiText(info.QtyMontaj, info.PhiMontaj), colTxt)
                        End If
                    Else
                        ' ----- UPPER (original, segmented)
                        For j As Integer = 0 To nMid
                            Dim rawX0 As Double, rawX1 As Double
                            If j = 0 Then
                                rawX0 = startX_U : rawX1 = If(nMid > 0, mids(0), endX_U)
                            ElseIf j = nMid Then
                                rawX0 = If(nMid > 0, mids(nMid - 1), startX_U) : rawX1 = endX_U
                            Else
                                rawX0 = mids(j - 1) : rawX1 = mids(j)
                            End If

                            Dim span As Double = Math.Max(0.0R, rawX1 - rawX0)
                            If span <= EPS Then Continue For

                            Dim x0 As Double = rawX0
                            Dim x1 As Double = rawX1
                            If x1 - x0 <= EPS Then Continue For

                            Dim yCoord As Double = baseY_Upper
                            Dim roundedLen As Double = Round5Up(x1 - x0)
                            If roundedLen <= 0 Then Continue For

                            Dim newStart As Double, newEnd As Double
                            If j = 0 Then
                                newStart = x0 : newEnd = Math.Min(x1, x0 + roundedLen)
                                Dim pl As New Polyline()
                                ' Miter at 45° to bottom-left (225°)
                                pl.AddVertexAt(0, New Point2d(newStart - miterLen, baseY_Upper - miterLen), 0, 0, 0)
                                pl.AddVertexAt(1, New Point2d(newStart, baseY_Upper), 0, 0, 0)
                                pl.AddVertexAt(2, New Point2d(newEnd, baseY_Upper), 0, 0, 0)
                                pl.ColorIndex = 256
                                AppendOnLayer(btr, tr, pl, LYR_REBAR1, forceByLayer:=True)
                            ElseIf j = nMid Then
                                newEnd = x1 : newStart = Math.Max(x0, x1 - roundedLen)
                                Dim pl As New Polyline()
                                pl.AddVertexAt(0, New Point2d(newStart, yCoord), 0, 0, 0)
                                pl.AddVertexAt(1, New Point2d(newEnd, yCoord), 0, 0, 0)
                                ' Miter at 45° to bottom-left (225°)
                                pl.AddVertexAt(2, New Point2d(newEnd - miterLen, yCoord - miterLen), 0, 0, 0)
                                pl.ColorIndex = 256
                                AppendOnLayer(btr, tr, pl, LYR_REBAR1, forceByLayer:=True)
                            Else
                                Dim midX As Double = (x0 + x1) / 2.0R
                                newStart = Math.Max(x0, midX - roundedLen / 2.0R)
                                newEnd = Math.Min(x1, midX + roundedLen / 2.0R)
                                Dim ln As New Line(New Point3d(newStart, yCoord, 0), New Point3d(newEnd, yCoord, 0))
                                ln.ColorIndex = 256
                                AppendOnLayer(btr, tr, ln, LYR_REBAR1, forceByLayer:=True)
                            End If

                            Dim midSeg As Double = (newStart + newEnd) / 2.0R
                            DrawPhiNote(btr, tr, midSeg, yCoord, QtyPhiText(info.QtyMontaj, info.PhiMontaj), colTxt)
                        Next
                    End If

                    ' ===== LOWER
                    Dim nSpaces As Integer = (c1 - c0)
                    Dim baseY_Lower As Double = bottomY + offsetLower
                    Dim startXFirstLower As Double = asmLeft + offsetLower
                    Dim endXLastLower As Double = asmRight - offsetLower

                    ' Build a per-span map: spanIdx (absolute, c0..c1-1) -> BeamExcelRow
                    Dim spansByIndex As New Dictionary(Of Integer, ReinforcementBarAutomation.BeamExcelRow)()

                    For iSpace As Integer = 0 To nSpaces - 1
                        Dim leftColIdx As Integer = c0 + iSpace
                        Dim rightColIdx As Integer = c0 + iSpace + 1

                        ' faces that bound this span
                        Dim spanLeft As Double = uniqueEdges(2 * leftColIdx + 1)   ' right face of left column
                        Dim spanRight As Double = uniqueEdges(2 * rightColIdx)     ' left face of right column

                        ' resolve K and row for THIS span (same logic you already use elsewhere)
                        Dim kSpan As String = FindClosestKNumber(labels, spanLeft, spanRight, topY, bottomY)
                        Dim infoSpan As ReinforcementBarAutomation.BeamExcelRow = Nothing
                        If Not excelMap.TryGetValue(kSpan, infoSpan) Then
                            ' optional fallback: search assembly-wide
                            Dim kAsm As String = FindClosestKNumber(labels, asmLeft, asmRight, topY, bottomY)
                            If Not excelMap.TryGetValue(kAsm, infoSpan) Then
                                infoSpan = New ReinforcementBarAutomation.BeamExcelRow()
                            End If
                        End If

                        spansByIndex(leftColIdx) = infoSpan   ' store by absolute span index
                    Next

                    ' Now wire per-span spec to the helper (no late binding; Option Strict-safe)
                    TieBarsHelper.DrawTieBarsForGroup(
    btr, tr, uniqueEdges,
    c0, c1,
    baseY_Upper, baseY_Lower,
    Function(spanIdx As Integer) As String

        Dim rowObj As ReinforcementBarAutomation.BeamExcelRow = Nothing
        If spansByIndex.TryGetValue(spanIdx, rowObj) Then
            ' 1) prefer the spec on the row
            Dim s As String = TieBarsHelper.GetTieSpecFromRow(rowObj)
            If Not String.IsNullOrWhiteSpace(s) Then Return s

            ' 2) fallback: prompt+cache using that row’s KNumber
            Dim kSpan As String = TieBarsHelper.TryGetKNumber(rowObj)
            If String.IsNullOrWhiteSpace(kSpan) Then kSpan = $"{kKey}_SPAN_{spanIdx}"
            Return TieBarsHelper.ResolveTieSpec(ed, rowObj, kSpan, tieSpecCache)
        End If

        ' no row → last resort prompt per span key
        Return TieBarsHelper.ResolveTieSpec(ed, Nothing, $"{kKey}_SPAN_{spanIdx}", tieSpecCache)
    End Function,
    LYR_REBAR1)

                    If PASS_THROUGH_COLUMNS Then
                        Dim sL As Double = asmLeft + offsetLower
                        Dim eL As Double = Math.Max(asmRight, uMax) - offsetLower
                        If eL - sL > EPS Then
                            Dim pl As New Polyline()
                            ' Miters at 45° to top-left (135°) - symmetric with upper
                            pl.AddVertexAt(0, New Point2d(sL - miterLen, baseY_Lower + miterLen), 0, 0, 0)
                            pl.AddVertexAt(1, New Point2d(sL, baseY_Lower), 0, 0, 0)
                            pl.AddVertexAt(2, New Point2d(eL, baseY_Lower), 0, 0, 0)
                            pl.AddVertexAt(3, New Point2d(eL - miterLen, baseY_Lower + miterLen), 0, 0, 0)
                            pl.ColorIndex = 256
                            AppendOnLayer(btr, tr, pl, LYR_REBAR1, forceByLayer:=True)

                            ' optional end-lap dims
                            Dim extL As Double = 50.0R * info.PhiLower
                            '  DrawDimIf(btr, tr, True, eL - extL, eL, baseY_Lower, FormatLen(extL), 3, placeAbove:=False)
                            ' DrawDimIf(btr, tr, True, sL, sL + extL, baseY_Lower, FormatLen(extL), 3, placeAbove:=False)

                            DrawPhiNote(btr, tr, 0.5R * (sL + eL), baseY_Lower, QtyPhiText(info.QtyLower, info.PhiLower), colTxt)
                        End If
                    Else


                        ' ===== LOWER (original)
                        For iSpace As Integer = 0 To nSpaces - 1
                            Dim leftColIdx As Integer = c0 + iSpace
                            Dim rightColIdx As Integer = c0 + iSpace + 1

                            Dim leftRightFace As Double = uniqueEdges(2 * leftColIdx + 1)
                            Dim rightLeftFace As Double = uniqueEdges(2 * rightColIdx)

                            Dim x0L As Double, x1L As Double
                            If iSpace = 0 Then
                                x0L = startXFirstLower
                                x1L = rightLeftFace + (50.0R * info.PhiLower)
                            ElseIf iSpace = nSpaces - 1 Then
                                x0L = leftRightFace - (50.0R * info.PhiLower)
                                x1L = endXLastLower
                            Else
                                x0L = leftRightFace - (50.0R * info.PhiLower)
                                x1L = rightLeftFace + (50.0R * info.PhiLower)
                            End If

                            If x1L - x0L <= EPS Then Continue For

                            If iSpace = 0 Then
                                Dim pl As New Polyline()
                                ' Miter at 45° to top-left (135°)
                                pl.AddVertexAt(0, New Point2d(x0L - miterLen, baseY_Lower + miterLen), 0, 0, 0)
                                pl.AddVertexAt(1, New Point2d(x0L, baseY_Lower), 0, 0, 0)
                                pl.AddVertexAt(2, New Point2d(x1L, baseY_Lower), 0, 0, 0)
                                pl.ColorIndex = 256
                                AppendOnLayer(btr, tr, pl, LYR_REBAR1, forceByLayer:=True)
                            ElseIf iSpace = nSpaces - 1 Then
                                Dim pl As New Polyline()
                                pl.AddVertexAt(0, New Point2d(x0L, baseY_Lower), 0, 0, 0)
                                pl.AddVertexAt(1, New Point2d(x1L, baseY_Lower), 0, 0, 0)
                                ' Miter at 45° to top-left (135°)
                                pl.AddVertexAt(2, New Point2d(x1L - miterLen, baseY_Lower + miterLen), 0, 0, 0)
                                pl.ColorIndex = 256
                                AppendOnLayer(btr, tr, pl, LYR_REBAR1, forceByLayer:=True)
                            Else
                                Dim ln As New Line(New Point3d(x0L, baseY_Lower, 0), New Point3d(x1L, baseY_Lower, 0))
                                ln.ColorIndex = 256
                                AppendOnLayer(btr, tr, ln, LYR_REBAR1, forceByLayer:=True)
                            End If

                            Dim midXL As Double = (x0L + x1L) / 2.0R
                            DrawPhiNote(btr, tr, midXL, baseY_Lower, QtyPhiText(info.QtyLower, info.PhiLower), colTxt)
                        Next
                    End If

                    ' ===== ADDITIONAL
                    Dim addlY As Double = (topY - offsetUL) - addlBelowUpper

                    Dim bayHalves As New List(Of Tuple(Of Double, Double, String))()
                    For j As Integer = c0 To c1 - 1
                        Dim leftRightFace As Double = uniqueEdges(2 * j + 1)
                        Dim rightLeftFace As Double = uniqueEdges(2 * (j + 1))
                        Dim bayInsideL As Double = leftRightFace + offsetUL
                        Dim bayInsideR As Double = rightLeftFace - offsetUL
                        If bayInsideR - bayInsideL <= EPS Then Continue For

                        Dim kKeyBay As String = FindClosestKNumber(labels, leftRightFace, rightLeftFace, topY, bottomY)
                        Dim infoBay As BeamExcelRow = Nothing
                        If Not excelMap.TryGetValue(kKeyBay, infoBay) Then infoBay = info

                        Dim solQty As Integer = infoBay.QtyAddlLeft
                        Dim solPhi As Double = infoBay.PhiAddlLeft
                        Dim sagQty As Integer = infoBay.QtyAddlRight
                        Dim sagPhi As Double = infoBay.PhiAddlRight

                        Dim txtSol As String = If(solQty > 0 AndAlso solPhi > 0.0R, QtyPhiText(solQty, solPhi) & " (ila.)", "")
                        Dim txtSag As String = If(sagQty > 0 AndAlso sagPhi > 0.0R, QtyPhiText(sagQty, sagPhi) & " (ila.)", "")

                        If txtSol = "" AndAlso txtSag <> "" Then txtSol = txtSag
                        If txtSag = "" AndAlso txtSol <> "" Then txtSag = txtSol

                        Dim fb As String = QtyPhiText(infoBay.QtyMontaj, infoBay.PhiMontaj) & " (ila.)"
                        If txtSol = "" Then txtSol = fb
                        If txtSag = "" Then txtSag = fb

                        Dim midBay As Double = 0.5R * (bayInsideL + bayInsideR)
                        bayHalves.Add(Tuple.Create(bayInsideL, midBay, txtSol))
                        bayHalves.Add(Tuple.Create(midBay, bayInsideR, txtSag))
                    Next

                    Dim drawnPieces As New List(Of Tuple(Of Double, Double))()

                    If PASS_THROUGH_COLUMNS Then
                        Dim sA As Double = asmLeft + offsetUL
                        Dim eA As Double = Math.Max(asmRight, uMax) - offsetUL
                        If eA - sA > EPS Then
                            Dim pl As New Polyline()
                            ' Miters at 45° to bottom-left (225°)
                            pl.AddVertexAt(0, New Point2d(sA - miterLen, addlY - miterLen), 0, 0, 0)
                            pl.AddVertexAt(1, New Point2d(sA, addlY), 0, 0, 0)
                            pl.AddVertexAt(2, New Point2d(eA, addlY), 0, 0, 0)
                            pl.AddVertexAt(3, New Point2d(eA - miterLen, addlY - miterLen), 0, 0, 0)
                            pl.ColorIndex = 256
                            AppendOnLayer(btr, tr, pl, LYR_REBAR1, forceByLayer:=True)
                            drawnPieces.Add(Tuple.Create(sA, eA))
                        End If
                    Else
                        ' ----- ADDITIONAL (original)
                        For i As Integer = c0 To c1
                            Dim leftFace As Double = uniqueEdges(2 * i)
                            Dim rightFace As Double = uniqueEdges(2 * i + 1)
                            Dim insideStart As Double = leftFace + offsetUL
                            Dim insideEnd As Double = rightFace - offsetUL
                            If insideEnd - insideStart <= EPS Then Continue For

                            Dim prevRightFace As Double = If(i > c0, uniqueEdges(2 * (i - 1) + 1), asmLeft)
                            Dim nextLeftFace As Double = If(i < c1, uniqueEdges(2 * (i + 1)), asmRight)

                            Dim Lleft As Double = If(i > c0, leftFace - prevRightFace, 0.0R)
                            Dim Lright As Double = If(i < c1, nextLeftFace - rightFace, 0.0R)

                            Dim ruleLeft As Double = If(i > c0, Math.Max(Lleft / 4.0R, 50.0R * info.PhiLower), 0.0R)
                            Dim ruleRight As Double = If(i < c1, Math.Max(Lright / 4.0R, 50.0R * info.PhiLower), 0.0R)

                            Dim extLeftOut As Double = Round5Up(Math.Min(ruleLeft, Lleft))
                            Dim extRightOut As Double = Round5Up(Math.Min(ruleRight, Lright))

                            Dim s As Double, e As Double
                            If i = c0 Then
                                s = insideStart : e = rightFace + extRightOut
                            ElseIf i = c1 Then
                                s = leftFace - extLeftOut : e = insideEnd
                            Else
                                s = leftFace - extLeftOut : e = rightFace + extRightOut
                            End If

                            If i > c0 AndAlso s < prevRightFace Then s = prevRightFace
                            If i < c1 AndAlso e > nextLeftFace Then e = nextLeftFace
                            If e - s <= EPS Then Continue For

                            If i = c0 Then
                                Dim pl As New Polyline()
                                ' Miter at 45° to bottom-left (225°)
                                pl.AddVertexAt(0, New Point2d(s - miterLen, addlY - miterLen), 0, 0, 0)
                                pl.AddVertexAt(1, New Point2d(s, addlY), 0, 0, 0)
                                pl.AddVertexAt(2, New Point2d(e, addlY), 0, 0, 0)
                                pl.ColorIndex = 256
                                AppendOnLayer(btr, tr, pl, LYR_REBAR1, forceByLayer:=True)
                            ElseIf i = c1 Then
                                Dim pl As New Polyline()
                                pl.AddVertexAt(0, New Point2d(s, addlY), 0, 0, 0)
                                pl.AddVertexAt(1, New Point2d(e, addlY), 0, 0, 0)
                                ' Miter at 45° to bottom-left (225°)
                                pl.AddVertexAt(2, New Point2d(e - miterLen, addlY - miterLen), 0, 0, 0)
                                pl.ColorIndex = 256
                                AppendOnLayer(btr, tr, pl, LYR_REBAR1, forceByLayer:=True)
                            Else
                                Dim ln As New Line(New Point3d(s, addlY, 0), New Point3d(e, addlY, 0))
                                ln.ColorIndex = 256
                                AppendOnLayer(btr, tr, ln, LYR_REBAR1, forceByLayer:=True)
                            End If

                            drawnPieces.Add(Tuple.Create(s, e))
                        Next
                    End If

                    ' labels for additional (works for both modes)
                    If bayHalves.Count > 0 AndAlso drawnPieces.Count > 0 Then
                        Const MIN_LBL_LEN As Double = 3.0R
                        For Each pe In drawnPieces
                            Dim ps As Double = pe.Item1
                            Dim peX As Double = pe.Item2

                            Dim segs As New List(Of Tuple(Of Double, Double, String))()
                            For Each bh In bayHalves
                                Dim x0 As Double = Math.Max(ps, bh.Item1)
                                Dim x1 As Double = Math.Min(peX, bh.Item2)
                                If x1 - x0 > MIN_LBL_LEN Then
                                    segs.Add(Tuple.Create(x0, x1, bh.Item3))
                                End If
                            Next
                            If segs.Count = 0 Then Continue For

                            segs.Sort(Function(a, b) a.Item1.CompareTo(b.Item1))
                            Dim merged As New List(Of Tuple(Of Double, Double, String))()
                            Dim cur = segs(0)
                            For idx As Integer = 1 To segs.Count - 1
                                Dim nx = segs(idx)
                                If String.Equals(cur.Item3, nx.Item3, StringComparison.Ordinal) AndAlso Math.Abs(cur.Item2 - nx.Item1) < 0.5R Then
                                    cur = Tuple.Create(cur.Item1, nx.Item2, cur.Item3)
                                Else
                                    merged.Add(cur) : cur = nx
                                End If
                            Next
                            merged.Add(cur)

                            For Each m In merged
                                Dim mx As Double = 0.5R * (m.Item1 + m.Item2)
                                DrawPhiNote(btr, tr, mx, addlY, m.Item3, colTxt)
                            Next
                        Next
                    End If

                    ' ----- WEB (original)
                    Dim firstX As Double, lastX As Double
                    If PASS_THROUGH_COLUMNS Then
                        firstX = asmLeft + offsetWeb
                        lastX = Math.Max(asmRight, uMax) - offsetWeb
                    Else
                        firstX = asmLeft + offsetWeb
                        lastX = asmRight - offsetWeb
                    End If

                    If lastX - firstX > EPS Then
                        ' === target count: EXACTLY half of Gövde ===
                        Dim target As Integer = VerifyAndGetHalvedWebQty(ed, kKey, info, "Normal-Dup-Target")
                        LogWebBeam(ed, "Normal-Orig-Web-Before", kKey, info,
               $"firstX={firstX:F1} lastX={lastX:F1} targetNWeb={target}")

                        Dim drawn As Integer = 0

                        Dim topBound As Double = baseY_Upper
                        Dim botBound As Double = baseY_Lower
                        Dim usableH As Double = Math.Max(0.0R, topBound - botBound)
                        Dim webGap As Double = If(target > 0, usableH / (target + 1), 0.0R)

                        Dim lapWeb As Double = 50.0R * info.PhiMontaj
                        Dim stepLen As Double = Math.Max(5.0R, segLen - lapWeb)

                        ' label text (use halved qty)
                        Const LABEL_DY_WEB As Double = 2.0R
                        Dim webPhi As Double = If(info.PhiWeb > 0.0R, info.PhiWeb, info.PhiMontaj)
                        Dim webText As String = QtyPhiText(2 * target, webPhi)

                        For k As Integer = 1 To target
                            Dim yBase As Double = topBound - k * webGap
                            Dim currentX As Double = firstX
                            Dim segIndex As Integer = 0

                            Do While currentX < (lastX - EPS)
                                Dim remaining As Double = lastX - currentX
                                Dim length As Double = If(remaining >= segLen, segLen, remaining)
                                Dim nextX As Double = currentX + length

                                ' ORIGINAL webs: no overlay staggering (flat row)
                                Dim ySeg As Double = yBase

                                ' draw a lightweight segment
                                Dim ln As New Line(New Point3d(currentX, ySeg, 0), New Point3d(nextX, ySeg, 0))
                                ln.ColorIndex = 256
                                AppendOnLayer(btr, tr, ln, LYR_REBAR1, forceByLayer:=True)

                                ' draw label once on first row if you like
                                If k = 1 AndAlso Not String.IsNullOrWhiteSpace(webText) Then
                                    DrawPhiNote(btr, tr,
                            0.5R * (currentX + nextX),
                            ySeg + LABEL_DY_WEB,
                            webText,
                            colTxt)
                                End If

                                currentX += stepLen
                                segIndex += 1
                            Loop
                            drawn += 1
                        Next

                        LogWebBeam(ed, "Normal-Orig-Web-After", kKey, info, $"drawnNWeb={drawn}")

                        ' Optional overall center label (kept, uses halved qty)
                        If Not String.IsNullOrWhiteSpace(webText) Then
                            DrawPhiNote(btr, tr,
                    0.5R * (firstX + lastX),
                    0.5R * (baseY_Upper + baseY_Lower),
                    webText,
                    colTxt)
                            LogWebBeam(ed, "Normal-Orig-Web-Label", kKey, info, $"labelQty={target}")
                        End If
                    End If

                    ' ================= DUPLICATE SET =================
                    Dim nWebDup As Integer = GetHalvedWebQty(info)

                    Dim dupAddlY As Double = (topY - offsetUL) - DUP_OFFSET
                    Dim dupUpperY As Double = dupAddlY - DUP_SPACING
                    ' Draw only ONE web line, so use 1 instead of nWebDup for Y position calculation
                    Dim dupLowerY As Double = dupUpperY - 1 * DUP_SPACING - DUP_SPACING
                    InstallLenUpdater()

                    ' --- use extended right end just like originals (keep start; extend right if pass-through)
                    Dim endX_U_ext As Double = If(PASS_THROUGH_COLUMNS, Math.Max(endX_U, Math.Max(asmRight, uMax) - offsetUL), endX_U)
                    Dim lastXDupExt_Web As Double = If(PASS_THROUGH_COLUMNS, Math.Max(asmRight, uMax) - offsetWeb, asmRight - offsetWeb)

                    ' ===== UPPER duplicate  =====
                    For j = 0 To nMid
                        Dim rawX0 As Double, rawX1 As Double
                        If j = 0 Then
                            rawX0 = startX_U
                            rawX1 = If(nMid > 0, mids(0), endX_U_ext)
                        ElseIf j = nMid Then
                            rawX0 = If(nMid > 0, mids(nMid - 1), startX_U)
                            rawX1 = endX_U_ext
                        Else
                            rawX0 = mids(j - 1) : rawX1 = mids(j)
                        End If

                        Dim span As Double = Math.Max(0.0R, rawX1 - rawX0)
                        If span <= EPS Then Continue For

                        ' FIX: Resolve specific info for this segment (based on midpoint)
                        Dim segMidX As Double = (rawX0 + rawX1) / 2.0R
                        Dim kKeySeg As String = FindClosestKNumber(labels, segMidX, segMidX, topY, bottomY)
                        Dim infoSeg As BeamExcelRow = Nothing
                        If Not excelMap.TryGetValue(kKeySeg, infoSeg) Then infoSeg = info

                        ' show a visual overlap oz 
                        Dim oz As Double = Math.Min(overlayZoneUpper, 0.45R * span)
                        Dim x0 As Double = If(j = 0, rawX0, rawX0 - oz)
                        Dim x1 As Double = rawX1
                        If x1 - x0 <= EPS Then Continue For

                        Dim yDup As Double = If(j Mod 2 = 0, dupUpperY, dupUpperY + overlayHeight)

                        ' ====== ROUNDING (geometry) ======
                        ' Calculate rounding BEFORE drawing so we draw the correct final length
                        Dim baseLen As Double = x1 - x0
                        Dim targetLen As Double = Round5Up(baseLen)
                        Dim pad As Double = targetLen - baseLen     ' >= 0

                        ' ---- base legal window (NO crossing beyond what is displayed) ----
                        Dim allowedMin As Double, allowedMax As Double
                        If j = 0 Then
                            allowedMin = startX_U
                            allowedMax = rawX1       ' *** do not cross the first joint in the base step
                        ElseIf j = nMid Then
                            allowedMin = rawX0       ' *** do not cross the last joint in the base step
                            allowedMax = endX_U_ext
                        Else
                            allowedMin = x0          ' displayed left (rawX0 - oz)
                            allowedMax = rawX1       ' joint face on the right
                        End If

                        ' start from what we show
                        Dim newStart As Double = x0
                        Dim newEnd As Double = x1

                        ' track overlay actually used (for dim)
                        Dim usedOverlayRight As Double = 0.0R  ' borrowed beyond rawX1
                        Dim usedOverlayLeft As Double = 0.0R  ' borrowed beyond rawX0

                        If pad > EPS Then
                            ' ... [Keep existing Rounding Logic lines 370-397 unchanged] ...
                            ' (For brevity, assuming standard rounding logic remains here as per original file)
                            ' Only checkingpad > EPS logic block is skipped in display but must be kept in code

                            If j = nMid Then
                                Dim growL As Double = Math.Min(pad, Math.Max(0.0R, newStart - allowedMin))
                                newStart -= growL
                                Dim remain As Double = pad - growL
                                If remain > EPS Then
                                    Dim growR As Double = Math.Min(remain, Math.Max(0.0R, allowedMax - newEnd))
                                    newEnd += growR
                                    remain -= growR
                                End If
                                If remain > EPS Then
                                    Dim overlayAllow As Double = If(overlayZoneUpper > EPS, overlayZoneUpper, 50.0R * info.PhiLower)
                                    Dim crossMin As Double = Math.Max(startX_U, rawX0 - overlayAllow)
                                    Dim extraL As Double = Math.Min(remain, Math.Max(0.0R, newStart - crossMin))
                                    If extraL > EPS Then
                                        newStart -= extraL
                                        usedOverlayLeft = extraL
                                        remain -= extraL
                                    End If
                                    If remain > EPS Then
                                        Dim crossMax As Double = endX_U_ext
                                        Dim extraR As Double = Math.Min(remain, Math.Max(0.0R, crossMax - newEnd))
                                        If extraR > EPS Then
                                            newEnd += extraR
                                            remain -= extraR
                                        End If
                                    End If
                                End If
                            Else
                                Dim growR As Double = Math.Min(pad, Math.Max(0.0R, allowedMax - newEnd))
                                newEnd += growR
                                Dim remain As Double = pad - growR
                                If remain > EPS Then
                                    Dim growL As Double = Math.Min(remain, Math.Max(0.0R, newStart - allowedMin))
                                    newStart -= growL
                                    remain -= growL
                                End If
                                If remain > EPS Then
                                    Dim overlayAllow As Double = If(overlayZoneUpper > EPS, overlayZoneUpper, 50.0R * info.PhiLower)
                                    Dim crossMax As Double = Math.Min(endX_U_ext, rawX1 + overlayAllow)
                                    Dim extraR As Double = Math.Min(remain, Math.Max(0.0R, crossMax - newEnd))
                                    If extraR > EPS Then
                                        newEnd += extraR
                                        usedOverlayRight = extraR
                                        remain -= extraR
                                    End If
                                    If remain > EPS Then
                                        Dim crossMin As Double = Math.Max(startX_U, rawX0 - overlayAllow)
                                        Dim extraL As Double = Math.Min(remain, Math.Max(0.0R, newStart - crossMin))
                                        If extraL > EPS Then
                                            newStart -= extraL
                                            usedOverlayLeft = extraL
                                            remain -= extraL
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        ' ====== END ROUNDING ======

                        ' guard against bad limits
                        If newEnd - newStart <= EPS Then Continue For

                        ' ids for reactive labels
                        Dim dupUpperId As ObjectId = ObjectId.Null   ' full entity (polyline or line)
                        Dim segIdx_NoM As Integer = -1      ' horizontal segment index (for NoMiter)
                        Dim segIdx_ML As Integer = -1                ' left miter segment index
                        Dim segIdx_MR As Integer = -1                ' right miter segment index
                        Dim segDrawn As Boolean = False

                        If j = 0 Then
                            ' LEFT end (LEFT miter)
                            Dim pl As New Polyline()
                            ' Miter at 45° to bottom-left (225°)
                            pl.AddVertexAt(0, New Point2d(newStart - miterLen, dupUpperY - miterLen), 0, 0, 0) ' seg 0: 45° (left miter)
                            pl.AddVertexAt(1, New Point2d(newStart, dupUpperY), 0, 0, 0)
                            pl.AddVertexAt(2, New Point2d(newEnd, dupUpperY), 0, 0, 0)            ' seg 1: horizontal
                            pl.ColorIndex = 256
                            AppendOnLayer(btr, tr, pl, LYR_REBAR2, forceByLayer:=True)
                            dupUpperId = pl.ObjectId
                            segIdx_ML = 0
                            segIdx_NoM = 1
                            segDrawn = True

                        ElseIf j = nMid Then
                            ' RIGHT end (RIGHT miter)
                            Dim pl As New Polyline()
                            pl.AddVertexAt(0, New Point2d(newStart, yDup), 0, 0, 0)                 ' seg 0: horizontal
                            pl.AddVertexAt(1, New Point2d(newEnd, yDup), 0, 0, 0)
                            ' Miter at 45° to bottom-left (225°)
                            pl.AddVertexAt(2, New Point2d(newEnd - miterLen, yDup - miterLen), 0, 0, 0)      ' seg 1: 45° (right miter)
                            pl.ColorIndex = 256
                            AppendOnLayer(btr, tr, pl, LYR_REBAR2, forceByLayer:=True)
                            dupUpperId = pl.ObjectId
                            segIdx_NoM = 0
                            segIdx_MR = 1
                            segDrawn = True

                        Else
                            ' interior (no miters)
                            Dim ln As New Line(New Point3d(newStart, yDup, 0), New Point3d(newEnd, yDup, 0))
                            ln.ColorIndex = 256
                            AppendOnLayer(btr, tr, ln, LYR_REBAR2, forceByLayer:=True)
                            dupUpperId = ln.ObjectId
                            segDrawn = True
                        End If

                        ' ===== CREATE POS (Now that dupUpperId exists) =====
                        Dim yBar As Double = dupUpperY + 5.5
                        ' Use newStart/newEnd for label centering to align with drawn bar
                        Dim mid As New Point3d(0.5 * (newStart + newEnd), yBar, 0)

                        Dim pieceLenCm As Integer = CInt(Math.Round(Math.Max(0.0R, newEnd - newStart)))

                        Dim hasLeftHookForThisPiece As Boolean = (j = 0)
                        Dim hasRightHookForThisPiece As Boolean = (j = nMid)
                        Dim m1 As Integer = If(hasLeftHookForThisPiece, CInt(Math.Round(miterLen)), 0)
                        Dim m2 As Integer = If(hasRightHookForThisPiece, CInt(Math.Round(miterLen)), 0)
                        Dim miterCount As Integer = CInt(If(hasLeftHookForThisPiece, 1, 0) + If(hasRightHookForThisPiece, 1, 0))
                        Dim straightPart As Integer = pieceLenCm

                        Dim yerCode As String = "mon."
                        Dim phiMm As Integer = CInt(Math.Round(infoSeg.PhiMontaj * 10.0R))
                        Dim adetFromExcel As Integer = infoSeg.QtyMontaj

                        PosManager.Manager.AddDuplicateRebar(
                                 mid:=mid,
                                 phiMm:=phiMm,
                                 totalLen:=CInt(Math.Round(pieceLenCm + m1 + m2)),
                                 yer:=yerCode,
                                 adet:=adetFromExcel,
                                 miterCount:=miterCount,
                                 miter1:=m1,
                                 straightPart:=straightPart,
                                 miter2:=m2,
                                 linkedEntId:=dupUpperId  ' <--- LINKED ID
                               )

                        ' Overlay dim on joint overlap 
                        If j > 0 AndAlso oz > EPS Then
                            Dim p1 As New Point3d(rawX0 - oz, yDup, 0)   ' overlap start (left of joint)
                            Dim p2 As New Point3d(rawX0, yDup, 0)   ' joint face
                            If p2.DistanceTo(p1) > EPS Then
                                AddOverlayDimLinear(btr, tr, p1, p2, +5.0, LYR_DIM)   ' +5 cm above
                            End If
                        End If

                        ' ===== REACTIVE LENGTH LABELS (NO FIELDS) =====
                        If segDrawn AndAlso newEnd - newStart > EPS Then
                            Dim curDb = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database

                            Dim addLeftMiter As Boolean = (j = 0)
                            Dim addRightMiter As Boolean = (j = nMid)

                            ' 1) TOTAL (bind to whole entity)
                            '   If Not dupUpperId.IsNull Then
                            '  Dim totLen = ComputeLen(CType(tr.GetObject(dupUpperId, OpenMode.ForRead), Entity), LenKind.Total, 0)
                            ' Dim midU = 0.5 * (newStart + newEnd)
                            'Dim lblTot = CreateLengthLabel(tr, btr, New Point3d(midU, yDup + 6.0, 0), "L= " & (Silindi) FormatLen(totLen), 3.333, colTxt)
                            'LinkLabelTo(tr, curDb, lblTot, dupUpperId, LenKind.Total, segIdx_NoM, 0, +6.0)
                            'End If

                            ' 2) MITER-ONLY: left / right
                            If addLeftMiter AndAlso segIdx_ML >= 0 AndAlso Not dupUpperId.IsNull Then
                                Dim lenML = ComputeLen(CType(tr.GetObject(dupUpperId, OpenMode.ForRead), Entity), LenKind.NoMiter, segIdx_ML)
                                Dim lblML = CreateLengthLabel(tr, btr, New Point3d(newStart - 8.0, yDup - (miterLen / 2.0), 0), FormatLen(lenML), 3.333, colTxt)
                                LinkLabelTo(tr, curDb, lblML, dupUpperId, LenKind.MiterLeft, segIdx_ML, 0, 0)
                            End If
                            If addRightMiter AndAlso segIdx_MR >= 0 AndAlso Not dupUpperId.IsNull Then
                                Dim lenMR = ComputeLen(CType(tr.GetObject(dupUpperId, OpenMode.ForRead), Entity), LenKind.NoMiter, segIdx_MR)
                                Dim lblMR = CreateLengthLabel(tr, btr, New Point3d(newEnd + 8.0, yDup - (miterLen / 2.0), 0), FormatLen(lenMR), 3.333, colTxt)
                                LinkLabelTo(tr, curDb, lblMR, dupUpperId, LenKind.MiterRight, segIdx_MR, 0, 0)
                            End If

                            ' 3) NO-MITER (horizontal only) — only when any miter exists
                            If (addLeftMiter OrElse addRightMiter) AndAlso segIdx_NoM >= 0 AndAlso Not dupUpperId.IsNull Then
                                Dim lenNM = ComputeLen(CType(tr.GetObject(dupUpperId, OpenMode.ForRead), Entity), LenKind.NoMiter, segIdx_NoM)
                                Dim midU_NoM = 0.5 * (newStart + newEnd)
                                Dim lblNM = CreateLengthLabel(tr, btr, New Point3d(midU_NoM, yDup - 4.0, 0), FormatLen(lenNM), 3.333, colTxt)
                                LinkLabelTo(tr, curDb, lblNM, dupUpperId, LenKind.NoMiter, segIdx_NoM, 0, -4.0)
                            End If
                        End If
                    Next

                    ' LOWER duplicate (per span) + dashed cuts
                    For iSpace As Integer = 0 To (c1 - c0) - 1
                        Dim leftColIdx As Integer = c0 + iSpace
                        Dim rightColIdx As Integer = c0 + iSpace + 1
                        Dim leftRightFace As Double = uniqueEdges(2 * leftColIdx + 1)
                        Dim rightLeftFace As Double = uniqueEdges(2 * rightColIdx)

                        ' FIX: Resolve specific info for this span
                        Dim spanMidX As Double = (leftRightFace + rightLeftFace) / 2.0R
                        Dim kKeySpan As String = FindClosestKNumber(labels, spanMidX, spanMidX, topY, bottomY)
                        Dim infoSpan As BeamExcelRow = Nothing
                        If Not excelMap.TryGetValue(kKeySpan, infoSpan) Then infoSpan = info

                        Dim x0L As Double, x1L As Double
                        If iSpace = 0 Then
                            x0L = asmLeft + offsetLower
                            x1L = rightLeftFace + (50.0R * infoSpan.PhiLower)
                        ElseIf iSpace = (c1 - c0) - 1 Then
                            x0L = leftRightFace - (50.0R * infoSpan.PhiLower)
                            x1L = If(PASS_THROUGH_COLUMNS, Math.Max(asmRight, uMax) - offsetLower, asmRight - offsetLower)
                        Else
                            x0L = leftRightFace - (50.0R * infoSpan.PhiLower)
                            x1L = rightLeftFace + (50.0R * infoSpan.PhiLower)
                        End If

                        Dim yDupL As Double = dupLowerY
                        If (iSpace Mod 2) = 1 Then yDupL -= overlayHeight

                        If x1L - x0L > EPS Then

                            ' =======================
                            ' ROUNDING PATCH (geometry) - Move BEFORE drawing
                            ' =======================
                            Dim rawLenL As Double = x1L - x0L
                            Dim targetLen As Double = Round5Up(rawLenL)
                            Dim pad As Double = targetLen - rawLenL

                            If pad > EPS Then
                                Dim allowedMin As Double, allowedMax As Double
                                If iSpace = 0 Then
                                    allowedMin = asmLeft + offsetLower
                                    allowedMax = rightLeftFace + (50.0R * info.PhiLower)
                                ElseIf iSpace = (c1 - c0) - 1 Then
                                    allowedMin = leftRightFace - (50.0R * info.PhiLower)
                                    allowedMax = If(PASS_THROUGH_COLUMNS, Math.Max(asmRight, uMax) - offsetLower, asmRight - offsetLower)
                                Else
                                    allowedMin = leftRightFace - (50.0R * info.PhiLower)
                                    allowedMax = rightLeftFace + (50.0R * info.PhiLower)
                                End If

                                If x0L < allowedMin Then x0L = allowedMin
                                If x1L > allowedMax Then x1L = allowedMax

                                ' 1) Use the base window first
                                Dim growR As Double = Math.Min(pad, Math.Max(0.0R, allowedMax - x1L))
                                x1L += growR
                                Dim remain As Double = pad - growR
                                If remain > EPS Then
                                    Dim growL As Double = Math.Min(remain, Math.Max(0.0R, x0L - allowedMin))
                                    x0L -= growL
                                    remain -= growL
                                End If

                                ' 2) Borrow from overlay zone
                                Dim usedOverlayRight As Double = 0.0R
                                Dim usedOverlayLeft As Double = 0.0R

                                If remain > EPS Then
                                    Const ALLOW_OVERLAY_FOR_ROUNDING As Boolean = True
                                    If ALLOW_OVERLAY_FOR_ROUNDING Then
                                        Dim overlayAllow As Double = 50.0R * info.PhiLower
                                        Dim absMin As Double = asmLeft + offsetLower
                                        Dim absMax As Double = If(iSpace = (c1 - c0) - 1 AndAlso PASS_THROUGH_COLUMNS,
                                                                  Math.Max(asmRight, uMax) - offsetLower,
                                                                  asmRight - offsetLower)

                                        Dim crossRightCap As Double = If(iSpace = (c1 - c0) - 1, absMax, Math.Min(absMax, rightLeftFace + (50.0R * info.PhiLower) + overlayAllow))
                                        Dim crossLeftCap As Double = If(iSpace = 0, absMin, Math.Max(absMin, leftRightFace - (50.0R * info.PhiLower) - overlayAllow))

                                        If remain > EPS Then
                                            Dim extraR As Double = Math.Min(remain, Math.Max(0.0R, crossRightCap - x1L))
                                            x1L += extraR
                                            usedOverlayRight = extraR
                                            remain -= extraR
                                        End If
                                        If remain > EPS Then
                                            Dim extraL As Double = Math.Min(remain, Math.Max(0.0R, x0L - crossLeftCap))
                                            x0L -= extraL
                                            usedOverlayLeft = extraL
                                            remain -= extraL
                                        End If
                                    End If
                                End If

                                ' Draw rounding dims
                                If usedOverlayRight > EPS Then
                                    AddOverlayDimLinear(btr, tr, New Point3d(rightLeftFace, yDupL, 0), New Point3d(x1L, yDupL, 0), +5.0, LYR_DIM)
                                End If
                                If usedOverlayLeft > EPS Then
                                    AddOverlayDimLinear(btr, tr, New Point3d(x0L, yDupL, 0), New Point3d(leftRightFace, yDupL, 0), +5.0, LYR_DIM)
                                End If
                            End If
                            ' --- END ROUNDING ---

                            ' --- draw the bar ---
                            Dim dupLowerId As ObjectId = ObjectId.Null
                            Dim segIdx_NoM As Integer = -1
                            Dim segIdx_ML As Integer = -1
                            Dim segIdx_MR As Integer = -1

                            If iSpace = 0 Then
                                ' Leftmost span
                                Dim pl As New Polyline()
                                ' Miter at 45° to top-left (135°)
                                pl.AddVertexAt(0, New Point2d(x0L - miterLen, yDupL + miterLen), 0, 0, 0)
                                pl.AddVertexAt(1, New Point2d(x0L, yDupL), 0, 0, 0)
                                pl.AddVertexAt(2, New Point2d(x1L, yDupL), 0, 0, 0)
                                pl.ColorIndex = 256
                                AppendOnLayer(btr, tr, pl, LYR_REBAR2, forceByLayer:=True)
                                dupLowerId = pl.ObjectId
                                segIdx_ML = 0
                                segIdx_NoM = 1

                            ElseIf iSpace = (c1 - c0) - 1 Then
                                ' Rightmost span
                                Dim pl As New Polyline()
                                pl.AddVertexAt(0, New Point2d(x0L, yDupL), 0, 0, 0)
                                pl.AddVertexAt(1, New Point2d(x1L, yDupL), 0, 0, 0)
                                ' Miter at 45° to top-left (135°)
                                pl.AddVertexAt(2, New Point2d(x1L - miterLen, yDupL + miterLen), 0, 0, 0)
                                pl.ColorIndex = 256
                                AppendOnLayer(btr, tr, pl, LYR_REBAR2, forceByLayer:=True)
                                dupLowerId = pl.ObjectId
                                segIdx_NoM = 0
                                segIdx_MR = 1
                            Else
                                ' Interior
                                Dim ln As New Line(New Point3d(x0L, yDupL, 0), New Point3d(x1L, yDupL, 0))
                                ln.ColorIndex = 256
                                AppendOnLayer(btr, tr, ln, LYR_REBAR2, forceByLayer:=True)
                                dupLowerId = ln.ObjectId
                            End If

                            ' ===== register POS (Now that we have dupLowerId) =====
                            Dim yBar As Double = dupLowerY - 5.5
                            Dim mid As New Point3d(0.5 * (x0L + x1L), yBar, 0)

                            Dim pieceLenCm As Integer = CInt(Math.Round(Math.Max(0.0R, x1L - x0L)))
                            Dim hasLeftHookForThisPiece As Boolean = (iSpace = 0)
                            Dim hasRightHookForThisPiece As Boolean = (iSpace = (c1 - c0) - 1)
                            Dim m1 As Integer = If(hasLeftHookForThisPiece, CInt(Math.Round(miterLen)), 0)
                            Dim m2 As Integer = If(hasRightHookForThisPiece, CInt(Math.Round(miterLen)), 0)
                            Dim miterCount As Integer = CInt(If(hasLeftHookForThisPiece, 1, 0) + If(hasRightHookForThisPiece, 1, 0))
                            Dim straightPart As Integer = pieceLenCm
                            Dim yerCode As String = "alt."
                            Dim phiMm As Integer = CInt(Math.Round(infoSpan.PhiLower * 10.0R))
                            Dim adetFromExcel As Integer = Math.Max(1, infoSpan.QtyLower)

                            PosManager.Manager.AddDuplicateRebar(
                                mid:=mid,
                                phiMm:=phiMm,
                                totalLen:=CInt(Math.Round(straightPart + m1 + m2)),
                                yer:=yerCode,
                                adet:=adetFromExcel,
                                miterCount:=miterCount,
                                miter1:=m1,
                                straightPart:=straightPart,
                                miter2:=m2,
                                linkedEntId:=dupLowerId  ' <--- LINKED ID
                            )

                            ' --- side lap dimension (left faces)
                            If iSpace > 0 Then
                                Dim lapLen As Double = 50.0R * info.PhiLower
                                Dim p1 As New Point3d(leftRightFace - lapLen, yDupL, 0) ' from lap start
                                Dim p2 As New Point3d(leftRightFace, yDupL, 0)          ' to face
                                AddOverlayDimLinear(btr, tr, p1, p2, -5.0, LYR_DIM)     ' 5 cm below
                            End If

                            ' --- dashed cuts 
                            Dim xLL As Double = uniqueEdges(2 * leftColIdx)
                            Dim xLR As Double = uniqueEdges(2 * leftColIdx + 1)
                            Dim xRL As Double = uniqueEdges(2 * rightColIdx)
                            Dim xRR As Double = uniqueEdges(2 * rightColIdx + 1)
                            DashedColumnCuts.DrawColumnBottomDashedCuts(btr, tr, xLL, xLR, bottomY, yDupL, 10.0)
                            DashedColumnCuts.DrawColumnBottomDashedCuts(btr, tr, xRL, xRR, bottomY, yDupL, 10.0)


                            ' ===== REACTIVE LENGTH LABELS (NO FIELDS) =====
                            Dim curDb = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
                            If Not dupLowerId.IsNull Then
                                ' 1) TOTAL length — above the bar
                                ' Dim totLen = ComputeLen(CType(tr.GetObject(dupLowerId, OpenMode.ForRead), Entity), LenKind.Total, 0)
                                ' Dim midL = 0.5 * (x0L + x1L)
                                'Dim lblTot = CreateLengthLabel(tr, btr, New Point3d(midL, yDupL + 6.0, 0), "L=" & (Silindi) FormatLen(totLen), 3.333, colTxt)
                                'LinkLabelTo(tr, curDb, lblTot, dupLowerId, LenKind.Total, segIdx_NoM, 0, +6.0)

                                ' 2) MITER-ONLY lengths (if present)
                                If segIdx_ML >= 0 Then
                                    Dim lenML = ComputeLen(CType(tr.GetObject(dupLowerId, OpenMode.ForRead), Entity), LenKind.NoMiter, segIdx_ML)
                                    Dim lblML = CreateLengthLabel(tr, btr, New Point3d(x0L - 8.0, yDupL + (miterLen / 2.0), 0), FormatLen(lenML), 3.333, colTxt)
                                    LinkLabelTo(tr, curDb, lblML, dupLowerId, LenKind.MiterLeft, segIdx_ML, 0, 0)
                                End If
                                If segIdx_MR >= 0 Then
                                    Dim lenMR = ComputeLen(CType(tr.GetObject(dupLowerId, OpenMode.ForRead), Entity), LenKind.NoMiter, segIdx_MR)
                                    Dim lblMR = CreateLengthLabel(tr, btr, New Point3d(x1L + 8.0, yDupL + (miterLen / 2.0), 0), FormatLen(lenMR), 3.333, colTxt)
                                    LinkLabelTo(tr, curDb, lblMR, dupLowerId, LenKind.MiterRight, segIdx_MR, 0, 0)
                                End If

                                ' 3) NO-MITER (horizontal only) — only when a miter exists
                                If (segIdx_NoM >= 0) AndAlso ((segIdx_ML >= 0) OrElse (segIdx_MR >= 0)) Then
                                    Dim lenNM = ComputeLen(CType(tr.GetObject(dupLowerId, OpenMode.ForRead), Entity), LenKind.NoMiter, segIdx_NoM)
                                    Dim midL_NoM = 0.5 * (x0L + x1L)
                                    Dim lblNM = CreateLengthLabel(tr, btr, New Point3d(midL_NoM, yDupL - 4.0, 0), FormatLen(lenNM), 3.333, colTxt)
                                    LinkLabelTo(tr, curDb, lblNM, dupLowerId, LenKind.NoMiter, segIdx_NoM, 0, -4.0)
                                End If
                            End If
                        End If
                    Next

                    ' ===== ADDITIONAL duplicate =====
                    Dim addlYdup As Double = dupAddlY

                    For i As Integer = c0 To c1
                        Dim leftFace As Double = uniqueEdges(2 * i)
                        Dim rightFace As Double = uniqueEdges(2 * i + 1)
                        Dim insideStart As Double = leftFace + offsetUL
                        Dim insideEnd As Double = rightFace - offsetUL

                        If insideEnd - insideStart <= EPS Then Continue For

                        Dim prevRightFace As Double = If(i > c0, uniqueEdges(2 * (i - 1) + 1), asmLeft)
                        Dim nextLeftFace As Double = If(i < c1, uniqueEdges(2 * (i + 1)), asmRight)

                        Dim Lleft As Double = If(i > c0, leftFace - prevRightFace, 0.0R)
                        Dim Lright As Double = If(i < c1, nextLeftFace - rightFace, 0.0R)

                        Dim ruleLeft As Double = If(i > c0, Math.Max(Lleft / 4.0R, 50.0R * info.PhiLower), 0.0R)
                        Dim ruleRight As Double = If(i < c1, Math.Max(Lright / 4.0R, 50.0R * info.PhiLower), 0.0R)

                        Dim extLeftOut As Double = Round5Up(Math.Min(ruleLeft, Lleft))
                        Dim extRightOut As Double = Round5Up(Math.Min(ruleRight, Lright))

                        Dim s As Double, e As Double
                        If i = c0 Then
                            s = insideStart
                            e = rightFace + extRightOut
                        ElseIf i = c1 Then
                            s = leftFace - extLeftOut
                            ' extend far-right piece to the right beam end when pass-through
                            Dim insideEndExt As Double = If(PASS_THROUGH_COLUMNS, Math.Max(insideEnd, Math.Max(asmRight, uMax) - offsetUL), insideEnd)
                            e = insideEndExt
                        Else
                            s = leftFace - extLeftOut
                            e = rightFace + extRightOut
                        End If

                        ' Safety: keep extensions within neighbor limits
                        If i > c0 AndAlso s < prevRightFace Then s = prevRightFace
                        If i < c1 AndAlso e > nextLeftFace Then e = nextLeftFace

                        ' --- ROUNDING PATCH (geometry) ---
                        Dim rawLen As Double = e - s
                        If rawLen > EPS Then
                            Dim targetLen As Double = Round5Up(rawLen)
                            Dim pad As Double = targetLen - rawLen
                            If pad > EPS Then
                                Dim allowedMin As Double = If(i = c0, insideStart, prevRightFace)
                                Dim allowedMax As Double
                                If i = c1 Then
                                    allowedMax = If(PASS_THROUGH_COLUMNS, Math.Max(insideEnd, Math.Max(asmRight, uMax) - offsetUL), insideEnd)
                                Else
                                    allowedMax = nextLeftFace
                                End If

                                If s < allowedMin Then s = allowedMin
                                If e > allowedMax Then e = allowedMax

                                Dim growR As Double = Math.Min(pad, Math.Max(0.0R, allowedMax - e))
                                e += growR
                                Dim remain As Double = pad - growR
                                If remain > EPS Then
                                    Dim growL As Double = Math.Min(remain, Math.Max(0.0R, s - allowedMin))
                                    s -= growL
                                End If
                            End If
                        End If
                        ' --- END ROUNDING PATCH ---

                        If e - s <= EPS Then Continue For

                        ' ---- DRAW ----
                        Dim dupAddlId As ObjectId = ObjectId.Null
                        Dim segIdx_NoM As Integer = -1
                        Dim segIdx_ML As Integer = -1
                        Dim segIdx_MR As Integer = -1

                        If i = c0 Then
                            ' Leftmost: LEFT miter (down)
                            Dim pl As New Polyline()
                            ' Miter at 45° to bottom-left (225°)
                            pl.AddVertexAt(0, New Point2d(s - miterLen, addlYdup - miterLen), 0, 0, 0)
                            pl.AddVertexAt(1, New Point2d(s, addlYdup), 0, 0, 0)
                            pl.AddVertexAt(2, New Point2d(e, addlYdup), 0, 0, 0)
                            pl.ColorIndex = 256
                            AppendOnLayer(btr, tr, pl, LYR_REBAR2, forceByLayer:=True)
                            dupAddlId = pl.ObjectId
                            segIdx_ML = 0
                            segIdx_NoM = 1
                        ElseIf i = c1 Then
                            ' Rightmost: RIGHT miter (down)
                            Dim pl As New Polyline()
                            pl.AddVertexAt(0, New Point2d(s, addlYdup), 0, 0, 0)
                            pl.AddVertexAt(1, New Point2d(e, addlYdup), 0, 0, 0)
                            ' Miter at 45° to bottom-left (225°)
                            pl.AddVertexAt(2, New Point2d(e - miterLen, addlYdup - miterLen), 0, 0, 0)
                            pl.ColorIndex = 256
                            AppendOnLayer(btr, tr, pl, LYR_REBAR2, forceByLayer:=True)
                            dupAddlId = pl.ObjectId
                            segIdx_NoM = 0
                            segIdx_MR = 1
                        Else
                            ' Interior
                            Dim ln As New Line(New Point3d(s, addlYdup, 0), New Point3d(e, addlYdup, 0))
                            ln.ColorIndex = 256
                            AppendOnLayer(btr, tr, ln, LYR_REBAR2, forceByLayer:=True)
                            dupAddlId = ln.ObjectId
                        End If

                        ' ---- EXTENSION DIMS ----
                        Dim leftExtActual As Double = If(i > c0, Math.Max(0.0R, leftFace - s), 0.0R)
                        Dim rightExtActual As Double = If(i < c1, Math.Max(0.0R, e - rightFace), 0.0R)

                        If rightExtActual > EPS Then
                            Dim p1 As New Point3d(rightFace, addlYdup, 0)
                            Dim p2 As New Point3d(rightFace + rightExtActual, addlYdup, 0)
                            AddOverlayDimLinear(btr, tr, p1, p2, +5.0, LYR_DIM)
                        End If
                        If leftExtActual > EPS Then
                            Dim p1 As New Point3d(leftFace - leftExtActual, addlYdup, 0)
                            Dim p2 As New Point3d(leftFace, addlYdup, 0)
                            AddOverlayDimLinear(btr, tr, p1, p2, +5.0, LYR_DIM)
                        End If

                        ' ===== REGISTER POS =====
                        Dim x0 As Double = s
                        Dim x1 As Double = e
                        Dim yBar As Double = addlYdup + 5.5
                        Dim mid As New Point3d(0.5 * (x0 + x1), yBar, 0)
                        Dim pieceLenCm As Integer = CInt(Math.Round(Math.Max(0.0R, x1 - x0)))

                        Dim hasLeftHookForThisPiece As Boolean = (i = c0)
                        Dim hasRightHookForThisPiece As Boolean = (i = c1)
                        Dim m1 As Integer = If(hasLeftHookForThisPiece, CInt(Math.Round(miterLen)), 0)
                        Dim m2 As Integer = If(hasRightHookForThisPiece, CInt(Math.Round(miterLen)), 0)
                        Dim miterCount As Integer = CInt(If(hasLeftHookForThisPiece, 1, 0) + If(hasRightHookForThisPiece, 1, 0))
                        Dim straightPart As Integer = pieceLenCm
                        Dim yerCode As String = "ila."
                        Dim targetQty As Integer = 0
                        Dim targetPhi As Double = 0.0

                        ' Check Right Span
                        If i < c1 Then
                            Dim leftFace_R = uniqueEdges(2 * i + 1)
                            Dim rightFace_R = uniqueEdges(2 * (i + 1))
                            Dim kR = FindClosestKNumber(labels, leftFace_R, rightFace_R, topY, bottomY)
                            Dim rowR As BeamExcelRow = Nothing
                            If excelMap.TryGetValue(kR, rowR) AndAlso rowR.QtyAddlLeft > 0 Then
                                targetQty = rowR.QtyAddlLeft
                                targetPhi = rowR.PhiAddlLeft
                            End If
                        End If
                        ' Check Left Span
                        If i > c0 Then
                            Dim leftFace_L = uniqueEdges(2 * (i - 1) + 1)
                            Dim rightFace_L = uniqueEdges(2 * i)
                            Dim kL = FindClosestKNumber(labels, leftFace_L, rightFace_L, topY, bottomY)
                            Dim rowL As BeamExcelRow = Nothing
                            If excelMap.TryGetValue(kL, rowL) AndAlso rowL.QtyAddlRight > 0 Then
                                If rowL.QtyAddlRight > targetQty Then
                                    targetQty = rowL.QtyAddlRight
                                    targetPhi = rowL.PhiAddlRight
                                End If
                            End If
                        End If
                        If targetQty <= 0 Then
                            targetQty = info.QtyLower
                            targetPhi = info.PhiLower
                        End If

                        Dim phiMm As Integer = CInt(Math.Round(targetPhi * 10.0R))
                        Dim adetFromExcel As Integer = targetQty

                        PosManager.Manager.AddDuplicateRebar(
                            mid:=mid,
                            phiMm:=phiMm,
                            totalLen:=CInt(Math.Round(straightPart + m1 + m2)),
                            yer:=yerCode,
                            adet:=adetFromExcel,
                            miterCount:=miterCount,
                            miter1:=m1,
                            straightPart:=straightPart,
                            miter2:=m2,
                            linkedEntId:=dupAddlId
                        )

                        ' ===== REACTIVE LABELS =====
                        If Not dupAddlId.IsNull Then
                            Dim curDb = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database

                            ' Miter Labels
                            If segIdx_ML >= 0 Then
                                Dim lenML = ComputeLen(CType(tr.GetObject(dupAddlId, OpenMode.ForRead), Entity), LenKind.NoMiter, segIdx_ML)
                                Dim lblML = CreateLengthLabel(tr, btr, New Point3d(s - 8.0, addlYdup - (miterLen / 2.0), 0), FormatLen(lenML), 3.333, colTxt)
                                LinkLabelTo(tr, curDb, lblML, dupAddlId, LenKind.MiterLeft, segIdx_ML, 0, 0)
                            End If
                            If segIdx_MR >= 0 Then
                                Dim lenMR = ComputeLen(CType(tr.GetObject(dupAddlId, OpenMode.ForRead), Entity), LenKind.NoMiter, segIdx_MR)
                                Dim lblMR = CreateLengthLabel(tr, btr, New Point3d(e + 8.0, addlYdup - (miterLen / 2.0), 0), FormatLen(lenMR), 3.333, colTxt)
                                LinkLabelTo(tr, curDb, lblMR, dupAddlId, LenKind.MiterRight, segIdx_MR, 0, 0)
                            End If

                            ' Horizontal Label
                            If (segIdx_NoM >= 0) AndAlso ((segIdx_ML >= 0) OrElse (segIdx_MR >= 0)) Then
                                Dim lenNM = ComputeLen(CType(tr.GetObject(dupAddlId, OpenMode.ForRead), Entity), LenKind.NoMiter, segIdx_NoM)
                                Dim midA_NoM = 0.5 * (s + e)
                                Dim lblNM = CreateLengthLabel(tr, btr, New Point3d(midA_NoM, addlYdup - 4.0, 0), FormatLen(lenNM), 3.333, colTxt)
                                LinkLabelTo(tr, curDb, lblNM, dupAddlId, LenKind.NoMiter, segIdx_NoM, 0, -4.0)
                            End If
                        End If
                    Next ' End of Loop

                    ' WEB duplicate - draw ONE line with nxm ADET format
                    Dim firstXDup As Double = asmLeft + offsetWeb
                    Dim lastXDup As Double = lastXDupExt_Web                                ' << extend right like originals
                    If lastXDup - firstXDup > EPS Then
                        Dim qtyDup2 As Integer = If(info.WebQty > 0, info.WebQty, 2)
                        Dim nWebDup2 As Integer = GetHalvedWebQty(info)

                        Dim lapWebDup As Double = 50.0R * info.PhiMontaj
                        Dim stepLenDup As Double = Math.Max(5.0R, segLen - lapWebDup)

                        ' Draw only ONE web rebar line (instead of nWebDup2 lines)
                        ' ADET will show "nxm" format (e.g., "3x2" for 6 total rebars)
                        Dim yDupWeb As Double = dupUpperY - 1 * DUP_SPACING
                        Dim currentX = firstXDup
                        Dim segIndex As Integer = 0

                        Do While currentX < (lastXDup - EPS)
                            Dim remaining = lastXDup - currentX
                            Dim length = If(remaining >= segLen, segLen, remaining)
                            Dim nextX = currentX + length
                            Dim ySeg As Double = If((segIndex And 1) = 0, yDupWeb, yDupWeb - overlayHeight)

                            ' draw the segment
                            Dim plW2 As New Polyline()
                            plW2.AddVertexAt(0, New Point2d(currentX, ySeg), 0, 0, 0)
                            plW2.AddVertexAt(1, New Point2d(nextX, ySeg), 0, 0, 0)
                            plW2.ColorIndex = 256
                            AppendOnLayer(btr, tr, plW2, LYR_REBAR2, forceByLayer:=True)

                            ' ===== register POS for this WEB (gövde) duplicate segment =====
                            Dim x0 As Double = currentX
                            Dim x1 As Double = nextX
                            Dim yBar As Double = ySeg + 5.5            ' nudge label a bit above the web segment (optional)
                            Dim mid As New Point3d(0.5 * (x0 + x1), yBar, 0)

                            Dim pieceLenCm As Integer = CInt(Math.Round(Math.Abs(x1 - x0)))

                            ' Web duplicates are straight pieces with no hooks
                            Dim miterCount As Integer = 0
                            Dim m1 As Integer = 0, m2 As Integer = 0
                            Dim straightPart As Integer = pieceLenCm

                            Dim yerCode As String = "gov."
                            ' If your web steel uses a dedicated diameter (e.g., info.PhiWeb) use that.
                            ' If not, keep PhiMontaj as you already used for lap length.
                            Dim phiMm As Integer = CInt(Math.Round(info.PhiWeb * 10.0R))   ' <-- change to PhiMontaj if that's what you use

                            ' Use nxm format for ADET (e.g., "3x2" for 6 total rebars)
                            Dim adetText As String = If(nWebDup2 > 1, $"{nWebDup2}x2", "2")

                            PosManager.Manager.AddDuplicateRebarWithAdetText(
                                                mid:=mid,
                                                phiMm:=phiMm,
                                                totalLen:=straightPart,      ' no hooks added for web piece
                                                yer:=yerCode,
                                                adetText:=adetText,
                                                miterCount:=miterCount,
                                                miter1:=m1,
                                                straightPart:=straightPart,
                                                miter2:=m2,
                                                linkedEntId:=plW2.ObjectId
                                            )
                            ' ===============================================================


                            ' LAP indicator as real DimLinear over THIS segment only
                            If segIndex > 0 Then
                                Dim ovEnd As Double = Math.Min(currentX + lapWebDup, nextX) ' cap to segment end
                                If ovEnd - currentX > EPS Then
                                    Dim p1 As New Point3d(currentX, ySeg, 0)
                                    Dim p2 As New Point3d(ovEnd, ySeg, 0)
                                    AddOverlayDimLinear(btr, tr, p1, p2, +5.0, LYR_DIM)  ' +5 above
                                End If
                            End If

                            currentX += stepLenDup
                            segIndex += 1
                        Loop


                    End If

                Next ' grp
            Next ' r
            Dim btrCur = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
            PosManager.Manager.Commit(db, tr, btrCur, markerRadius:=5.5)
            PosManager.Manager.ClearPending()
            tr.Commit()
        End Using

        If doc IsNot Nothing Then
            Using doc.LockDocument()
                RefreshAllLenLabelsForDb(doc.Database)  ' lock-free helper
            End Using
        End If

        PlacePlanBeamNameSizeLabels(db, ed, columnEdges, beamLines, labels,
                            xGap:=1.5, yGap:=1.5, textH:=10.0)

        ' 6) Draw sections to the right of the plan if user requested
        If dlg.DrawSections Then
            DrawSectionsRightOfPlan(
            db:=db,
            ed:=ed,
            excelMap:=excelMap,
            columnEdges:=columnEdges,
            beamLines:=beamLines,
            labels:=labels,
            START_GAP_RIGHT_OF_LAST:=200.0,
            GAP_BETWEEN_SECTIONS:=150.0,
            coverParam:=offsetUL)
        End If



    End Sub


#End Region


#Region "SplitAssemblies + UI"

    Private Function SplitAssemblies(uniqueEdges As List(Of Double)) As List(Of Tuple(Of Integer, Integer))
        Dim ranges As New List(Of Tuple(Of Integer, Integer))()
        Dim numCols As Integer = uniqueEdges.Count \ 2
        If numCols <= 0 Then Return ranges

        Dim cx As New List(Of Double)()
        For i = 0 To numCols - 1
            Dim leftX = uniqueEdges(2 * i)
            Dim rightX = uniqueEdges(2 * i + 1)
            cx.Add((leftX + rightX) / 2.0)
        Next

        Dim diffs As New List(Of Double)()
        For i = 0 To cx.Count - 2
            diffs.Add(cx(i + 1) - cx(i))
        Next
        Dim medianSpan As Double = If(diffs.Count = 0, 0, diffs.OrderBy(Function(d) d).ElementAt(diffs.Count \ 2))
        If medianSpan <= 0 Then
            ranges.Add(Tuple.Create(0, numCols - 1))
            Return ranges
        End If

        Dim thresh = 1.8 * medianSpan
        Dim startIdx = 0
        For i = 0 To cx.Count - 2
            If (cx(i + 1) - cx(i)) > thresh Then
                ranges.Add(Tuple.Create(startIdx, i))
                startIdx = i + 1
            End If
        Next
        ranges.Add(Tuple.Create(startIdx, numCols - 1))
        Return ranges
    End Function

    Private Class RebarParamsForm
        Inherits Windows.Forms.Form

        Private nudOverlay As Windows.Forms.NumericUpDown
        Private nudOffsetUL As Windows.Forms.NumericUpDown
        Public nudCarpanDonati As Windows.Forms.NumericUpDown
        Private chkDrawSections As Windows.Forms.CheckBox


        Public ReadOnly Property OverlayUpper As Double
            Get
                Return CDbl(nudOverlay.Value)
            End Get
        End Property

        Public ReadOnly Property OffsetUL As Double
            Get
                Return CDbl(nudOffsetUL.Value)
            End Get
        End Property

        Public ReadOnly Property CarpanDonati As Double
            Get
                Return CDbl(nudCarpanDonati.Value)
            End Get
        End Property

        Public ReadOnly Property DrawSections As Boolean
            Get
                Return chkDrawSections.Checked
            End Get
        End Property

        Public Sub New()
            Me.Text = "Rebar Parameters"
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
            Me.StartPosition = Windows.Forms.FormStartPosition.CenterParent
            Me.MaximizeBox = False : Me.MinimizeBox = False
            Me.Width = 380 : Me.Height = 320

            Dim overlayDefault As Decimal = CDec(Settings.GetDouble("OverlayUpper", 80))
            Dim offsetULDefault As Decimal = CDec(Settings.GetDouble("OffsetUL", 30))
            Dim carpanDefault As Decimal = CDec(Settings.GetDouble("CarpanDonati", 1))

            Dim y As Integer = 20
            Dim rowGap As Integer = 40

            Dim lbl1 As New Windows.Forms.Label() With {
                .Text = "Montaj Bindirme (cm):",
                .AutoSize = True,
                .Left = 16,
                .Top = y
            }
            nudOverlay = New Windows.Forms.NumericUpDown() With {
                .Left = 220,
                .Top = y - 4,
                .Width = 120,
                .Minimum = 1,
                .Maximum = 100000,
                .DecimalPlaces = 0,
                .Increment = 5,
                .Value = overlayDefault
            }
            y += rowGap

            Dim lbl2 As New Windows.Forms.Label() With {
                .Text = "Pas payı (cm):",
                .AutoSize = True,
                .Left = 16,
                .Top = y
            }
            nudOffsetUL = New Windows.Forms.NumericUpDown() With {
                .Left = 220,
                .Top = y - 4,
                .Width = 120,
                .Minimum = 1,
                .Maximum = 100000,
                .DecimalPlaces = 0,
                .Increment = 1,
                .Value = offsetULDefault
            }
            y += rowGap

            Dim lbl3 As New Windows.Forms.Label() With {
                .Text = "Donati Carpani",
                .AutoSize = True,
                .Left = 16,
                .Top = y
            }

            nudCarpanDonati = New Windows.Forms.NumericUpDown() With {
                .Left = 220,
                .Top = y - 4,
                .Width = 120,
                .Minimum = 1,
                .Maximum = 100000,
                .DecimalPlaces = 0,
                .Increment = 1,
                .Value = carpanDefault
                }
            y += rowGap

            ' new checkbox to toggle drawing sections
            chkDrawSections = New Windows.Forms.CheckBox() With {
                .Left = 16,
                .Top = y,
                .AutoSize = True,
                .Text = "Also draw sections (use same Excel/CSV)"
            }
            chkDrawSections.Checked = True  ' default on
            y += rowGap

            Dim btnOK As New Windows.Forms.Button() With {
                .Text = "OK",
                .Left = 170,
                .Top = y,
                .DialogResult = Windows.Forms.DialogResult.OK,
                .Width = 75
            }
            Dim btnCancel As New Windows.Forms.Button() With {
                .Text = "Cancel",
                .Left = 255,
                .Top = y,
                .DialogResult = Windows.Forms.DialogResult.Cancel,
                .Width = 75
            }

            Me.AcceptButton = btnOK : Me.CancelButton = btnCancel
            Me.Controls.AddRange({lbl1, nudOverlay, lbl2, nudOffsetUL, lbl3, nudCarpanDonati, chkDrawSections, btnOK, btnCancel})

            AddHandler btnOK.Click, AddressOf OnOKSave
        End Sub

        Private Sub OnOKSave(sender As Object, e As EventArgs)
            Settings.SetDouble("OverlayUpper", CDbl(nudOverlay.Value))
            Settings.SetDouble("OffsetUL", CDbl(nudOffsetUL.Value))
            Settings.SetDouble("CarpanDonati", CDbl(nudCarpanDonati.Value))
        End Sub

        Private Sub RebarParamsForm_DragOver(sender As Object, e As Windows.Forms.DragEventArgs) Handles Me.DragOver

        End Sub

        Private Sub RebarParamsForm_HelpRequested(sender As Object, hlpevent As Windows.Forms.HelpEventArgs) Handles Me.HelpRequested

        End Sub
    End Class

#End Region

    Private Delegate Function TryPlaceDelegate(
    xTarget As Double,
    yRow As Double,
    ByRef lastXOnRow As Double,
    rngL As Double,
    rngR As Double
) As Boolean

#Region "Section"

    ' ===== Section helpers =====
    ' 0-based index -> A, B, ... Z, AA, AB, ...
    Private Function SecLetter(ByVal idx As Integer) As String
        Dim s As String = ""
        Dim n As Integer = idx
        Do
            Dim r As Integer = n Mod 26
            s = ChrW(AscW("A"c) + r) & s
            n = (n \ 26) - 1
        Loop While n >= 0
        Return s
    End Function

    ' Draw the letter tag at a given XY
    Private Sub DrawSectionLetterTag(
    btr As BlockTableRecord, tr As Transaction,
    x As Double, y As Double, letter As String,
    Optional txtH As Double = 12.0,
    Optional lay As String = LYR_DIM,
    Optional col As Short = 3)

        Dim mt As New MText()
        mt.Location = New Point3d(x, y, 0)
        mt.TextHeight = txtH
        mt.Contents = letter
        mt.Attachment = AttachmentPoint.MiddleCenter
        mt.Layer = lay
        mt.ColorIndex = col

        ' Preserve your existing rotation (90 degrees)
        mt.Rotation = Math.PI / 2.0

        ' --- NEW: Apply the Romans Text Style ---
        mt.TextStyleId = GetRomansTextStyleId(tr, btr.Database)
        ' ----------------------------------------

        ' Assuming AppendOnLayer is your own existing helper sub
        AppendOnLayer(btr, tr, mt, lay, forceByLayer:=True)
    End Sub

    ' =================================================================
    ' HELPER: Finds the "Romans" style or creates it if it's missing
    ' =================================================================
    Private Function GetRomansTextStyleId(tr As Transaction, db As Database) As ObjectId
        Dim tsTable As TextStyleTable = CType(tr.GetObject(db.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
        Dim styleName As String = "Romans"

        ' 1. If it already exists, return its ID
        If tsTable.Has(styleName) Then
            Return tsTable(styleName)
        End If

        ' 2. If not, create it pointing to romans.shx
        tsTable.UpgradeOpen()
        Dim newStyle As New TextStyleTableRecord()
        newStyle.Name = styleName
        newStyle.FileName = "romans.shx" ' The standard AutoCAD font file

        Dim id As ObjectId = tsTable.Add(newStyle)
        tr.AddNewlyCreatedDBObject(newStyle, True)

        Return id
    End Function
#End Region


#Region "Section Command"

    <CommandMethod("DRAW_REBAR_SECTION")>
    Public Sub DrawRebarSection()
        Dim doc = AcApp.DocumentManager.MdiActiveDocument
        Dim db = doc.Database
        Dim ed = doc.Editor

        ' ---------- layout gaps (cm) ----------
        Const START_GAP_RIGHT_OF_LAST As Double = 200.0   ' big gap from last beam to first section
        Const GAP_BETWEEN_SECTIONS As Double = 150.0      ' big gap between adjacent sections

        ' 1) Select geometry
        Dim selOpts As New PromptSelectionOptions() With {.MessageForAdding = vbLf & "Select columns (vertical) and beams (S-BEAM lines): "}
        Dim selRes = ed.GetSelection(selOpts)
        If selRes.Status <> PromptStatus.OK Then Return
        Dim ids = selRes.Value.GetObjectIds()

        ' 2) Load Excel/CSV
        Dim ofo As New PromptOpenFileOptions(vbLf & "Select Excel/CSV with beam info (KNumber, Alt Donatı, Montaj, [Gövde]):")
        ofo.Filter = "Excel (*.xlsx;*.xls)|*.xlsx;*.xls|CSV (*.csv)|*.csv|All files (*.*)|*.*"
        Try
            Dim lastPath = Settings.GetString("LastDataPath", "")
            If Not String.IsNullOrEmpty(lastPath) Then
                Dim ip = Path.GetDirectoryName(lastPath)
                If Not String.IsNullOrEmpty(ip) AndAlso Directory.Exists(ip) Then ofo.InitialDirectory = ip
                ofo.InitialFileName = Path.GetFileName(lastPath)
            End If
        Catch
        End Try
        Dim ofr = ed.GetFileNameForOpen(ofo)
        If ofr.Status <> PromptStatus.OK Then Return
        Dim dataPath As String = ofr.StringResult : Settings.SetString("LastDataPath", dataPath)
        Dim ext = Path.GetExtension(dataPath).ToLowerInvariant()

        Dim excelMap As Dictionary(Of String, BeamExcelRow) = Nothing
        Dim ok As Boolean = False
        If ext = ".csv" Then ok = TryLoadCsvMap(dataPath, excelMap, ed) Else ok = TryLoadExcelMap(dataPath, excelMap, ed)
        If Not ok Then
            ed.WriteMessage(vbLf & "No usable beam info loaded; using defaults.")
            excelMap = New Dictionary(Of String, BeamExcelRow)(StringComparer.OrdinalIgnoreCase)
        End If

        ' 3) Collect geometry + labels
        Dim columnEdges As New List(Of Double)()
        Dim beamLines As New List(Of HLine)()
        CollectGeometry(ids, columnEdges, beamLines)
        If beamLines.Count = 0 Then ed.WriteMessage(vbLf & "No S-BEAM lines detected.") : Return

        Dim labels As New List(Of LabelPt)()
        CollectLabels(ids, labels)

        Dim uniqueEdges = UniqueSorted(columnEdges, 0.000001)
        Dim numColumns As Integer = uniqueEdges.Count \ 2
        Const EPS_Y As Double = 0.001, EPS_LEN As Double = 0.01

        ' global rightmost X of all selected beams (for placing the first section)
        Dim allXs As New List(Of Double)()
        For Each h In beamLines : allXs.Add(h.X0) : allXs.Add(h.X1) : Next
        Dim globalRight As Double = If(allXs.Count > 0, allXs.Max(), 0.0)

        ' Collect all sections first
        Dim jobs As New List(Of SectionJob)()
        ' Tuple: (topY, bottomY, wCm, hCm, info, spanRightForRefOnly)

        ' ---- WITH COLUMNS: one job per span between columns ----
        If numColumns >= 2 Then
            Dim asmRanges = SplitAssemblies(uniqueEdges)
            For Each r In asmRanges
                Dim c0 = r.Item1, c1 = r.Item2
                Dim asmLeft = uniqueEdges(2 * c0)
                Dim asmRight = uniqueEdges(2 * c1 + 1)

                Dim yValsAsm As New List(Of Double)()
                For Each h In beamLines
                    If h.X1 > asmLeft + EPS_LEN AndAlso h.X0 < asmRight - EPS_LEN Then yValsAsm.Add(h.Y)
                Next
                Dim uniqueYAsm = UniqueSorted(yValsAsm, EPS_Y)
                Dim groups = PairBeamLevels(uniqueYAsm)
                Dim nSpaces As Integer = (c1 - c0)
                For Each grp In groups
                    Dim topY = grp.Item1, bottomY = grp.Item2
                    For iSpace As Integer = 0 To nSpaces - 1
                        Dim leftColIdx As Integer = c0 + iSpace
                        Dim rightColIdx As Integer = c0 + iSpace + 1
                        Dim spanLeft As Double = uniqueEdges(2 * leftColIdx + 1)
                        Dim spanRight As Double = uniqueEdges(2 * rightColIdx)

                        Dim kKey = FindClosestKNumber(labels, spanLeft, spanRight, topY, bottomY)
                        Dim info As BeamExcelRow = Nothing
                        If Not excelMap.TryGetValue(kKey, info) Then
                            kKey = FindClosestKNumber(labels, asmLeft, asmRight, topY, bottomY)
                            If Not excelMap.TryGetValue(kKey, info) Then info = New BeamExcelRow()
                        End If

                        Dim wCm As Double, hCm As Double
                        If Not FindClosestBeamSize(labels, spanLeft, spanRight, topY, bottomY, wCm, hCm) Then
                            If Not FindClosestBeamSize(labels, asmLeft, asmRight, topY, bottomY, wCm, hCm) Then
                                wCm = 30.0 : hCm = Math.Abs(topY - bottomY)
                            End If
                        End If



                        jobs.Add(New SectionJob With {
    .TopY = topY,
    .BottomY = bottomY,
    .WCm = wCm,
    .HCm = hCm,
    .Info = info,
    .SpanLeft = spanLeft,
    .SpanRight = spanRight,
    .Side = BeamSide.LeftSide,
    .BeamName = kKey
})

                        ' Right-side section for this beam
                        jobs.Add(New SectionJob With {
    .TopY = topY,
    .BottomY = bottomY,
    .WCm = wCm,
    .HCm = hCm,
    .Info = info,
    .SpanLeft = spanLeft,
    .SpanRight = spanRight,
    .Side = BeamSide.RightSide,
    .BeamName = kKey
})
                    Next
                Next
            Next


        Else
            ' ---- NO COLUMNS: split spans from S-BEAM coverage ----
            Dim yVals As New List(Of Double)()
            For Each h In beamLines : yVals.Add(h.Y) : Next
            Dim uniqueY = UniqueSorted(yVals, EPS_Y)
            Dim yPairs = PairBeamLevels(uniqueY)

            For Each grp In yPairs
                Dim topY = grp.Item1, bottomY = grp.Item2
                Dim spans = BuildSpansNoColumns(beamLines, topY, bottomY, EPS_Y, EPS_LEN)

                If spans.Count = 0 Then
                    Dim xs As New List(Of Double)()
                    For Each h In beamLines
                        If Math.Abs(h.Y - topY) < EPS_Y OrElse Math.Abs(h.Y - bottomY) < EPS_Y Then
                            xs.Add(h.X0) : xs.Add(h.X1)
                        End If
                    Next
                    If xs.Count = 0 Then Continue For
                    Dim minX = xs.Min(), maxX = xs.Max()

                    Dim kKey = FindClosestKNumber(labels, minX, maxX, topY, bottomY)
                    Dim info As BeamExcelRow = Nothing
                    If Not excelMap.TryGetValue(kKey, info) Then info = New BeamExcelRow()

                    Dim wCm As Double, hCm As Double
                    If Not FindClosestBeamSize(labels, minX, maxX, topY, bottomY, wCm, hCm) Then
                        wCm = 30.0 : hCm = Math.Abs(topY - bottomY)
                    End If

                    jobs.Add(New SectionJob With {
    .TopY = topY,
    .BottomY = bottomY,
    .WCm = wCm,
    .HCm = hCm,
    .Info = info,
   .SpanRight = maxX,
    .Side = BeamSide.LeftSide,
    .BeamName = kKey
})

                    ' Right-side section for this beam
                    jobs.Add(New SectionJob With {
    .TopY = topY,
    .BottomY = bottomY,
    .WCm = wCm,
    .HCm = hCm,
    .Info = info,
   .SpanRight = maxX,
    .Side = BeamSide.RightSide,
    .BeamName = kKey
})
                Else
                    For Each sp In spans
                        Dim spanLeft = sp.Item1, spanRight = sp.Item2
                        Dim kKey = FindClosestKNumber(labels, spanLeft, spanRight, topY, bottomY)
                        Dim info As BeamExcelRow = Nothing
                        If Not excelMap.TryGetValue(kKey, info) Then info = New BeamExcelRow()

                        Dim wCm As Double, hCm As Double
                        If Not FindClosestBeamSize(labels, spanLeft, spanRight, topY, bottomY, wCm, hCm) Then
                            wCm = 30.0 : hCm = Math.Abs(topY - bottomY)
                        End If

                        jobs.Add(New SectionJob With {
    .TopY = topY,
    .BottomY = bottomY,
    .WCm = wCm,
    .HCm = hCm,
    .Info = info,
    .SpanRight = spanRight,
    .Side = BeamSide.LeftSide,
    .BeamName = kKey
})

                        ' Right-side section for this beam
                        jobs.Add(New SectionJob With {
    .TopY = topY,
    .BottomY = bottomY,
    .WCm = wCm,
    .HCm = hCm,
    .Info = info,
    .SpanRight = spanRight,
    .Side = BeamSide.RightSide,
    .BeamName = kKey
})
                    Next
                End If
            Next
        End If



        If jobs.Count = 0 Then ed.WriteMessage(vbLf & "No spans found for sections.") : Return

        ' 4) Draw: side-by-side to the right of the last beam
        Using tr = db.TransactionManager.StartTransaction()
            Dim btr = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
            EnsureAllLayers(db, tr)

            Dim cursorX As Double = globalRight + START_GAP_RIGHT_OF_LAST
            For Each jb In jobs
                DrawRebarSectionForGroup(
            btr, tr,
            jb.SpanRight,          ' spanRightForRefOnly
            jb.TopY, jb.BottomY,   ' top & bottom
            jb.WCm, jb.HCm,        ' section size
            jb.Info,               ' bar info
            20.0,                  ' addlBelowUpper (same as before)
            baseXOverride:=cursorX,
            side:=jb.Side,         ' <<< pass side
            beamName:=jb.BeamName  ' <<< pass K name to print
        )


                cursorX += jb.WCm + GAP_BETWEEN_SECTIONS

            Next

            tr.Commit()
        End Using
    End Sub


#End Region

    ' =====================================================================
    ' Draw sections to the right of the plan using existing data/selection
    ' =====================================================================
    Private Sub DrawSectionsRightOfPlan(
    db As Database,
    ed As Editor,
    excelMap As Dictionary(Of String, BeamExcelRow),
    columnEdges As List(Of Double),
    beamLines As List(Of HLine),
    labels As List(Of LabelPt),
    Optional START_GAP_RIGHT_OF_LAST As Double = 200.0,
    Optional GAP_BETWEEN_SECTIONS As Double = 150.0,
    Optional coverParam As Double = 3.0
)
        Const EPS_Y As Double = 0.001
        Const EPS_LEN As Double = 0.01

        Dim cv As Double = coverParam  ' single local cover for this routine

        If beamLines Is Nothing OrElse beamLines.Count = 0 Then
            ed.WriteMessage(vbLf & "No S-BEAM lines detected for sections.")
            Return
        End If

        Dim uniqueEdges = UniqueSorted(columnEdges, 0.000001)
        Dim numColumns As Integer = uniqueEdges.Count \ 2

        ' global rightmost X among selected beams
        Dim allXs As New List(Of Double)()
        For Each h In beamLines : allXs.Add(h.X0) : allXs.Add(h.X1) : Next
        Dim globalRight As Double = If(allXs.Count > 0, allXs.Max(), 0.0)

        Dim jobs As New List(Of SectionJob)()

        If numColumns >= 2 Then
            ' ===== With columns =====
            Dim asmRanges = SplitAssemblies(uniqueEdges)
            For Each r In asmRanges
                Dim c0 = r.Item1, c1 = r.Item2
                Dim asmLeft = uniqueEdges(2 * c0)
                Dim asmRight = uniqueEdges(2 * c1 + 1)

                ' Y groups in this assembly
                Dim yValsAsm As New List(Of Double)()
                For Each h In beamLines
                    If h.X1 > asmLeft + EPS_LEN AndAlso h.X0 < asmRight - EPS_LEN Then yValsAsm.Add(h.Y)
                Next
                Dim uniqueYAsm = UniqueSorted(yValsAsm, EPS_Y)
                Dim groups = PairBeamLevels(uniqueYAsm)
                Dim nSpaces As Integer = (c1 - c0)

                For Each grp In groups
                    Dim topY = grp.Item1, bottomY = grp.Item2

                    For iSpace As Integer = 0 To nSpaces - 1
                        Dim leftColIdx As Integer = c0 + iSpace
                        Dim rightColIdx As Integer = c0 + iSpace + 1

                        ' span faces (inside the two columns)
                        Dim spanLeft As Double = uniqueEdges(2 * leftColIdx + 1)     ' right face of left column
                        Dim spanRight As Double = uniqueEdges(2 * rightColIdx)      ' left face of right column

                        ' K and size
                        Dim kKey = FindClosestKNumber(labels, spanLeft, spanRight, topY, bottomY)
                        Dim info As BeamExcelRow = Nothing
                        If Not excelMap.TryGetValue(kKey, info) Then
                            kKey = FindClosestKNumber(labels, asmLeft, asmRight, topY, bottomY)
                            If Not excelMap.TryGetValue(kKey, info) Then info = New BeamExcelRow()
                        End If

                        Dim wCm As Double, hCm As Double
                        If Not FindClosestBeamSize(labels, spanLeft, spanRight, topY, bottomY, wCm, hCm) Then
                            If Not FindClosestBeamSize(labels, asmLeft, asmRight, topY, bottomY, wCm, hCm) Then
                                wCm = 30.0 : hCm = Math.Abs(topY - bottomY)
                            End If
                        End If

                        ' Tag X equals ADDL duplicate ends:
                        '   Left side  -> startX_UL = spanLeft  + cover
                        '   Right side -> endX_UL   = spanRight - cover
                        Dim tagXLeft As Double = spanLeft + cv
                        Dim tagXRight As Double = spanRight - cv

                        jobs.Add(New SectionJob With {
                        .TopY = topY, .BottomY = bottomY,
                        .WCm = wCm, .HCm = hCm, .Info = info,
                        .SpanLeft = spanLeft, .SpanRight = spanRight, .Side = BeamSide.LeftSide, .BeamName = kKey,
                        .TagX = tagXLeft
                    })
                        jobs.Add(New SectionJob With {
                        .TopY = topY, .BottomY = bottomY,
                        .WCm = wCm, .HCm = hCm, .Info = info,
                        .SpanLeft = spanLeft, .SpanRight = spanRight, .Side = BeamSide.RightSide, .BeamName = kKey,
                        .TagX = tagXRight
                    })
                    Next
                Next
            Next
        Else
            ' ===== No columns: infer spans =====
            Dim yVals As New List(Of Double)()
            For Each h In beamLines : yVals.Add(h.Y) : Next
            Dim uniqueY = UniqueSorted(yVals, EPS_Y)
            Dim yPairs = PairBeamLevels(uniqueY)

            For Each grp In yPairs
                Dim topY = grp.Item1, bottomY = grp.Item2
                Dim spans = BuildSpansNoColumns(beamLines, topY, bottomY, EPS_Y, EPS_LEN)

                If spans.Count = 0 Then
                    ' fallback: one span between min/max X on these two lines
                    Dim xs As New List(Of Double)()
                    For Each h In beamLines
                        If Math.Abs(h.Y - topY) < EPS_Y OrElse Math.Abs(h.Y - bottomY) < EPS_Y Then
                            xs.Add(h.X0) : xs.Add(h.X1)
                        End If
                    Next
                    If xs.Count = 0 Then Continue For
                    Dim minX = xs.Min(), maxX = xs.Max()

                    Dim kKey = FindClosestKNumber(labels, minX, maxX, topY, bottomY)
                    Dim info As BeamExcelRow = Nothing
                    If Not excelMap.TryGetValue(kKey, info) Then info = New BeamExcelRow()

                    Dim wCm As Double, hCm As Double
                    If Not FindClosestBeamSize(labels, minX, maxX, topY, bottomY, wCm, hCm) Then
                        wCm = 30.0 : hCm = Math.Abs(topY - bottomY)
                    End If

                    jobs.Add(New SectionJob With {
                    .TopY = topY, .BottomY = bottomY, .WCm = wCm, .HCm = hCm, .Info = info,
                    .SpanRight = maxX, .Side = BeamSide.LeftSide, .BeamName = kKey,
                    .TagX = minX + cv
                })
                    jobs.Add(New SectionJob With {
                    .TopY = topY, .BottomY = bottomY, .WCm = wCm, .HCm = hCm, .Info = info,
                    .SpanRight = maxX, .Side = BeamSide.RightSide, .BeamName = kKey,
                    .TagX = maxX - cv
                })
                Else
                    For Each sp In spans
                        Dim spanLeft = sp.Item1, spanRight = sp.Item2
                        Dim kKey = FindClosestKNumber(labels, spanLeft, spanRight, topY, bottomY)
                        Dim info As BeamExcelRow = Nothing
                        If Not excelMap.TryGetValue(kKey, info) Then info = New BeamExcelRow()

                        Dim wCm As Double, hCm As Double
                        If Not FindClosestBeamSize(labels, spanLeft, spanRight, topY, bottomY, wCm, hCm) Then
                            wCm = 30.0 : hCm = Math.Abs(topY - bottomY)
                        End If

                        jobs.Add(New SectionJob With {
                        .TopY = topY, .BottomY = bottomY, .WCm = wCm, .HCm = hCm, .Info = info,
                        .SpanRight = spanRight, .Side = BeamSide.LeftSide, .BeamName = kKey,
                        .TagX = spanLeft + cv
                    })
                        jobs.Add(New SectionJob With {
                        .TopY = topY, .BottomY = bottomY, .WCm = wCm, .HCm = hCm, .Info = info,
                        .SpanRight = spanRight, .Side = BeamSide.RightSide, .BeamName = kKey,
                        .TagX = spanRight - cv
                    })
                    Next
                End If
            Next
        End If

        If jobs.Count = 0 Then
            ed.WriteMessage(vbLf & "No spans found for sections.")
            Return
        End If

        Using tr = db.TransactionManager.StartTransaction()
            Dim btr = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
            EnsureAllLayers(db, tr)

            Dim cursorX As Double = globalRight + START_GAP_RIGHT_OF_LAST
            Dim secCounter As Integer = 0

            For Each jb In jobs
                ' ==== PLAN letters ====
                Const ScaleFactorFlag As Double = 0.6
                Const LETTER_INSET As Double = 100.0
                Dim baseX As Double = If(jb.TagX <> 0, jb.TagX, jb.SpanRight)
                Dim xCut As Double = If(jb.Side = BeamSide.LeftSide, baseX + LETTER_INSET, baseX - LETTER_INSET)
                Dim letter As String = SecLetter(secCounter)
                Const TAG_DY_TOP As Double = 35.0 * ScaleFactorFlag
                Const TAG_DY_BOT As Double = 35.0 * ScaleFactorFlag
                DrawSectionLetterTag(btr, tr, xCut, jb.TopY + TAG_DY_TOP, letter)
                DrawSectionLetterTag(btr, tr, xCut, jb.BottomY - TAG_DY_BOT, letter)

                Const FLAG_DX As Double = 10.0 * ScaleFactorFlag
                Const POLE As Double = 28.0 * ScaleFactorFlag
                Const TRIL As Double = 18.0 * ScaleFactorFlag
                Const TRIH As Double = 10.0 * ScaleFactorFlag

                Dim dir As Integer = -1                               ' always Left
                Dim poleTopX As Double = xCut + 3 * FLAG_DX           ' place just right of the letter

                ' TOP
                PlanTriangleFlag.DrawTriangleFlag(
                btr, tr, poleTopX, jb.TopY + TAG_DY_TOP, dir,
                poleLen:=POLE, triLen:=TRIL, triHt:=TRIH,
                rotate180:=False)

                ' BOTTOM flag — mirror across the X-axis so the triangle is below the pole
                PlanTriangleFlag.DrawTriangleFlag(
                btr, tr, poleTopX, jb.BottomY - TAG_DY_BOT, -dir,
                poleLen:=POLE, triLen:=TRIL, triHt:=TRIH,
                rotate180:=True)

                ' ==== SECTION drawing ====

                'Kesit
                DrawRebarSectionForGroup(
                btr, tr, jb.SpanRight, jb.TopY, jb.BottomY,
                jb.WCm, jb.HCm, jb.Info, 20.0,
                baseXOverride:=cursorX, side:=jb.Side, beamName:=jb.BeamName,
                secLetter:=letter, coverParam:=cv)

                'Çiroz

                Dim tc = PosTiebar.ComputeFromSectionInputs(
            baseX:=cursorX,
            secW:=Math.Max(10.0, jb.WCm),
            topY:=jb.TopY, bottomY:=jb.BottomY,
            secHcm:=jb.HCm,
            coverParam:=cv, _                          ' same cover used above
            rWeb:=Math.Max(0.6, If(jb.Info IsNot Nothing, jb.Info.PhiWeb, 0) / 2.0R),
            tieBarsSpec:=If(jb.Info IsNot Nothing, jb.Info.TieBarsSpec, ""))

                ' Etriye 

                Dim stirQty As Integer = 0
                Dim stirPhiMm As Integer = 0
                Dim stirPitch As Integer = 0

                Dim stirSpec As String = If(jb.Info IsNot Nothing, jb.Info.TieBarsSpec, "")
                If Not String.IsNullOrWhiteSpace(stirSpec) Then
                    Dim qty As Integer
                    Dim phiD As Double
                    Dim pitchD As Double

                    If TieBarsHelper.TryParseTieSpec(stirSpec, qty, phiD, pitchD) Then
                        stirQty = qty
                        stirPhiMm = CInt(Math.Round(phiD))
                        stirPitch = CInt(Math.Round(pitchD))
                    End If
                End If

                ' === draw separate stirrup UNDER the section (matching lengths) ===
                Dim gapRight As Double = 60.0                    ' same as SectionsModule
                Dim secW As Double = jb.WCm
                Dim secH As Double = jb.HCm
                Dim secBaseX As Double = cursorX
                Dim secBaseY As Double = 0.5 * (jb.TopY + jb.BottomY) - jb.HCm / 2.0

                Dim gapBelow As Double = 400.0                   ' use variable so we can reuse
                Dim skewRight As Double = 12.0                   ' same as skewRightCm

                StirrupModule.DrawBelow(
    btr, tr,
    secBaseX, secBaseY,
    Math.Max(10.0, jb.WCm), jb.HCm,
    coverCm:=cv,                  ' same cover as section
    gapBelowCm:=gapBelow,         ' fixed 400 cm below
    gapCornerCm:=6.0,             ' top-right separation
    skewRightCm:=skewRight,       ' lean right
    leaderLenX:=10.0,             ' inward leaders
    leaderLenY:=8.0,
    hookText:="12",
    textH:=3.333)                 ' small skew for readability

                If stirQty > 0 AndAlso stirPhiMm > 0 AndAlso stirPitch > 0 Then
                    ' Recompute stirrup geometry exactly like StirrupModule
                    Dim stirrupW As Double = Math.Max(1.0, secW - 2 * cv)
                    Dim stirrupH As Double = Math.Max(1.0, secH - 2 * cv)

                    ' reuse secBaseX/secBaseY instead of redefining baseX/baseY
                    Dim stirrupBaseX As Double = secBaseX
                    Dim stirrupBaseY As Double = secBaseY - gapBelow - stirrupH

                    ' Approximate total length: 2*(W+H) + 2*hook
                    Dim hookLen As Double = 0.0
                    Double.TryParse(
        "12",
        Globalization.NumberStyles.Float,
        Globalization.CultureInfo.InvariantCulture,
        hookLen)

                    Dim totalLenCm As Double = 2.0 * (stirrupW + stirrupH)
                    If hookLen > 0 Then
                        totalLenCm += 2.0 * hookLen
                    End If
                    Dim totalLen As Integer = CInt(Math.Round(totalLenCm))

                    ' POS point under the stirrup
                    Dim midX As Double = stirrupBaseX + (stirrupW + skewRight) / 2.0 - 20.0
                    Dim midY As Double = stirrupBaseY - 20.0
                    Dim posMid As New Point3d(midX, midY, 0)

                    PosManager.Manager.AddStirrupPos(
                        mid:=posMid,
                        phiMm:=stirPhiMm,
                        totalLen:=totalLen,
                        adet:=stirQty,
                        aralik:=stirPitch)

                    'Kiriş altı etriye pozu ( metraj için)
                    If (jb.Side = BeamSide.LeftSide) AndAlso Not Double.IsNaN(_planBeamDimlineY) Then
                        Dim beamMidX As Double = 0.5R * (jb.SpanLeft + jb.SpanRight)
                        Dim posMidDim As New Point3d(beamMidX, _planBeamDimlineY + 22.0R, 0)
                        Dim beamLen As Double = jb.SpanRight - jb.SpanLeft

                        PosManager.Manager.AddStirrupPosMetraj(
                            mid:=posMidDim,
                            phiMm:=stirPhiMm,
                            totalLen:=totalLen,
                            beamLen:=beamLen,
                            aralik:=stirPitch)
                    End If
                End If


                ' === B) Draw the separate tie 200 cm below; text white; your PosTiebar settings apply ===
                PosTiebar.DrawSeparateTie(btr, tr, tc, gapDown:=200.0, slantDeg:=45.0)

                ' === POS for tie (Çiroz) under its separate symbol ===
                Dim dzSpec As String = If(jb.Info IsNot Nothing, jb.Info.TieBarsSpec, "")
                If Not String.IsNullOrWhiteSpace(dzSpec) Then
                    Dim qty As Integer, phiD As Double, pitchD As Double

                    If TieBarsHelper.TryParseTieSpec(dzSpec, qty, phiD, pitchD) AndAlso qty > 0 Then
                        Dim phiMm As Integer = CInt(Math.Round(phiD))
                        Dim aralik As Integer = CInt(Math.Round(pitchD))

                        ' Total tie length from the computed geometry
                        Dim tieBoy As Integer = CInt(Math.Round(tc.Leg1 + tc.StirrupH + tc.Leg2))

                        ' Place POS marker centered under the drawn tie bar
                        Const GAP_DOWN As Double = 200.0    ' same as used in DrawSeparateTie
                        Dim xMid As Double = 0.5 * (tc.XLeft + tc.XRight)
                        Dim baseY As Double = tc.YLower - GAP_DOWN
                        Dim posY As Double = baseY - 30.0   ' 30 units below the tie base so it's clearly under

                        Dim posMid As New Point3d(xMid, posY, 0)

                        PosManager.Manager.AddTiebarPos(
                                    mid:=posMid,
                                    phiMm:=phiMm,
                                    totalLen:=tieBoy,
                                    adet:=qty,
                                    aralik:=aralik)

                        Dim beamLen As Double = jb.SpanRight - jb.SpanLeft
                        'Kiriş altı etriye pozu metraj için)
                        If (jb.Side = BeamSide.LeftSide) AndAlso Not Double.IsNaN(_planBeamDimlineY) Then
                            Dim beamMidX As Double = 0.5R * (jb.SpanLeft + jb.SpanRight)
                            Dim posTieDim As New Point3d(beamMidX, _planBeamDimlineY + 10.0R, 0)

                            PosManager.Manager.AddTiebarPosMetraj(
                                mid:=posTieDim,
                                phiMm:=phiMm,
                                totalLen:=tieBoy,
                                beamLen:=beamLen,
                                aralik:=aralik)
                        End If
                    End If
                End If
                Dim usedWidth As Double = Math.Max(10.0, jb.WCm)   ' mirror the min width used inside drawer
                cursorX += usedWidth + GAP_BETWEEN_SECTIONS
                secCounter += 1
                ' ==== Standalone tie bar directly under this section ====
                ' base of the drawn section sits at X = cursorX; compute baseY like in sections:
                Dim centerY As Double = 0.5 * (jb.TopY + jb.BottomY)
                Dim baseYSection As Double = centerY - jb.HCm / 2.0




            Next
            PosManager.Manager.Commit(db, tr, btr, markerRadius:=5.5)

            tr.Commit()
        End Using
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    ' ========= BAR PACKING (section) =========

    Private Class BarSpec
        Public Qty As Integer
        Public PhiCm As Double
        Public Label As String        ' e.g. "2Ø16"
        Public Sub New(q As Integer, phi As Double, lbl As String)
            Qty = q : PhiCm = phi : Label = lbl
        End Sub
    End Class

    Private Class PlacedBar
        Public Center As Point3d
        Public Radius As Double
        Public Label As String
    End Class

End Class
