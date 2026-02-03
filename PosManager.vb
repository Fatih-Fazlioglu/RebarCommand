' ===========================================
' PosBlockGenerator.vb
' ===========================================
Option Strict On
Option Infer On

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Public Module PosManager

    ' =======================
    ' Public Facade (Manager)
    ' =======================
    Public NotInheritable Class Manager
        Private Sub New()
        End Sub

        Private Shared _nextPosNumber As Integer = 1
        Private Shared ReadOnly _keyToPos As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Private Shared ReadOnly _keyToCount As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Private Shared ReadOnly _instances As New List(Of Instance)()
        Private Shared _carpanDonati As Integer = 0

        Public Shared Property CarpanDonati As Integer
            Get
                Return _carpanDonati
            End Get
            Set(value As Integer)
                _carpanDonati = Math.Max(0, CInt(value))   ' 0 => fall back to auto count
            End Set
        End Property

        ' =========================================================
        ' 1. MAIN REBAR METHOD (Modified to accept linkedEntId)
        ' =========================================================
        Public Shared Sub AddDuplicateRebar(
            mid As Point3d,
            phiMm As Integer,
            totalLen As Integer,
            yer As String,
            adet As Integer,
            miterCount As Integer,
            miter1 As Integer,
            straightPart As Integer,
            miter2 As Integer,
            Optional linkedEntId As ObjectId = Nothing)

            Dim key As String = BuildSignature(phiMm, totalLen, yer, miterCount, miter1, straightPart, miter2)

            Dim posNo As Integer
            If Not _keyToPos.TryGetValue(key, posNo) Then
                posNo = _nextPosNumber
                _keyToPos(key) = posNo
                _nextPosNumber += 1
            End If

            If _keyToCount.ContainsKey(key) Then _keyToCount(key) += 1 Else _keyToCount(key) = 1

            _instances.Add(New Instance With {
                .Key = key,
                .PosNo = posNo,
                .Mid = mid,
                .PhiMm = phiMm,
                .TotalLen = totalLen,
                .Yer = yer,
                .Adet = adet.ToString(),
                .MiterCount = Math.Max(0, Math.Min(2, miterCount)),
                .Miter1 = miter1,
                .StraightPart = straightPart,
                .Miter2 = miter2,
                .IsRebar = True,
                .LinkedEntityId = linkedEntId ' <--- Stored for Field creation
            })
        End Sub

        ' =========================================================
        ' 1b. REBAR METHOD with String ADET (for nxm format)
        ' =========================================================
        Public Shared Sub AddDuplicateRebarWithAdetText(
            mid As Point3d,
            phiMm As Integer,
            totalLen As Integer,
            yer As String,
            adetText As String,
            miterCount As Integer,
            miter1 As Integer,
            straightPart As Integer,
            miter2 As Integer,
            Optional linkedEntId As ObjectId = Nothing)

            Dim key As String = BuildSignature(phiMm, totalLen, yer, miterCount, miter1, straightPart, miter2)

            Dim posNo As Integer
            If Not _keyToPos.TryGetValue(key, posNo) Then
                posNo = _nextPosNumber
                _keyToPos(key) = posNo
                _nextPosNumber += 1
            End If

            If _keyToCount.ContainsKey(key) Then _keyToCount(key) += 1 Else _keyToCount(key) = 1

            _instances.Add(New Instance With {
                .Key = key,
                .PosNo = posNo,
                .Mid = mid,
                .PhiMm = phiMm,
                .TotalLen = totalLen,
                .Yer = yer,
                .Adet = adetText,
                .MiterCount = Math.Max(0, Math.Min(2, miterCount)),
                .Miter1 = miter1,
                .StraightPart = straightPart,
                .Miter2 = miter2,
                .IsRebar = True,
                .LinkedEntityId = linkedEntId
            })
        End Sub

        ' =========================================================
        ' 2. RESTORED METHOD: Tiebar (Çiroz)
        ' =========================================================
        Public Shared Sub AddTiebarPos(
            mid As Point3d,
            phiMm As Integer,
            totalLen As Integer,
            adet As Integer,
            aralik As Integer)

            If phiMm <= 0 OrElse totalLen <= 0 OrElse adet <= 0 OrElse aralik <= 0 Then Exit Sub

            Dim yer As String = "Crz"
            Dim key As String = $"tie|φ{phiMm}|L{totalLen}|yer:{yer.ToLowerInvariant()}|arl:{aralik}"

            Dim posNo As Integer
            If Not _keyToPos.TryGetValue(key, posNo) Then
                posNo = _nextPosNumber
                _keyToPos(key) = posNo
                _nextPosNumber += 1
            End If

            If _keyToCount.ContainsKey(key) Then
                _keyToCount(key) += adet
            Else
                _keyToCount(key) = adet
            End If

            _instances.Add(New Instance With {
                .Key = key,
                .PosNo = posNo,
                .Mid = mid,
                .PhiMm = phiMm,
                .TotalLen = totalLen,
                .Yer = yer,
                .Adet = adet.ToString(),
                .MiterCount = 0,
                .Miter1 = 0,
                .StraightPart = 0,
                .Miter2 = 0,
                .IsTie = True,
                .Aralik = aralik
            })
        End Sub

        ' =========================================================
        ' 3. RESTORED METHOD: Tiebar Metraj
        ' =========================================================
        Public Shared Sub AddTiebarPosMetraj(
            mid As Point3d,
            phiMm As Integer,
            totalLen As Integer,
            beamLen As Double,
            aralik As Integer)

            If phiMm <= 0 OrElse totalLen <= 0 OrElse beamLen <= 0 OrElse aralik <= 0 Then Exit Sub

            Dim adetCalc As Integer = CInt(Math.Ceiling(beamLen / aralik))
            Dim yer As String = "Crz"
            Dim key As String = $"tie|φ{phiMm}|L{totalLen}|yer:{yer.ToLowerInvariant()}|arl:{aralik}"

            Dim posNo As Integer
            If Not _keyToPos.TryGetValue(key, posNo) Then
                posNo = _nextPosNumber
                _keyToPos(key) = posNo
                _nextPosNumber += 1
            End If

            If _keyToCount.ContainsKey(key) Then
                _keyToCount(key) += adetCalc
            Else
                _keyToCount(key) = adetCalc
            End If

            _instances.Add(New Instance With {
                .Key = key,
                .PosNo = posNo,
                .Mid = mid,
                .PhiMm = phiMm,
                .TotalLen = totalLen,
                .Yer = yer,
                .Adet = adetCalc.ToString(),
                .MiterCount = 0,
                .Miter1 = 0,
                .StraightPart = 0,
                .Miter2 = 0,
                .IsTie = True,
                .IsForMetraj = True,
                .Aralik = aralik
            })
        End Sub

        ' =========================================================
        ' 4. RESTORED METHOD: Stirrup Metraj (Etriye Metraj)
        ' =========================================================
        Public Shared Sub AddStirrupPosMetraj(
            mid As Point3d,
            phiMm As Integer,
            totalLen As Integer,
            beamLen As Double,
            aralik As Integer)

            If phiMm <= 0 OrElse totalLen <= 0 OrElse beamLen <= 0 OrElse aralik <= 0 Then Exit Sub

            Dim adetCalc As Integer = CInt(Math.Ceiling(beamLen / aralik))
            Dim yer As String = "Etr"
            Dim key As String = $"stirrup|φ{phiMm}|L{totalLen}|yer:{yer.ToLowerInvariant()}|arl:{aralik}"

            Dim posNo As Integer
            If Not _keyToPos.TryGetValue(key, posNo) Then
                posNo = _nextPosNumber
                _keyToPos(key) = posNo
                _nextPosNumber += 1
            End If

            If _keyToCount.ContainsKey(key) Then
                _keyToCount(key) += adetCalc
            Else
                _keyToCount(key) = adetCalc
            End If

            _instances.Add(New Instance With {
                .Key = key,
                .PosNo = posNo,
                .Mid = mid,
                .PhiMm = phiMm,
                .TotalLen = totalLen,
                .Yer = yer,
                .Adet = adetCalc.ToString(),
                .MiterCount = 0,
                .Miter1 = 0,
                .StraightPart = totalLen,
                .Miter2 = 0,
                .IsTie = False,
                .IsForMetraj = True,
                .Aralik = aralik,
                .IsStirrup = True
            })
        End Sub

        ' =========================================================
        ' 5. RESTORED METHOD: Stirrup Normal (Etriye)
        ' =========================================================
        Public Shared Sub AddStirrupPos(
            mid As Point3d,
            phiMm As Integer,
            totalLen As Integer,
            adet As Integer,
            aralik As Integer)

            If phiMm <= 0 OrElse totalLen <= 0 OrElse adet <= 0 OrElse aralik <= 0 Then Exit Sub

            Dim yer As String = "Etr"
            Dim key As String = $"stirrup|φ{phiMm}|L{totalLen}|yer:{yer.ToLowerInvariant()}|arl:{aralik}"

            Dim posNo As Integer
            If Not _keyToPos.TryGetValue(key, posNo) Then
                posNo = _nextPosNumber
                _keyToPos(key) = posNo
                _nextPosNumber += 1
            End If

            If _keyToCount.ContainsKey(key) Then
                _keyToCount(key) += adet
            Else
                _keyToCount(key) = adet
            End If

            _instances.Add(New Instance With {
                .Key = key,
                .PosNo = posNo,
                .Mid = mid,
                .PhiMm = phiMm,
                .TotalLen = totalLen,
                .Yer = yer,
                .Adet = adet.ToString(),
                .MiterCount = 0,
                .Miter1 = 0,
                .StraightPart = totalLen,
                .Miter2 = 0,
                .IsTie = False,
                .Aralik = aralik,
                .IsStirrup = True
            })
        End Sub

        ' =========================================================
        ' COMMIT: Handles Block Insert and Field Creation
        ' =========================================================
        Public Shared Sub Commit(db As Database, tr As Transaction, btr As BlockTableRecord,
                                 Optional markerRadius As Double = 40.0)
            EnsurePozLayer(db, tr)
            EnsurePozTextStyle(db, tr)
            Dim blkId = EnsurePozBlock(db, tr, markerRadius)

            For Each inst In _instances
                Dim isTie As Boolean = inst.IsTie
                Dim isStirrup As Boolean = inst.IsStirrup
                Dim isForMetraj As Boolean = inst.IsForMetraj
                Dim carpan As Integer

                ' Logic for Bukum Boy
                Dim bukumBoy As String = ""
                Dim parts As New List(Of String)()
                Dim hasM1 As Boolean = inst.Miter1 > 0
                Dim hasM2 As Boolean = inst.Miter2 > 0
                Dim hasStraight As Boolean = inst.StraightPart > 0

                Select Case inst.MiterCount
                    Case 0
                        bukumBoy = ""
                    Case 1
                        If hasM1 Then
                            parts.Add($"{inst.Miter1}")
                        ElseIf hasM2 Then
                            parts.Add($"{inst.Miter2}")
                        End If
                        If hasStraight Then parts.Add($"{inst.StraightPart}")
                        bukumBoy = String.Join(" ", parts)
                    Case 2
                        If hasM1 Then parts.Add($"{inst.Miter1}")
                        If hasStraight Then parts.Add($"{inst.StraightPart}")
                        If hasM2 Then parts.Add($"{inst.Miter2}")
                        bukumBoy = String.Join(" ", parts)
                    Case Else
                        If hasM1 Then parts.Add($"{inst.Miter1}")
                        If hasStraight Then parts.Add($"{inst.StraightPart}")
                        If hasM2 Then parts.Add($"{inst.Miter2}")
                        bukumBoy = String.Join(" ", parts)
                End Select

                If (isTie OrElse isStirrup) AndAlso Not isForMetraj Then
                    carpan = 0
                Else
                    carpan = CarpanDonati
                End If

                ' Insert block
                Dim br As New BlockReference(inst.Mid, blkId)
                br.Layer = LYR_POZ
                btr.AppendEntity(br) : tr.AddNewlyCreatedDBObject(br, True)

                Dim bd = CType(tr.GetObject(blkId, OpenMode.ForRead), BlockTableRecord)
                For Each id In bd
                    Dim dbo = TryCast(tr.GetObject(id, OpenMode.ForRead), DBObject)
                    Dim ad = TryCast(dbo, AttributeDefinition)
                    If ad Is Nothing OrElse ad.Constant Then Continue For

                    Dim ar As New AttributeReference()
                    ar.SetAttributeFromBlock(ad, br.BlockTransform)
                    ar.Layer = LYR_POZ

                    ' Positioning Logic
                    Dim H As Double = Math.Max(1.0, ar.Height)
                    Dim extraOut As Double = Math.Max(2.0, 0.9 * H)
                    Dim extraBetween As Double = Math.Max(3.0, 0.8 * H)
                    Dim baseX As Double = br.Position.X + extraOut

                    ' Position adjustment: Etriye/Ciroz Normal
                    If (inst.IsTie OrElse inst.IsStirrup) And Not inst.IsForMetraj Then
                        If ad.Tag.Equals("ADET", StringComparison.OrdinalIgnoreCase) Then
                            ar.Invisible = True
                        End If
                        Select Case ad.Tag.ToUpperInvariant()
                            Case "NO" : ar.Position = New Point3d(ar.Position.X, br.Position.Y, ar.Position.Z)
                            Case "CAP" : ar.Position = New Point3d(ar.Position.X + (extraOut + extraBetween), br.Position.Y, ar.Position.Z)
                            Case "ARALIK" : ar.Position = New Point3d(ar.Position.X + (extraOut + extraBetween * 2.9), br.Position.Y, ar.Position.Z)
                            Case "YER" : ar.Position = New Point3d(ar.Position.X + (extraOut + extraBetween * 5.5), br.Position.Y, ar.Position.Z)
                            Case "BOY" : ar.Position = New Point3d(ar.Position.X + (extraOut + 8 * extraBetween), br.Position.Y, ar.Position.Z)
                        End Select
                    End If

                    ' Position adjustment: Metraj
                    If inst.IsForMetraj Then
                        Select Case ad.Tag.ToUpperInvariant()
                            Case "NO" : ar.Position = New Point3d(br.Position.X, br.Position.Y, ar.Position.Z)
                            Case "ADET" : ar.Position = New Point3d(baseX + extraOut, br.Position.Y, ar.Position.Z)
                            Case "CAP" : ar.Position = New Point3d(baseX + 3.0 * extraBetween, br.Position.Y, ar.Position.Z)
                            Case "ARALIK" : ar.Position = New Point3d(baseX + extraOut + 4.0 * extraBetween, br.Position.Y, ar.Position.Z)
                            Case "YER" : ar.Position = New Point3d(baseX + extraOut + 6.5 * extraBetween, br.Position.Y, ar.Position.Z)
                            Case "BOY" : ar.Position = New Point3d(baseX + extraOut + 8.5 * extraBetween, br.Position.Y, ar.Position.Z)
                        End Select
                    End If

                    ' Position adjustment: Rebar
                    If inst.IsRebar Then
                        Select Case ad.Tag.ToUpperInvariant()
                            Case "NO" : ar.Position = New Point3d(ar.Position.X, br.Position.Y, ar.Position.Z)
                            Case "ADET" : ar.Position = New Point3d(ar.Position.X + (extraOut + extraBetween), br.Position.Y, ar.Position.Z)
                            Case "CAP" : ar.Position = New Point3d(ar.Position.X + (extraOut + extraBetween) * 1.3, br.Position.Y, ar.Position.Z)
                            Case "YER" : ar.Position = New Point3d(ar.Position.X + (extraOut + extraBetween) * 2.5, br.Position.Y, ar.Position.Z)
                            Case "BOY" : ar.Position = New Point3d(ar.Position.X + (extraOut + extraBetween) * 3.8, br.Position.Y, ar.Position.Z)
                        End Select
                    End If

                    ' Set Attribute Content & FIELDS
                    Select Case ad.Tag.ToUpperInvariant()
                        Case "NO"
                            ar.TextString = inst.PosNo.ToString()
                            ar.ColorIndex = If(carpan = 0, 1, 2)
                        Case "ADET" : ar.TextString = inst.Adet
                        Case "CAP" : ar.TextString = "Ø" & inst.PhiMm.ToString()
                        Case "ARALIK" : ar.TextString = If(inst.Aralik > 0, "/" + inst.Aralik.ToString(), "")

                        Case "BOY"
                            If inst.LinkedEntityId.IsValid Then
                                Try
                                    ' 1. Get pointer
                                    Dim ptrStr As String = inst.LinkedEntityId.OldIdPtr.ToInt64().ToString()

                                    ' 2. Construct Field Code
                                    Dim fieldCode As String = "L=%<\AcObjProp Object(%<\_ObjId " & ptrStr & ">%).Length \f ""%lu2%pr0"">%"

                                    ' 3. Create Field and Set
                                    Dim fd As New Field(fieldCode)
                                    ar.SetField(fd)
                                    ' DO NOT call tr.AddNewlyCreatedDBObject(fd, True) here. 
                                    ' The AttributeReference owns the field.
                                Catch
                                    ' Fallback to static text if field creation fails
                                    ar.TextString = $"L={inst.TotalLen}"
                                End Try
                            Else
                                ar.TextString = $"L={inst.TotalLen}"
                            End If

                        Case "BUKUM_BOY"
                            ' =================================================
                            ' DYNAMIC BUKUM BOY (FORMULA FIELD)
                            ' =================================================
                            ' Logic: We assume Miters (hooks) are fixed. 
                            ' We create a field: "Miter1 [TotalLen - Miters] Miter2"
                            If inst.LinkedEntityId.IsValid AndAlso inst.MiterCount > 0 Then
                                Try
                                    ' 1. Get Pointer and Static Values
                                    Dim ptrStr As String = inst.LinkedEntityId.OldIdPtr.ToInt64().ToString()
                                    Dim miterSum As Integer = inst.Miter1 + inst.Miter2

                                    ' 2. Create inner field for Dynamic Total Length
                                    Dim lenField As String = "%<\AcObjProp Object(%<\_ObjId " & ptrStr & ">%).Length>%"

                                    ' 3. Create Formula for Straight Part: (TotalLength - MiterSum)
                                    ' \f "%lu2%pr0" ensures decimal format with precision 0
                                    Dim straightField As String = "%<\AcExpr " & lenField & " - " & miterSum & " \f ""%lu2%pr0"">%"

                                    Dim finalField As String = ""

                                    ' 4. Construct the final string based on shape (matching original text logic)
                                    If inst.MiterCount = 1 Then
                                        If inst.Miter1 > 0 Then
                                            ' Format: "Miter1 Straight"
                                            finalField = inst.Miter1.ToString() & " " & straightField
                                        ElseIf inst.Miter2 > 0 Then
                                            ' Format: "Miter2 Straight" (Preserving original order logic)
                                            finalField = inst.Miter2.ToString() & " " & straightField
                                        Else
                                            finalField = straightField
                                        End If
                                    ElseIf inst.MiterCount >= 2 Then
                                        ' Format: "Miter1 Straight Miter2"
                                        finalField = inst.Miter1.ToString() & " " & straightField & " " & inst.Miter2.ToString()
                                    Else
                                        finalField = straightField
                                    End If

                                    ' 5. Set the Field
                                    Dim fd As New Field(finalField)
                                    ar.SetField(fd)
                                Catch
                                    ' Fallback to static text
                                    ar.TextString = bukumBoy
                                End Try
                            Else
                                ' Static text for ties, stirrups, or failed links
                                ar.TextString = bukumBoy
                            End If
                        Case "DEMIR_TIPI"
                            If inst.IsTie Then
                                ar.TextString = "13"
                            ElseIf inst.IsStirrup Then
                                ar.TextString = "15"
                            Else
                                ar.TextString = (Math.Min(2, inst.MiterCount) + 1).ToString()
                            End If
                        Case "CARPAN" : ar.TextString = carpan.ToString()
                        Case "DEMIR_CINSI" : ar.TextString = ""
                        Case "YER" : ar.TextString = inst.Yer
                    End Select

                    ar.AlignmentPoint = ar.Position
                    ar.AdjustAlignment(db)

                    br.AttributeCollection.AppendAttribute(ar)
                    tr.AddNewlyCreatedDBObject(ar, True)
                Next
            Next
        End Sub

        Public Shared Sub ClearPending()
            _instances.Clear()
            _keyToCount.Clear()
        End Sub

        Public Shared Sub Reset(Optional nextStart As Integer = 1)
            _nextPosNumber = Math.Max(1, nextStart)
            _keyToPos.Clear()
            _keyToCount.Clear()
            _instances.Clear()
            _carpanDonati = 1
        End Sub

        Private Shared Function BuildSignature(phiMm As Integer, totalLen As Integer, yer As String,
                                miterCount As Integer, m1 As Integer, straightLen As Integer, m2 As Integer) As String
            Dim y = (If(yer, "")).Trim().ToLowerInvariant()
            Return $"φ{phiMm}|L{totalLen}|yer:{y}|m{miterCount}|a{m1}|s{straightLen}|b{m2}"
        End Function

        Private Class Instance
            Public Key As String
            Public PosNo As Integer
            Public Mid As Point3d
            Public PhiMm As Integer
            Public TotalLen As Integer
            Public Yer As String
            Public Adet As String
            Public MiterCount As Integer
            Public Miter1 As Integer
            Public StraightPart As Integer
            Public Miter2 As Integer
            Public IsTie As Boolean
            Public Aralik As Integer
            Public IsStirrup As Boolean
            Public IsRebar As Boolean
            Public IsForMetraj As Boolean
            Public LinkedEntityId As ObjectId
        End Class

    End Class

    ' =====================
    ' Helpers
    ' =====================
    Private Const LYR_POZ As String = "poz"
    Private Const BLK_POZ As String = "POZ_MARKER"
    Private Const TXTSTYLE_POZ As String = "poz"

    Private Sub EnsurePozLayer(db As Database, tr As Transaction)
        Dim lt = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
        If Not lt.Has(LYR_POZ) Then
            lt.UpgradeOpen()
            Dim lr As New LayerTableRecord() With {
                .Name = LYR_POZ,
                .Color = Color.FromColorIndex(ColorMethod.ByAci, 7)
            }
            lt.Add(lr) : tr.AddNewlyCreatedDBObject(lr, True)
        End If
    End Sub

    Private Sub EnsurePozTextStyle(db As Database, tr As Transaction)
        Dim tst = CType(tr.GetObject(db.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
        If Not tst.Has(TXTSTYLE_POZ) Then
            tst.UpgradeOpen()
            Dim ts As New TextStyleTableRecord() With {.Name = TXTSTYLE_POZ}
            ts.XScale = 0.6
            ts.TextSize = 55.0
            tst.Add(ts) : tr.AddNewlyCreatedDBObject(ts, True)
        Else
            Dim ts = CType(tr.GetObject(tst(TXTSTYLE_POZ), OpenMode.ForWrite), TextStyleTableRecord)
            If ts.TextSize = 0 Then ts.TextSize = 55.0
            If Math.Abs(ts.XScale - 0.6) > 0.0001 Then ts.XScale = 0.6
        End If
    End Sub

    Private Function MakePozAttributeDefinition(db As Database,
                                            tr As Transaction,
                                            tag As String,
                                            prompt As String,
                                            visible As Boolean,
                                            pos As Point3d,
                                            justify As AttachmentPoint,
                                            height As Double) As AttributeDefinition
        Dim ad As New AttributeDefinition() With {
        .Tag = tag, .Prompt = prompt, .TextString = "", .Position = pos, .Layer = LYR_POZ,
        .Verifiable = False, .Invisible = Not visible, .LockPositionInBlock = False,
        .Height = height, .Rotation = 0.0, .WidthFactor = 0.6, .Justify = justify
    }
        Dim tst = CType(tr.GetObject(db.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
        If tst.Has(TXTSTYLE_POZ) Then ad.TextStyleId = tst(TXTSTYLE_POZ)
        Return ad
    End Function

    Private Function EnsurePozBlock(db As Database,
                                tr As Transaction,
                                markerRadius As Double) As ObjectId
        Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
        If bt.Has(BLK_POZ) Then Return bt(BLK_POZ)

        bt.UpgradeOpen()
        Dim btr As New BlockTableRecord() With {.Name = BLK_POZ}
        Dim blkId = bt.Add(btr) : tr.AddNewlyCreatedDBObject(btr, True)

        Dim c As New Circle(Point3d.Origin, Vector3d.ZAxis, Math.Max(1.0, markerRadius)) With {
        .Layer = LYR_POZ, .ColorIndex = 256
    }
        btr.AppendEntity(c) : tr.AddNewlyCreatedDBObject(c, True)

        Dim H As Double = 5.5

        Dim attNO = MakePozAttributeDefinition(db, tr, "NO", "POS No", True, Point3d.Origin, AttachmentPoint.MiddleCenter, H)
        btr.AppendEntity(attNO) : tr.AddNewlyCreatedDBObject(attNO, True)

        Dim OUTSIDE_GAP As Double = 2.0
        Dim BETWEEN_GAP As Double = 3.0

        Dim adetPos As New Point3d(markerRadius + OUTSIDE_GAP, 0, 0)
        Dim attADET = MakePozAttributeDefinition(db, tr, "ADET", "Adet", True, adetPos, AttachmentPoint.MiddleLeft, H)
        btr.AppendEntity(attADET) : tr.AddNewlyCreatedDBObject(attADET, True)

        Dim capPos As New Point3d(markerRadius + OUTSIDE_GAP + BETWEEN_GAP, 0, 0)
        Dim attCAP = MakePozAttributeDefinition(db, tr, "CAP", "Cap", True, capPos, AttachmentPoint.MiddleLeft, H)
        btr.AppendEntity(attCAP) : tr.AddNewlyCreatedDBObject(attCAP, True)

        Dim boyPos As New Point3d(markerRadius + OUTSIDE_GAP + 8.5 * BETWEEN_GAP, 0, 0)
        Dim attBOY = MakePozAttributeDefinition(db, tr, "BOY", "Toplam Boy", True, boyPos, AttachmentPoint.MiddleLeft, H)
        btr.AppendEntity(attBOY) : tr.AddNewlyCreatedDBObject(attBOY, True)

        Dim aralikPos As New Point3d(markerRadius + OUTSIDE_GAP + 6 * BETWEEN_GAP, 0, 0)
        Dim attARALIK = MakePozAttributeDefinition(db, tr, "ARALIK", "Aralik", True, aralikPos, AttachmentPoint.MiddleLeft, H)
        btr.AppendEntity(attARALIK) : tr.AddNewlyCreatedDBObject(attARALIK, True)

        Dim yerPos As New Point3d(markerRadius + OUTSIDE_GAP + 4.5 * BETWEEN_GAP, 0, 0)
        Dim attYer = MakePozAttributeDefinition(db, tr, "YER", "Demir Yeri", True, yerPos, AttachmentPoint.MiddleLeft, H)
        btr.AppendEntity(attYer) : tr.AddNewlyCreatedDBObject(attYer, True)

        Dim baseP As New Point3d(markerRadius + 10.0, 0, 0)
        Dim invs As (Tag As String, Prompt As String)() = {
        ("BUKUM_BOY", "Bukum Boylari"),
        ("DEMIR_TIPI", "Demir Tipi"),
        ("CARPAN", "Carpan"),
        ("DEMIR_CINSI", "Demir Cinsi")
    }

        For i = 0 To invs.Length - 1
            Dim ad = MakePozAttributeDefinition(db, tr, invs(i).Tag, invs(i).Prompt, False, New Point3d(baseP.X, baseP.Y - i * 5.0, 0), AttachmentPoint.MiddleLeft, 55.0)
            btr.AppendEntity(ad) : tr.AddNewlyCreatedDBObject(ad, True)
        Next

        Return blkId
    End Function

End Module