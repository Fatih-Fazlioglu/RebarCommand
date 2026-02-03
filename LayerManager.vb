Option Strict On
Option Infer On

' ========= AutoCAD =========
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Colors

' =====================================================================
' LAYER MANAGER — create/ensure layers & append entities BYLAYER
' =====================================================================
Public Module LayerManager

    ' --- Canonical layer names used across your commands ---
    Public Const LYR_REBAR1 As String = "kr_donati1"
    Public Const LYR_REBAR2 As String = "kr_donati2"
    Public Const LYR_PHI As String = "kr_dnyazi"
    Public Const LYR_DIM As String = "kr_olcu"
    Public Const LYR_BEAMTITLE As String = "kr_kiris_baslik"
    Public Const LYR_MEASURE As String = "kr_measure_aux"

    ' Brown
    Public ReadOnly REBAR_BROWN As Color = Color.FromRgb(127, 95, 63)

    ' -----------------------------------------------------------------
    ' Ensure a single layer exists (creates if missing) and returns Id
    ' -----------------------------------------------------------------
    Public Function EnsureLayer(db As Database,
                                tr As Transaction,
                                layerName As String,
                                Optional byColor As Color = Nothing,
                                Optional linetypeName As String = "Continuous",
                                Optional lw As LineWeight = LineWeight.ByLineWeightDefault,
                                Optional isPlottable As Boolean = True) As ObjectId

        If String.IsNullOrWhiteSpace(layerName) Then
            Throw New System.ArgumentException("Layer name cannot be empty.", NameOf(layerName))
        End If

        Dim lt = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)

        ' Return existing
        If lt.Has(layerName) Then
            Return lt(layerName)
        End If

        ' Need to create
        lt.UpgradeOpen()

        ' Ensure requested linetype exists (loads if needed)
        Dim ltypeId As ObjectId = EnsureLinetypeLoaded(db, tr, linetypeName)

        Dim ltr As New LayerTableRecord() With {
            .Name = layerName,
            .IsPlottable = isPlottable,
            .LineWeight = lw,
            .LinetypeObjectId = ltypeId
        }

        ' If caller provided a color, use it; otherwise keep ByLayer (256)
        If byColor IsNot Nothing Then
            ltr.Color = byColor
        Else
            ltr.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256)
        End If

        Dim id = lt.Add(ltr)
        tr.AddNewlyCreatedDBObject(ltr, True)
        Return id
    End Function

    ' -----------------------------------------------------------------
    ' Ensure a set of standard layers you use everywhere
    ' (Call once per drawing, e.g., at command start)
    ' -----------------------------------------------------------------
    Public Sub EnsureAllLayers(db As Database, tr As Transaction)
        ' Rebar layers (brown, continuous)
        EnsureLayer(db, tr, LYR_REBAR1, REBAR_BROWN, "Continuous", LineWeight.LineWeight025, True)
        EnsureLayer(db, tr, LYR_REBAR2, REBAR_BROWN, "Continuous", LineWeight.LineWeight020, True)

        ' Text/labels (white by default)
        EnsureLayer(db, tr, LYR_PHI, Color.FromColorIndex(ColorMethod.ByAci, 5), "Continuous", LineWeight.LineWeight018, True)
        EnsureLayer(db, tr, LYR_BEAMTITLE, Color.FromColorIndex(ColorMethod.ByAci, 5), "Continuous", LineWeight.LineWeight020, True)

        ' Dimensions / measures
        EnsureLayer(db, tr, LYR_DIM, Nothing, "Continuous", LineWeight.LineWeight018, True)
        EnsureLayer(db, tr, LYR_MEASURE, Nothing, "Continuous", LineWeight.ByLineWeightDefault, False)
    End Sub

    ' -----------------------------------------------------------------
    ' Append entity to BTR on a specific layer (creates layer if needed)
    ' If forceByLayer=True → ent.ColorIndex = 256 (BYLAYER)
    ' If colorOverride is provided and forceByLayer=False → set that color
    ' -----------------------------------------------------------------
    Public Sub AppendOnLayer(btr As BlockTableRecord,
                             tr As Transaction,
                             ent As Entity,
                             layerName As String,
                             Optional forceByLayer As Boolean = True,
                             Optional colorOverride As Color = Nothing)

        If ent Is Nothing Then Throw New System.ArgumentNullException(NameOf(ent))
        If btr Is Nothing Then Throw New System.ArgumentNullException(NameOf(btr))

        Dim db = btr.Database

        ' Ensure target layer exists
        EnsureLayer(db, tr, layerName)

        ent.Layer = layerName
        If forceByLayer Then
            ent.ColorIndex = 256 ' BYLAYER
        ElseIf colorOverride IsNot Nothing Then
            ent.Color = colorOverride
        End If

        btr.AppendEntity(ent)
        tr.AddNewlyCreatedDBObject(ent, True)
    End Sub

    ' -----------------------------------------------------------------
    ' Make a layer current (creates if needed)
    ' -----------------------------------------------------------------
    Public Sub SetCurrentLayer(db As Database, tr As Transaction, layerName As String)
        Dim id = EnsureLayer(db, tr, layerName)
        db.Clayer = id
    End Sub

    ' -----------------------------------------------------------------
    ' Ensure a linetype exists in the drawing; load if missing
    ' Returns ObjectId of the linetype
    ' -----------------------------------------------------------------
    Public Function EnsureLinetypeLoaded(db As Database,
                                         tr As Transaction,
                                         linetypeName As String) As ObjectId
        If String.IsNullOrWhiteSpace(linetypeName) Then
            linetypeName = "Continuous"
        End If

        Dim ltt = CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)
        If ltt.Has(linetypeName) Then
            Return ltt(linetypeName)
        End If

        ' Try to load from acad.lin / acadiso.lin – silently fall back to Continuous
        Try
            ltt.UpgradeOpen()
            Using ltr As New LinetypeTableRecord()
                ltr.Name = linetypeName
                ' The simplest and most robust way in code-only scenarios is to call LoadLineTypeFile.
                ' If that throws, we'll just return Continuous.
            End Using
            db.LoadLineTypeFile(linetypeName, "acad.lin")
            Return CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)(linetypeName)
        Catch
            ' Fallback: Continuous always exists
            Return CType(tr.GetObject(db.LinetypeTableId, OpenMode.ForRead), LinetypeTable)("Continuous")
        End Try
    End Function

End Module