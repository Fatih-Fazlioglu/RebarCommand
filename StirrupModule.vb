Option Strict On
Option Infer On

Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry


' Draw a standalone stirrup UNDER a section:
' - right edge leaning to RIGHT
' - top & right edges DO NOT TOUCH (gap at corner)
' - two leaders drawn inward (down-left) from top/right near the corner
Public Module StirrupModule

    Private Const LYR_REBAR1 As String = "kr_donati1"
    Private Const LYR_TEXT As String = "kr_dnyazi"

    ' ---------- helpers ----------
    Private Sub AddEnt(btr As BlockTableRecord, tr As Transaction, ent As Entity, layerName As String)
        ent.Layer = layerName : ent.ColorIndex = 256
        btr.AppendEntity(ent) : tr.AddNewlyCreatedDBObject(ent, True)
    End Sub

    Private Sub Txt(btr As BlockTableRecord, tr As Transaction, p As Point3d, s As String,
                    Optional h As Double = 3.333, Optional anchor As AttachmentPoint = AttachmentPoint.BottomCenter)
        Dim mt As New MText With {.Location = p, .Attachment = anchor, .TextHeight = h, .Contents = s}
        AddEnt(btr, tr, mt, LYR_TEXT)
    End Sub

    ' ---------- main ----------
    Public Sub DrawBelow(
        btr As BlockTableRecord, tr As Transaction,
        secBaseX As Double, secBaseY As Double, secW As Double, secH As Double,
        Optional coverCm As Double = 3.0,
        Optional gapBelowCm As Double = 400.0, _                 ' << 400 cm below section
        Optional gapCornerCm As Double = 6.0, _                  ' << (Unused now, kept for signature compatibility)
        Optional skewRightCm As Double = 12.0, _                 ' << right edge leans to RIGHT
        Optional leaderLenX As Double = 10.0, _                  ' << leader vector X (inward, left)
        Optional leaderLenY As Double = 8.0, _                   ' << leader vector Y (inward, down)
        Optional hookText As String = "12", _                    ' << the two small labels
        Optional textH As Double = 3.333)

        ' ---- inner clear dimensions (must match section’s stirrup) ----
        Dim stirrupW As Double = Math.Max(1.0, secW - 2 * coverCm)
        Dim stirrupH As Double = Math.Max(1.0, secH - 2 * coverCm)

        ' ---- place below section (keep exact X) ----
        Dim baseX As Double = secBaseX
        Dim baseY As Double = secBaseY - gapBelowCm - stirrupH

        ' ---- key points ----
        Dim TL As New Point3d(baseX, baseY + stirrupH, 0)                                        ' top-left
        Dim BL As New Point3d(baseX, baseY, 0)                                                           ' bottom-left
        Dim BR As New Point3d(baseX + stirrupW + skewRightCm, baseY, 0)                  ' bottom-right (skewed)

        ' Corner Logic: 
        ' Top bar goes to the absolute top-right corner (Outer).
        ' Right bar goes to a slightly offset inner corner (Inner) to simulate overlap.
        Dim TR_Outer As New Point3d(baseX + stirrupW, baseY + stirrupH, 0)

        ' Offset for the inner hook (right leg) so lines don't clash perfectly
        Dim innerOffset As New Vector3d(-1.5, -2.5, 0)
        Dim TR_Inner As Point3d = TR_Outer.Add(innerOffset)

        ' ---- outline segments ----
        AddEnt(btr, tr, New Line(TL, TR_Outer), LYR_REBAR1)      ' Top Bar
        AddEnt(btr, tr, New Line(TL, BL), LYR_REBAR1)            ' Left Bar
        AddEnt(btr, tr, New Line(BL, BR), LYR_REBAR1)            ' Bottom Bar
        AddEnt(btr, tr, New Line(BR, TR_Inner), LYR_REBAR1)      ' Right Bar (goes to inner corner)

        ' ---- Hooks (Connected Lines) ----
        ' Vector pointing down-left
        Dim hookVec As New Vector3d(-leaderLenX, -leaderLenY, 0)

        ' Hook 1: From Top Bar End (Outer)
        Dim hook1_End As Point3d = TR_Outer.Add(hookVec)
        AddEnt(btr, tr, New Line(TR_Outer, hook1_End), LYR_REBAR1)

        ' Hook 2: From Right Bar End (Inner)
        Dim hook2_End As Point3d = TR_Inner.Add(hookVec)
        AddEnt(btr, tr, New Line(TR_Inner, hook2_End), LYR_REBAR1)

        ' ---- "12" texts near hook ends ----
        Txt(btr, tr, New Point3d(hook1_End.X - 1.0, hook1_End.Y + 1.0, 0), hookText, textH, AttachmentPoint.BottomRight)
        Txt(btr, tr, New Point3d(hook2_End.X - 1.0, hook2_End.Y - 1.0, 0), hookText, textH, AttachmentPoint.TopRight)

        ' ---- numeric side labels (just numbers, mid-sides) ----
        ' Top (stirrupW)
        Dim topMid As New Point3d((TL.X + TR_Outer.X) / 2.0, TL.Y + 6.0, 0)
        Txt(btr, tr, topMid, CInt(Math.Round(stirrupW)).ToString(), textH)

        ' Bottom (stirrupW)
        Dim botMid As New Point3d((BL.X + BR.X) / 2.0, BL.Y - 6.0, 0)
        Txt(btr, tr, botMid, CInt(Math.Round(stirrupW)).ToString(), textH, AttachmentPoint.TopCenter)

        ' Left (stirrupH)
        Dim leftMid As New Point3d(TL.X - 6.0, (TL.Y + BL.Y) / 2.0, 0)
        Txt(btr, tr, leftMid, CInt(Math.Round(stirrupH)).ToString(), textH, AttachmentPoint.MiddleRight)

        ' Right (stirrupH) – mid of right slanted segment
        Dim rightMid As New Point3d((TR_Inner.X + BR.X) / 2.0 + 6.0, (TR_Inner.Y + BR.Y) / 2.0, 0)
        Txt(btr, tr, rightMid, CInt(Math.Round(stirrupH)).ToString(), textH, AttachmentPoint.MiddleLeft)
    End Sub

End Module