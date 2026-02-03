Option Strict On
Option Infer On

Module BeamSpanModule

    ''' <summary>
    ''' Builds continuous (x0,x1) spans common to both the top and bottom beam lines on a level
    ''' when there are NO columns interrupting them.
    ''' </summary>
    Public Function BuildSpansNoColumns(
        allBeams As List(Of ReinforcementBarAutomation.HLine),
        topY As Double, bottomY As Double,
        epsY As Double, epsLen As Double
    ) As List(Of Tuple(Of Double, Double))

        Dim topSegs As New List(Of Tuple(Of Double, Double))()
        Dim botSegs As New List(Of Tuple(Of Double, Double))()

        For Each h In allBeams
            If Math.Abs(h.Y - topY) < epsY Then
                Dim a = Math.Min(h.X0, h.X1), b = Math.Max(h.X0, h.X1)
                If b - a > epsLen Then topSegs.Add(Tuple.Create(Of Double, Double)(a, b))
            ElseIf Math.Abs(h.Y - bottomY) < epsY Then
                Dim a = Math.Min(h.X0, h.X1), b = Math.Max(h.X0, h.X1)
                If b - a > epsLen Then botSegs.Add(Tuple.Create(Of Double, Double)(a, b))
            End If
        Next

        Dim spans As New List(Of Tuple(Of Double, Double))()
        If topSegs.Count = 0 OrElse botSegs.Count = 0 Then Return spans

        ' Breakpoints from both rows
        Dim bps As New List(Of Double)()
        For Each s In topSegs : bps.Add(s.Item1) : bps.Add(s.Item2) : Next
        For Each s In botSegs : bps.Add(s.Item1) : bps.Add(s.Item2) : Next
        bps.Sort()

        ' Unique X’s (with eps)
        Dim uniq As New List(Of Double)()
        For Each x In bps
            If uniq.Count = 0 OrElse Math.Abs(x - uniq(uniq.Count - 1)) > epsLen Then uniq.Add(x)
        Next
        If uniq.Count < 2 Then Return spans

        ' Helper: does [a,b] lie fully inside any segment of the list?
        Dim covers As Func(Of List(Of Tuple(Of Double, Double)), Double, Double, Boolean) =
        Function(list As List(Of Tuple(Of Double, Double)), a As Double, b As Double)
            For Each s In list
                If s.Item1 <= a + epsLen AndAlso s.Item2 >= b - epsLen Then Return True
            Next
            Return False
        End Function

        ' Merge adjacent covered intervals into maximal spans
        Dim i As Integer = 0
        While i < uniq.Count - 1
            Dim a = uniq(i), b = uniq(i + 1)
            If b - a > epsLen AndAlso covers(topSegs, a, b) AndAlso covers(botSegs, a, b) Then
                Dim j = i + 1, endX = b
                While j < uniq.Count - 1 AndAlso covers(topSegs, a, uniq(j + 1)) AndAlso covers(botSegs, a, uniq(j + 1))
                    j += 1 : endX = uniq(j)
                End While
                spans.Add(Tuple.Create(Of Double, Double)(a, endX))
                i = j
            Else
                i += 1
            End If
        End While

        Return spans
    End Function

End Module