Option Strict On
Option Infer On

Imports System.Globalization

Public Module MathUtils

    ' ---------- Rounds positive numbers up to the next multiple of 5 ----------
    Public Function Round5Up(value As Double) As Double
        If value <= 0 OrElse Double.IsNaN(value) OrElse Double.IsInfinity(value) Then Return 0
        Return Math.Ceiling(value / 5.0R) * 5.0R
    End Function

    ' ---------- Formats a length with 0 decimals using invariant culture ----------
    Public Function FormatLen(v As Double) As String
        Return Math.Round(v, 0).ToString(CultureInfo.InvariantCulture)
    End Function

    ' ---------- Unique + sorted with tolerance ----------
    ' Preferred flexible overload
    Public Function UniqueSorted(values As IEnumerable(Of Double), tol As Double) As List(Of Double)
        Dim list As List(Of Double) = If(values IsNot Nothing, New List(Of Double)(values), New List(Of Double)())
        list.Sort()

        Dim u As New List(Of Double)()
        For Each v In list
            If u.Count = 0 OrElse Math.Abs(v - u(u.Count - 1)) > tol Then
                u.Add(v)
            End If
        Next
        Return u
    End Function

    ' Back-compat overload with your original signature
    Public Function UniqueSorted(values As List(Of Double), tol As Double) As List(Of Double)
        Return UniqueSorted(DirectCast(values, IEnumerable(Of Double)), tol)
    End Function

    ' ---------- Pair consecutive beam levels into (top,bottom) tuples ----------
    ' Expects levels already sorted (ascending or descending is fine).
    Public Function PairBeamLevels(sortedLevels As List(Of Double)) As List(Of Tuple(Of Double, Double))
        Dim groups As New List(Of Tuple(Of Double, Double))()
        If sortedLevels Is Nothing Then Return groups

        For i As Integer = 0 To sortedLevels.Count - 2 Step 2
            Dim y0 = sortedLevels(i)
            Dim y1 = sortedLevels(i + 1)
            ' Ensure tuple is (top, bottom)
            groups.Add(Tuple.Create(Of Double, Double)(Math.Max(y0, y1), Math.Min(y0, y1)))
        Next

        Return groups
    End Function

End Module