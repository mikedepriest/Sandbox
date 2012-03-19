Module Module1

    Sub Main()
        Dim cc As Int64
        Dim c1d As Int64 = 1100
        Dim c2d As Int64 = 1000
        Dim c3d As Int64 = 1050
        codeBlock(c1d, c2d, c3d, cc)
        Console.Write(cc.ToString)
    End Sub

    Sub codeBlock(ByRef C1Delta As Int64, ByRef C2Delta As Int64, ByRef C3Delta As Int64, ByRef computedCount As Int64)
        Dim deltaHi As Int64
        Dim deltaLo As Int64
        Dim deltaVar As Single

        Dim deltaGroup As Int32() = {0, 1, 2}
        Dim deltaRange As Int64() = {999999999, 999999999, 999999999}

        Dim useD1 As Boolean = False
        Dim useD2 As Boolean = False
        Dim useD3 As Boolean = False
        Dim useNumerator As Int64 = 0
        Dim useDenom As Int32 = 0

        'Don't use counters that didn't change, they could be
        'broken or not present

        If (C1Delta > 0) Then
            useD1 = True
            useDenom = useDenom + 1
            deltaHi = C1Delta
            deltaLo = C1Delta
        End If

        If (C2Delta > 0) Then
            useD2 = True
            useDenom = useDenom + 1
            If C2Delta > deltaHi Then deltaHi = C2Delta
            If C2Delta < deltaLo Then deltaLo = C2Delta
        End If

        If (C3Delta > 0) Then
            useD3 = True
            useDenom = useDenom + 1
            If C3Delta > deltaHi Then deltaHi = C3Delta
            If C3Delta < deltaLo Then deltaLo = C3Delta
        End If

        'Now assign the deltas and find the "closest" pair
        'of counters, which is the one with the lowest delta.
        'If two are equally close, use the higher delta.
        If (useD1 And useD2) Then
            deltaRange(0) = System.Math.Abs(C1Delta - C2Delta)
        End If
        If (useD1 And useD3) Then
            deltaRange(1) = System.Math.Abs(C1Delta - C3Delta)
        End If
        If (useD2 And useD3) Then
            deltaRange(2) = System.Math.Abs(C2Delta - C3Delta)
        End If

        'Sort by the range
        Array.Sort(deltaRange, deltaGroup)

        'If no counters had changes, the outcome is zero by definition
        If Not (useD1 Or useD2 Or useD3) Then
            computedCount = 0
        Else
            'At least one counter registered production
            'Compute the average of all counters present as the
            'default business case value
            If useD1 Then useNumerator = C1Delta
            If useD2 Then useNumerator = useNumerator + C2Delta
            If useD3 Then useNumerator = useNumerator + C3Delta
            computedCount = CLng(Int(useNumerator / useDenom))

            deltaVar = (deltaHi - deltaLo) / deltaHi
            'Business rule, per FRS:
            If (System.Math.Abs(deltaHi - deltaLo) > 25) And (0.003 <= deltaVar) Then
                'Pick the average of the "closest two"
                Select Case deltaGroup(0)
                    Case 0
                        computedCount = CLng(Int((C1Delta + C2Delta) / 2))
                    Case 1
                        computedCount = CLng(Int((C1Delta + C3Delta) / 2))
                    Case Else
                        computedCount = CLng(Int((C2Delta + C3Delta) / 2))
                End Select
            End If

        End If


    End Sub
End Module
