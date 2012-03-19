Module Module1
    Dim computedCount As Int64 = 0
    Dim C1Delta As Int64
    Dim C2Delta As Int64
    Dim C3Delta As Int64

    Dim C1S As Int64 = 16968696
    Dim C1F As Int64 = 17239680
    Dim C2S As Int64 = 36181656
    Dim C2F As Int64 = 36456552
    Dim C3S As Int64 = 16960512
    Dim C3F As Int64 = 17232288

    Dim PackCount As Int32 = 24

    Sub Main()
        Dim logicUsed As String = ""

        C1Delta = C1F - C1S
        C2Delta = C2F - C2S
        C3Delta = C3F - C3S
        CountLogic(PackCount, C1Delta, C2Delta, C3Delta, logicUsed)
        Debug.Print("Result was: " + computedCount.ToString)
        Debug.Print("Logic used: " + logicUsed)

    End Sub

    Sub CountLogic(ByRef PackCount As Int32, ByRef C1Delta As Int64, ByRef C2Delta As Int64, ByRef C3Delta As Int64, ByRef logicUsed As String)
        Dim C1D As Int64
        Dim C2D As Int64
        Dim C3D As Int64

        Dim deltaHi As Int64
        Dim deltaLo As Int64
        Dim deltaVar As Double

        Dim deltaGroup As Int32() = {0, 1, 2}
        Dim deltaRange As Int64() = {999999999, 999999999, 999999999}
        Dim deltaName As String() = {"C1-C2", "C1-C3", "C2-C3"}

        Dim useD1 As Boolean = False
        Dim useD2 As Boolean = False
        Dim useD3 As Boolean = False
        Dim useNumerator As Int64 = 0
        Dim useDenom As Int64 = 0

        logicUsed = ("Pack Count: " + PackCount.ToString)

        C1D = Convert.ToInt64(C1Delta / PackCount)
        C2D = Convert.ToInt64(C2Delta / PackCount)
        C3D = Convert.ToInt64(C3Delta / PackCount)

        logicUsed = logicUsed + vbCrLf + ("Adjusted Counter Deltas: C1 " + C1D.ToString + " C2 " + C2D.ToString + " C3 " + C3D.ToString)

        'Don't use counters that didn't change, they could be
        'broken or not present

        If (C1Delta > 0) Then
            useD1 = True
            useDenom = useDenom + 1
            deltaHi = C1D
            deltaLo = C1D
        End If

        'logicUsed = logicUsed + vbCrLf + ("Use C1Delta is " + useD1.ToString)
        'logicUsed = logicUsed + vbCrLf + ("deltaHi is " + deltaHi.ToString)
        'logicUsed = logicUsed + vbCrLf + ("deltaLo is " + deltaLo.ToString)

        If (C2Delta > 0) Then
            useD2 = True
            useDenom = useDenom + 1
            If C2Delta > deltaHi Then deltaHi = C2D
            If C2Delta < deltaLo Then deltaLo = C2D
        End If

        'logicUsed = logicUsed + vbCrLf + ("Use C2Delta is " + useD2.ToString)
        'logicUsed = logicUsed + vbCrLf + ("deltaHi is " + deltaHi.ToString)
        'logicUsed = logicUsed + vbCrLf + ("deltaLo is " + deltaLo.ToString)

        If (C3Delta > 0) Then
            useD3 = True
            useDenom = useDenom + 1
            If C3Delta > deltaHi Then deltaHi = C3D
            If C3Delta < deltaLo Then deltaLo = C3D
        End If

        'logicUsed = logicUsed + vbCrLf + ("Use C3Delta is " + useD3.ToString)
        'logicUsed = logicUsed + vbCrLf + ("deltaHi is " + deltaHi.ToString)
        'logicUsed = logicUsed + vbCrLf + ("deltaLo is " + deltaLo.ToString)


        'If no counters had changes, the outcome is zero by definition
        If Not (useD1 Or useD2 Or useD3) Then
            computedCount = 0
        Else
            'At least one counter registered production
            deltaVar = (deltaHi - deltaLo) / deltaHi

            logicUsed = logicUsed + vbCrLf + ("Checking business rule inputs:")
            logicUsed = logicUsed + vbCrLf + ("delta range is " + (deltaHi - deltaLo).ToString)
            logicUsed = logicUsed + vbCrLf + ("delta variance is " + deltaVar.ToString)

            'Business rule, per FRS:
            If (System.Math.Abs(deltaHi - deltaLo) > 25) And (0.003 <= deltaVar) Then
                'Pick the average of the "closest two"
                logicUsed = logicUsed + vbCrLf + ("Rule says pick average of closest two counters")
                'Now assign the deltas and find the "closest" pair
                'of counters, which is the one with the lowest delta.
                'If two are equally close, use the higher delta.
                If (useD1 And useD2) Then
                    deltaRange(0) = System.Math.Abs(C1D - C2D)
                End If
                If (useD1 And useD3) Then
                    deltaRange(1) = System.Math.Abs(C1D - C3D)
                End If
                If (useD2 And useD3) Then
                    deltaRange(2) = System.Math.Abs(C2D - C3D)
                End If

                logicUsed = logicUsed + vbCrLf + ("Spread (C1-C2) is " + deltaRange(0).ToString)
                logicUsed = logicUsed + vbCrLf + ("Spread (C1-C3) is " + deltaRange(1).ToString)
                logicUsed = logicUsed + vbCrLf + ("Spread (C2-C3) is " + deltaRange(2).ToString)

                'Sort by the range
                Array.Sort(deltaRange, deltaGroup)

                logicUsed = logicUsed + vbCrLf + ("Closest spread is " + deltaRange(0).ToString + " between " + deltaName(deltaGroup(0)))
                logicUsed = logicUsed + vbCrLf + ("Computing by averaging " + deltaName(deltaGroup(0)))

                Select Case deltaGroup(0)
                    Case 0
                        computedCount = Convert.ToInt64((C1D + C2D) / 2)
                    Case 1
                        computedCount = Convert.ToInt64((C1D + C3D) / 2)
                    Case Else
                        computedCount = Convert.ToInt64((C2D + C3D) / 2)
                End Select
            Else
                logicUsed = logicUsed + vbCrLf + ("Rule says average all available counters")
                'Compute the average of all counters present as the
                'default business case value
                If useD1 Then useNumerator = C1D
                If useD2 Then useNumerator = useNumerator + C2D
                If useD3 Then useNumerator = useNumerator + C3D
                computedCount = CLng(Int(useNumerator / useDenom))
            End If

        End If
        logicUsed = logicUsed + vbCrLf + ("Rule result: " + computedCount.ToString + " units")
    End Sub
End Module
