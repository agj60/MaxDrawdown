Attribute VB_Name = "Module1"
Function MaxDrawdown(rng As Range) As Variant
    Dim n As Long, m As Long
    Dim prices() As Double
    Dim dates() As Variant
    Dim maxDD() As Double
    Dim startDD() As Variant
    Dim maxDDDate() As Variant
    Dim recoveryDate() As Variant
    Dim i As Long, j As Long
    Dim peak As Double
    Dim trough As Double
    Dim dd As Double
    Dim recDate As Variant
    Dim startPeakIndex As Long
    Dim startThroughIndex As Long
    
    n = rng.Rows.Count
    m = rng.Columns.Count
    
    ReDim prices(1 To n - 1, 1 To m - 1)
    ReDim dates(1 To n - 1)
    ReDim maxDD(1 To m - 1)
    ReDim startDD(1 To m - 1)
    ReDim maxDDDate(1 To m - 1)
    ReDim recoveryDate(1 To m - 1)
    
    ' Fill arrays with data
    For i = 2 To n
        dates(i - 1) = rng.Cells(i, 1).value
        For j = 2 To m
            prices(i - 1, j - 1) = rng.Cells(i, j).value
        Next j
    Next i
    
    ' Calculate drawdowns
    For j = 1 To m - 1
        peak = prices(1, j)
        trough = prices(1, j)
        dd = 0
        For i = 2 To n - 1
            If prices(i, j) > peak Then
                peak = prices(i, j)
                trough = prices(i, j)
            End If
            If prices(i, j) < trough Then
                trough = prices(i, j)
                If (peak - trough) / peak > dd Then
                    dd = (peak - trough) / peak
                    maxDD(j) = dd
                    startPeakIndex = FindIndex(prices, peak, j)
                    startDD(j) = dates(startPeakIndex)
                    maxDDDate(j) = dates(i)
                    startThroughIndex = i
                    peakPriorToDD = peak
                End If
            End If
        Next i
        
        ' Find recovery date
        recDate = "N/A"
        For i = startThroughIndex To n - 1
            If prices(i, j) >= peakPriorToDD Then
                recDate = dates(i)
                Exit For
            End If
        Next i
        recoveryDate(j) = recDate
    Next j
    
    ' Fill in results
    Dim result As Variant
    ReDim result(1 To 4, 1 To m - 1)
    For j = 1 To m - 1
        result(1, j) = -maxDD(j)
        result(2, j) = startDD(j)
        result(3, j) = maxDDDate(j)
        result(4, j) = recoveryDate(j)
    Next j
    
    MaxDrawdown = result
End Function

Function FindIndex(arr As Variant, value As Double, col As Long) As Long
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, col) = value Then
            FindIndex = i
            Exit Function
        End If
    Next i
    FindIndex = -1 ' in case value is not found
End Function
