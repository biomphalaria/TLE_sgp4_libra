Option Explicit

Public Function loadTLEs(tles() As Variant) As Long
    Dim row As Long
    Dim tl As Variant
    Dim sht As Worksheet
    Dim l1 As String
    Dim l2 As String
    Dim ind As Long

    Set sht = ThisWorkbook.Worksheets("SGP4-VER.TLE")

    row = 0
    ind = 0
    Do While Len(sht.Cells(row + 1, 1).Value) > 1
        l1 = sht.Cells(row + 2, 1).Value
        l2 = sht.Cells(row + 3, 1).Value
        tl = parseTLE(l1, l2)
        tles(ind) = tl
        ind = ind + 1
        row = row + 3
    Loop

    loadTLEs = ind
End Function
