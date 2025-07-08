Option Explicit

Public Sub runVer(tles() As Variant, cnt As Long)
    Dim tl As Variant
    Dim sht As Worksheet
    Dim sht2 As Worksheet
    Dim i As Long
    Dim row As Long
    Dim str As String
    Dim mins As Double
    Dim rv(3) As Double
    Dim vv(3) As Double
    Dim r(3) As Double
    Dim v(3) As Double
    Dim sa As Variant
    Dim rdist As Double
    Dim vdist As Double
    Dim rerr As Double
    Dim verr As Double
    Dim cnt2 As Long

    rerr = 0
    verr = 0
    cnt2 = 0

    i = -1
    Set sht = ThisWorkbook.Worksheets("tcppver.out")
    Set sht2 = ThisWorkbook.Worksheets("TestSGP4")

    row = 0
    Do While Len(sht.Cells(row + 1, 1).Value) > 1
        str = sht.Cells(row + 1, 1).Value
        If InStr(str, "xx") > 0 Then
            i = i + 1
            tl = tles(i)
        Else
            str = ws(str)
            sa = Split(str, " ")
            mins = CDbl(sa(0))
            rv(0) = CDbl(sa(1))
            rv(1) = CDbl(sa(2))
            rv(2) = CDbl(sa(3))
            vv(0) = CDbl(sa(4))
            vv(1) = CDbl(sa(5))
            vv(2) = CDbl(sa(6))

            getRV tl, mins, r, v

            rdist = dist(r, rv)
            vdist = dist(v, vv)
            rerr = rerr + rdist
            verr = verr + vdist
            cnt2 = cnt2 + 1
            If rdist > 1E-4 Or vdist > 1E-6 Then
                sht2.Cells(row + 4, 1).Value = str
                sht2.Cells(row + 4, 2).Value = rerr
                sht2.Cells(row + 4, 3).Value = verr
            End If
        End If
        row = row + 1
    Loop

    rerr = rerr / cnt2
    verr = verr / cnt2

    rerr = rerr * 1E6
    verr = verr * 1E6
    sht2.Cells(1, 1).Value = "Typical errors"
    sht2.Cells(2, 1).Value = rerr
    sht2.Cells(3, 1).Value = verr
    sht2.Cells(2, 2).Value = "mm"
    sht2.Cells(3, 2).Value = "mm/s"
End Sub
