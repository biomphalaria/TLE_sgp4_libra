Option Explicit

Public Function dist(v1 As Variant, v2 As Variant) As Double
    Dim d As Double
    Dim tmp As Double

    tmp = v1(0) - v2(0)
    d = tmp * tmp
    tmp = v1(1) - v2(1)
    d = d + tmp * tmp
    tmp = v1(2) - v2(2)
    d = d + tmp * tmp

    dist = Sqr(d)
End Function

Public Function ws(str As String) As String
    Dim t As String
    t = Trim(str)
    t = Replace(t, "  ", " ")
    t = Replace(t, "  ", " ")
    t = Replace(t, "  ", " ")
    t = Replace(t, "  ", " ")
    ws = t
End Function
