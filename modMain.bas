Option Explicit

Sub Main()
    Dim tles(0 To 100) As Variant
    Dim numTLEs As Long

    numTLEs = loadTLEs(tles)
    runVer tles, numTLEs
End Sub
