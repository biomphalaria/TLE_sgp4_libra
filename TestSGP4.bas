
REM Loads the test TLEs from SGP4-VER.TLE
function loadTLEs(tles)
  Dim row as Integer
  Dim tl as Variant
  Dim sht as Variant
  Dim l1 as String
  Dim l2 as String
  Dim ind as Integer
  
  sht = ThisComponent.Sheets.getByName("SGP4-VER.TLE")
  
  ind = 0
  
 
  Do while(Len(sht.getCellByPosition(0,row).String) > 1)
     l1 = sht.getCellByPosition(0,row+1).String
     l2 = sht.getCellByPosition(0,row+2).String
     
     tl = parseTLE(l1,l2)
     tles(ind)=tl
     ind = ind+1
     row = row +3
  Loop

  loadTLEs = ind
end function


function dist(v1 as Variant, v2 as Variant)
  Dim d as Double
  Dim tmp as Double
  d = 0
  tmp = 0
  
  tmp = v1(0)-v2(0)
  d = d + (tmp*tmp)
  tmp = v1(1)-v2(1)
  d = d + (tmp*tmp)
  tmp = v1(2)-v2(2)
  d = d + (tmp*tmp)
  
  dist = Sqr(d)
end function

function ws(str as String) as String
  Dim t as String
  t = trim(str)
  t = replace(t,"  "," ")
  t = replace(t,"  "," ")
  t = replace(t,"  "," ")
  t = replace(t,"  "," ")
  ws = t
end function

REM Verify the propagation of the TLEs
sub runVer(tles As Variant, cnt as Integer)
  Dim tl as Variant
  Dim sht as Variant
  Dim sht2 as Variant
  Dim i as Integer
  Dim row as Integer
  Dim str as String
  Dim mins as Double
  Dim rv(3) as Double
  Dim vv(3) as Double
  Dim r(3) as Double
  Dim v(3) as Double
  Dim sa as Variant
  Dim rdist as Variant
  Dim vdist as Variant
  Dim rerr as Variant
  Dim verr as Variant
  Dim cnt2 as Integer
  
  rerr = 0
  verr = 0
  cnt2 = 0
  
  i = -1
  sht = ThisComponent.Sheets.getByName("tcppver.out")
  sht2 = ThisComponent.Sheets.getByName("TestSGP4")

  Do while(Len(sht.getCellByPosition(0,row).String) > 1)
    str = sht.getCellByPosition(0,row).String
    if(InStr(str,"xx")) then
      i = i + 1
      tl = tles(i)
    else
      str = ws(str)
      sa = Split(str," ") 
      mins = CDbl(sa(0))
      rv(0) = CDbl(sa(1))
      rv(1) = CDbl(sa(2))
      rv(2) = CDbl(sa(3))
      vv(0) = CDbl(sa(4))
      vv(1) = CDbl(sa(5))
      vv(2) = CDbl(sa(6))
      
      getRV(tl,mins,r,v)
      
      rdist = dist(r,rv)
      vdist = dist(v,vv)
      rerr = rerr + rdist
      verr = verr + vdist
      cnt2 = cnt2 + 1    
      if(rdist > 1e-4 OR vdist > 1e-6) then
        sht2.getCellByPosition(0,row+3).String = str
        sht2.getCellByPosition(1,row+3).setValue(rerr)
        sht2.getCellByPosition(2,row+3).setValue(verr)
      end if
    end if
    row = row + 1
  Loop

  rerr = rerr / cnt2
  verr = verr / cnt2
  
  rerr = rerr * 1e6
  verr = verr * 1e6
  sht2.getCellByPosition(0,0).String = "Typical errors"
  sht2.getCellByPosition(0,1).setValue(rerr)
  sht2.getCellByPosition(0,2).setValue(verr)
  sht2.getCellByPosition(1,1).String = "mm"
  sht2.getCellByPosition(1,2).String = "mm/s"
  
end sub


REM runs the TLE verification tests
Sub Main

  Dim tles(100) as Variant
  Dim numTLEs as Integer

  numTLEs = loadTLES(tles)

  call runVer(tles,numTLEs)

End Sub
