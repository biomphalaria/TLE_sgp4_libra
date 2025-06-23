REM  *****  TLE types and functions *****
Type TLE
    rec as Variant
    line1 as String
    line2 as String
    intlid as String
    objectID as String
    epoch as Date
    ndot as Double
    nddot as Double
    bstar as Double
    elnum as Integer
    incDeg as Double
    raanDeg as Double
    ecc as Double
    argpDeg as Double
    maDeg as Double
    n as Double
    revnum as Long
    sgp4Error as Integer
End Type

REM construct an empty TLE
function newTLE
  Dim tl as new TLE
  tl.rec =  newElsetRec()
  newTLE = tl
end function

REM construct a TLE using the supplied lines
function parseTLE(line1 as String, line2 as String)
  Dim tl as new TLE
  tl.rec =  newElsetRec()
  call parseLines(tl,line1,line2)
  parseTLE = tl
end function

REM fill r and v vectos for specific datetime
sub getRVForDate(tl as TLE, dt as Date, r, v)
  Dim diff as Double
  
  diff = dt - tl.epoch
  diff = diff * 1440.0
  
  call getRV(tl,diff,r,v)
end sub

REM fill r and v vectors for specific time after epoch
sub getRV(tl as TLE, minutesAfterEpoch as Double, r, v)
  tl.rec.error = 0
  call sgp4.sgp4(tl.rec,minutesAfterEpoch,r,v)
  tl.sgp4Error = tl.rec.error
end sub

REM parse the lines into the provided TLE
sub parseLines(tl as TLE, line1 as String, line2 as String)
  Dim texp as Double
  tl.line1 = line1
  tl.line2 = line2
  tl.rec.whichconst=sgp4.wgs72
  
  tl.objectID = Trim(Mid(line1,2,5))
  
  tl.rec.classification = Mid(line1,8,1)
 '     #     //          1         2         3         4         5         6
 '          #//0123456789012345678901234567890123456789012345678901234567890123456789
 '   #//line1="1 00005U 58002B   00179.78495062  .00000023  00000-0  28098-4 0  4753"
 '   #//line2="2 00005  34.2682 348.7242 1859667 331.7664  19.3264 10.82419157413667"

    tl.intlid = Trim(Mid(line1,10,8))
    
    tl.epoch = parseEpoch(tl,Trim(Mid(line1,19,14)))
    
    tl.ndot = gdi(Mid(line1,34,1),line1,35,44)

    tl.nddot = gdi(Mid(line1,45,1),line1,45,50)
    texp = Val(Mid(line1,51,2))
    tl.nddot = tl.nddot *(10^texp)

    tl.bstar = gdi(Mid(line1,54,1),line1,54,59)
    texp = Val(Mid(line1,60,2))
    tl.bstar = tl.bstar * (10^texp)
    
    tl.elnum = Int(gd(line1,64,68))

    tl.incDeg = gd(line2,8,16)
    tl.raanDeg = gd(line2,17,25)
    tl.ecc = gdi("+",line2,26,33)
    tl.argpDeg = gd(line2,34,42)
    tl.maDeg = gd(line2,43,51)

    tl.n = gd(line2,52,63)

    tl.revnum = gd(line2,63,68)

  call setValsToRec(tl)
  
end sub

sub setValsToRec(tl as TLE)
  dim deg2rad as Double
  dim xpdotp as Double
  
  deg2rad = PI() / 180.0 
 '//   0.0174532925199433
  xpdotp = 1440.0 / (2.0 * PI())
 '// 229.1831180523293

  tl.rec.elnum = tl.elnum
  tl.rec.revnum = tl.revnum
  tl.rec.satid = tl.objectID
  tl.rec.bstar = tl.bstar
  tl.rec.inclo = tl.incDeg*deg2rad
  tl.rec.nodeo = tl.raanDeg*deg2rad
  tl.rec.argpo = tl.argpDeg*deg2rad
  tl.rec.mo = tl.maDeg*deg2rad
  tl.rec.ecco = tl.ecc
  tl.rec.no_kozai = tl.n/xpdotp
  tl.rec.ndot = tl.ndot / (xpdotp*1440.0)
  tl.rec.nddot = tl.nddot / (xpdotp*1440.0*1440.0)

  call sgp4.sgp4init("a",tl.rec)
end sub

function parseEpoch(tl as TLE, str as String)
   Dim year as Integer
   Dim doy as Integer
   Dim mon as Integer
   Dim dy as Integer
   Dim hr as Integer
   Dim mn as Integer
   Dim sec as Double
   Dim sc as Integer
   Dim dfrac as Double
   Dim ts as String
   Dim td as Date
   Dim jdout(2) as Double
   
   year = Val(Trim(Mid(str,1,2)))
   tl.rec.epochyr = year
   if year > 56 then
     year = year + 1900
   else
     year = year + 2000
   end if
   
   doy = Val(Trim(Mid(str,3,3)))
   ts = "0"+Mid(str,6)
   dfrac = Val(ts)
   
   tl.rec.epochdays = doy+dfrac
   
   ts = year-1
   ts = "12/31/"+ts
   td = CDate(ts)
   td = td+doy+dfrac
   parseEpoch = td

   
   dfrac = 24.0*dfrac
   hr = Int(dfac)
   dfrac = 60.0*(dfrac-hr)
   mn = Int(dfrac)
   dfrac = 60.0*(dfrac-mn)
   sec = dfrac
   sc = Int(sc)
   
   mon = Month(td)
   dy = Day(td)
   
   call SGP4.jday(year,mon,dy,hr,mn,sec,jdout)
   tl.rec.jdsatepoch = jdout(0)
   tl.rec.jdsatepochF = jdout(1)
end function

function gd(str as String, startn as Integer, endn as Integer)
  gd = gd2(str,startn,endn,0)
end function

function gd2(str as String, startn as Integer, endn as Integer, defVal as Double)
  Dim num as Double
  num = defVal
  endn = endn-startn
  num = Val(Trim(Mid(str,startn+1,endn)))
  gd2 = num
end function

function gdi(sign as String, str as String, startn as Integer, endn as Integer)
  gdi = gdi2(sign,str,startn,endn,0)
end function

function gdi2(sign as String, str as String, startn as Integer, endn as Integer, defVal as Double)
  Dim num as Double
  Dim ts as String
  
  num = defVal
  endn = endn-startn
  
  ts = Mid(str,startn+1,endn)
  num = Val(Trim("0."+ts))
  if sign = "-" then
    num = num * -1.0
  end if
  gdi2 = num
end function

Sub Main
  l1 = "1 00005U 58002B   00179.78495062  .00000023  00000-0  28098-4 0  4753"
  l2 = "2 00005  34.2682 348.7242 1859667 331.7664  19.3264 10.82419157413667" 

  l1 = "1 04632U 70093B   04031.91070959 -.00000084  00000-0  10000-3 0  9955"
  l2 = "2 04632  11.4628 273.1101 1450506 207.6000 143.9350  1.20231981 44145"
  
  Dim tl as Variant
  
  tl = parseTLE(l1,l2)
  
  ThisComponent.Sheets.getByName("Scratch").getCellByPosition(0,0).setString(tl.line1)
  ThisComponent.Sheets.getByName("Scratch").getCellByPosition(0,1).setString(tl.line2)
  ThisComponent.Sheets.getByName("Scratch").getCellByPosition(0,2).setValue(tl.epoch)
  
  Dim r(3) as Double
  Dim v(3) as Double
  
  call getRV(tl,-5184,r,v)

  ThisComponent.Sheets.getByName("Scratch").getCellByPosition(0,3).setValue(r(0))
  ThisComponent.Sheets.getByName("Scratch").getCellByPosition(0,4).setValue(r(1))
  ThisComponent.Sheets.getByName("Scratch").getCellByPosition(0,5).setValue(r(2))
  ThisComponent.Sheets.getByName("Scratch").getCellByPosition(0,6).setValue(v(0))
  ThisComponent.Sheets.getByName("Scratch").getCellByPosition(0,7).setValue(v(1))
  ThisComponent.Sheets.getByName("Scratch").getCellByPosition(0,8).setValue(v(2))
  
  
End Sub
