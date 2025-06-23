REM  *****  SGP4 Types and Functions  *****

Public Const pi as Double = 3.14159265358979
Public Const twopi as Double = 6.28318530717959
Public Const deg2rad as Double = pi / 180.0

Public Const wgs72old As Integer = 1
Public Const wgs72 As Integer = 2
Public Const wgs84 As Integer = 3

Public Type ElsetRec
	whichconst As Integer
	satid As String
	epochyr As Integer
	epochtynumrev As Long
	error As Integer
	operationmode As String
	init As String
	method As String
	a As Double
	altp As Double
	alta As Double
	epochdays As Double
	jdsatepoch As Double
	jdsatepochF As Double
	nddot As Double
	ndot As Double
	bstar As Double
	rcse As Double
	inclo As Double
	nodeo As Double
	ecco As Double
	argpo As Double
	mo As Double
	no_kozai As Double
	classification As String
	intldesg As String
	ephtype As Integer
	elnum As Long
	revnum As Long
	no_unkozai As Double
	am As Double
	em As Double
	im As Double
	Om As Double
	oom As Double  &apos;replaces om 
	mm As Double
	nm As Double
	t As Double
	tumin As Double
	mu As Double
	radiusearthkm As Double
	xke As Double
	j2 As Double
	j3 As Double
	j4 As Double
	j3oj2 As Double
	dia_mm As Long
	period_sec As Double
	active As String
	not_orbital As String
	rcs_m2 As Double
	ep As Double
	inclp As Double
	nodep As Double
	argpp As Double
	mp As Double
	isimp As Integer
	aycof As Double
	con41 As Double
	cc1 As Double
	cc4 As Double
	cc5 As Double
	d2 As Double
	d3 As Double
	d4 As Double
	delmo As Double
	eta As Double
	argpdot As Double
	omgcof As Double
	sinmao As Double
	t2cof As Double
	t3cof As Double
	t4cof As Double
	t5cof As Double
	x1mth2 As Double
	x7thm1 As Double
	mdot As Double
	nodedot As Double
	xlcof As Double
	xmcof As Double
	nodecf As Double
	irez As Integer
	d2201 As Double
	d2211 As Double
	d3210 As Double
	d3222 As Double
	d4410 As Double
	d4422 As Double
	d5220 As Double
	d5232 As Double
	d5421 As Double
	d5433 As Double
	dedt As Double
	del1 As Double
	del2 As Double
	del3 As Double
	didt As Double
	dmdt As Double
	dnodt As Double
	domdt As Double
	e3 As Double
	ee2 As Double
	peo As Double
	pgho As Double
	pho As Double
	pinco As Double
	plo As Double
	se2 As Double
	se3 As Double
	sgh2 As Double
	sgh3 As Double
	sgh4 As Double
	sh2 As Double
	sh3 As Double
	si2 As Double
	si3 As Double
	sl2 As Double
	sl3 As Double
	sl4 As Double
	gsto As Double
	xfact As Double
	xgh2 As Double
	xgh3 As Double
	xgh4 As Double
	xh2 As Double
	xh3 As Double
	xi2 As Double
	xi3 As Double
	xl2 As Double
	xl3 As Double
	xl4 As Double
	xlamo As Double
	zmol As Double
	zmos As Double
	atime As Double
	xli As Double
	xni As Double
	snodm As Double
	cnodm As Double
	sinim As Double
	cosim As Double
	sinomm As Double
	cosomm As Double
	day As Double
	emsq As Double
	gam As Double
	rtemsq As Double
	s1 As Double
	s2 As Double
	s3 As Double
	s4 As Double
	s5 As Double
	s6 As Double
	s7 As Double
	ss1 As Double
	ss2 As Double
	ss3 As Double
	ss4 As Double
	ss5 As Double
	ss6 As Double
	ss7 As Double
	sz1 As Double
	sz2 As Double
	sz3 As Double
	sz11 As Double
	sz12 As Double
	sz13 As Double
	sz21 As Double
	sz22 As Double
	sz23 As Double
	sz31 As Double
	sz32 As Double
	sz33 As Double
	z1 As Double
	z2 As Double
	z3 As Double
	z11 As Double
	z12 As Double
	z13 As Double
	z21 As Double
	z22 As Double
	z23 As Double
	z31 As Double
	z32 As Double
	z33 As Double
	argpm As Double
	inclm As Double
	nodem As Double
	dndt As Double
	eccsq As Double
	ainv As Double
	ao As Double
	con42 As Double
	cosio As Double
	cosio2 As Double
	omeosq As Double
	posq As Double
	rp As Double
	rteosq As Double
	sinio As Double
End Type


function newElsetRec
  Dim er as new ElsetRec
  newElsetRec = er
end function


sub dpper(e3 as Double, ee2 as Double, peo as Double, pgho as Double, pho as Double, _
        pinco as Double, plo as Double, se2 as Double, se3 as Double, sgh2 as Double, _
        sgh3 as Double, sgh4 as Double, sh2 as Double, sh3 as Double, si2 as Double, _
        si3 as Double, sl2 as Double, sl3 as Double, sl4 as Double, t as Double, _
        xgh2 as Double, xgh3 as Double, xgh4 as Double, xh2 as Double, xh3 as Double, _
        xi2 as Double, xi3 as Double, xl2 as Double, xl3 as Double, xl4 as Double, _
        zmol as Double, zmos as Double, _
        init as String, rec as ElsetRec, opsmode as String )

        &apos;/* --------------------- local variables ------------------------ */
        Dim alfdp, betdp, cosip, cosop, dalf, dbet, dls, _
            f2, f3, pe, pgh, ph, pinc, pl, _
            sel, ses, sghl, sghs, shll, shs, sil, _
            sinip, sinop, sinzf, sis, sll, sls, xls, _
            xnoh, zf, zm, zel, zes, znl, zns as Double

        
        &apos;/* ---------------------- constants ----------------------------- */
        zns = 1.19459e-5
        zes = 0.01675
        znl = 1.5835218e-4
        zel = 0.05490

        &apos;/* --------------- calculate time varying periodics ----------- */
        zm = zmos + zns * t
        &apos;// be sure that the initial call has time set to zero
        if (init = &quot;y&quot;) then
            zm = zmos
        end if
        zf = zm + 2.0 * zes * sin(zm)
        sinzf = sin(zf)
        f2 = 0.5 * sinzf * sinzf - 0.25
        f3 = -0.5 * sinzf * cos(zf)
        ses = se2* f2 + se3 * f3
        sis = si2 * f2 + si3 * f3
        sls = sl2 * f2 + sl3 * f3 + sl4 * sinzf
        sghs = sgh2 * f2 + sgh3 * f3 + sgh4 * sinzf
        shs = sh2 * f2 + sh3 * f3
        zm = zmol + znl * t
        if (init = &quot;y&quot;) then
            zm = zmol
        end if
        zf = zm + 2.0 * zel * sin(zm)
        sinzf = sin(zf)
        f2 = 0.5 * sinzf * sinzf - 0.25
        f3 = -0.5 * sinzf * cos(zf)
        sel = ee2 * f2 + e3 * f3
        sil = xi2 * f2 + xi3 * f3
        sll = xl2 * f2 + xl3 * f3 + xl4 * sinzf
        sghl = xgh2 * f2 + xgh3 * f3 + xgh4 * sinzf
        shll = xh2 * f2 + xh3 * f3
        pe = ses + sel
        pinc = sis + sil
        pl = sls + sll
        pgh = sghs + sghl
        ph = shs + shll

        if (init = &quot;n&quot;) then
        
            pe = pe - peo
            pinc = pinc - pinco
            pl = pl - plo
            pgh = pgh - pgho
            ph = ph - pho
            rec.inclp = rec.inclp + pinc
            rec.ep = rec.ep + pe
            sinip = sin(rec.inclp)
            cosip = cos(rec.inclp)

            if (rec.inclp &gt;= 0.2) then
            
                ph = ph / sinip
                pgh = pgh - cosip * ph
                rec.argpp = rec.argpp + pgh
                rec.nodep = rec.nodep + ph
                rec.mp = rec.mp + pl
            
            else
            
                sinop = sin(rec.nodep)
                cosop = cos(rec.nodep)
                alfdp = sinip * sinop
                betdp = sinip * cosop
                dalf = ph * cosop + pinc * cosip * sinop
                dbet = -ph * sinop + pinc * cosip * cosop
                alfdp = alfdp + dalf
                betdp = betdp + dbet
                rec.nodep = fmod(rec.nodep, twopi)
                if ((rec.nodep &lt; 0.0) AND (opsmode = &quot;a&quot;)) then
                    rec.nodep = rec.nodep + twopi
                end if
                xls = rec.mp + rec.argpp + cosip * rec.nodep
                dls = pl + pgh - pinc * rec.nodep * sinip
                xls = xls + dls
                xls = fmod(xls,twopi)
                xnoh = rec.nodep
                rec.nodep = Atn2(alfdp, betdp)
                if ((rec.nodep &lt; 0.0) AND (opsmode = &quot;a&quot;)) then
                    rec.nodep = rec.nodep + twopi
                end if
                if (abs(xnoh - rec.nodep) &gt; pi) then
                    if (rec.nodep &lt; xnoh) then
                        rec.nodep = rec.nodep + twopi
                    else
                        rec.nodep = rec.nodep - twopi
                    end if
                end if
                rec.mp = rec.mp + pl
                rec.argpp = xls - rec.mp - cosip * rec.nodep
            end if
        end if  &apos; // if init == &apos;n&apos;

end sub &apos;// dpper


sub dscom( epoch as Double, ep as Double, argpp as Double, tc as Double, inclp as Double, _
           nodep as Double, np as Double, rec as ElsetRec)
           
        &apos;/* -------------------------- constants ------------------------- */
        Const zes as Double = 0.01675
        Const zel = 0.05490
        Const c1ss = 2.9864797e-6
        Const c1l = 4.7968065e-7
        Const zsinis = 0.39785416
        Const zcosis = 0.91744867
        Const zcosgs = 0.1945905
        Const zsings = -0.98088458

        &apos;/* --------------------- local variables ------------------------ */
        Dim lsflg as Integer
        Dim a1, a2, a3, a4, a5, a6, a7, _
            a8, a9, a10, betasq, cc, ctem, stem, _
            x1, x2, x3, x4, x5, x6, x7, _
            x8, xnodce, xnoi, zcosg, zcosgl, zcosh, zcoshl, _
            zcosi, zcosil, zsing, zsingl, zsinh, zsinhl, zsini, _
            zsinil, zx, zy as Double

        rec.nm = np
        rec.em = ep
        rec.snodm = sin(nodep)
        rec.cnodm = cos(nodep)
        rec.sinomm = sin(argpp)
        rec.cosomm = cos(argpp)
        rec.sinim = sin(inclp)
        rec.cosim = cos(inclp)
        rec.emsq = rec.em * rec.em
        betasq = 1.0 - rec.emsq
        rec.rtemsq = Sqr(betasq)

        &apos;/* ----------------- initialize lunar solar terms --------------- */
        rec.peo = 0.0
        rec.pinco = 0.0
        rec.plo = 0.0
        rec.pgho = 0.0
        rec.pho = 0.0
        rec.day = epoch + 18261.5 + tc / 1440.0
        xnodce = fmod(4.5236020 - 9.2422029e-4 * rec.day, twopi)
        stem = sin(xnodce)
        ctem = cos(xnodce)
        zcosil = 0.91375164 - 0.03568096 * ctem
        zsinil = Sqr(1.0 - zcosil * zcosil)
        zsinhl = 0.089683511 * stem / zsinil
        zcoshl = Sqr(1.0 - zsinhl * zsinhl)
        rec.gam = 5.8351514 + 0.0019443680 * rec.day
        zx = 0.39785416 * stem / zsinil
        zy = zcoshl * ctem + 0.91744867 * zsinhl * stem
        zx = Atn2(zx, zy)
        zx = rec.gam + zx - xnodce
        zcosgl = cos(zx)
        zsingl = sin(zx)

        &apos;/* ------------------------- do solar terms --------------------- */
        zcosg = zcosgs
        zsing = zsings
        zcosi = zcosis
        zsini = zsinis
        zcosh = rec.cnodm
        zsinh = rec.snodm
        cc = c1ss
        xnoi = 1.0 / rec.nm

        For lsflg = 1 to 2 Step 1
            a1 = zcosg * zcosh + zsing * zcosi * zsinh
            a3 = -zsing * zcosh + zcosg * zcosi * zsinh
            a7 = -zcosg * zsinh + zsing * zcosi * zcosh
            a8 = zsing * zsini
            a9 = zsing * zsinh + zcosg * zcosi * zcosh
            a10 = zcosg * zsini
            a2 = rec.cosim * a7 + rec.sinim * a8
            a4 = rec.cosim * a9 + rec.sinim * a10
            a5 = -rec.sinim * a7 + rec.cosim * a8
            a6 = -rec.sinim * a9 + rec.cosim * a10

            x1 = a1 * rec.cosomm + a2 * rec.sinomm
            x2 = a3 * rec.cosomm + a4 * rec.sinomm
            x3 = -a1 * rec.sinomm + a2 * rec.cosomm
            x4 = -a3 * rec.sinomm + a4 * rec.cosomm
            x5 = a5 * rec.sinomm
            x6 = a6 * rec.sinomm
            x7 = a5 * rec.cosomm
            x8 = a6 * rec.cosomm

            rec.z31 = 12.0 * x1 * x1 - 3.0 * x3 * x3
            rec.z32 = 24.0 * x1 * x2 - 6.0 * x3 * x4
            rec.z33 = 12.0 * x2 * x2 - 3.0 * x4 * x4
            rec.z1 = 3.0 *  (a1 * a1 + a2 * a2) + rec.z31 * rec.emsq
            rec.z2 = 6.0 *  (a1 * a3 + a2 * a4) + rec.z32 * rec.emsq
            rec.z3 = 3.0 *  (a3 * a3 + a4 * a4) + rec.z33 * rec.emsq
            rec.z11 = -6.0 * a1 * a5 + rec.emsq *  (-24.0 * x1 * x7 - 6.0 * x3 * x5)
            rec.z12 = -6.0 *  (a1 * a6 + a3 * a5) + rec.emsq * _
                (-24.0 * (x2 * x7 + x1 * x8) - 6.0 * (x3 * x6 + x4 * x5))
            rec.z13 = -6.0 * a3 * a6 + rec.emsq * (-24.0 * x2 * x8 - 6.0 * x4 * x6)
            rec.z21 = 6.0 * a2 * a5 + rec.emsq * (24.0 * x1 * x5 - 6.0 * x3 * x7)
            rec.z22 = 6.0 *  (a4 * a5 + a2 * a6) + rec.emsq * _
                (24.0 * (x2 * x5 + x1 * x6) - 6.0 * (x4 * x7 + x3 * x8))
            rec.z23 = 6.0 * a4 * a6 + rec.emsq * (24.0 * x2 * x6 - 6.0 * x4 * x8)
            rec.z1 = rec.z1 + rec.z1 + betasq * rec.z31
            rec.z2 = rec.z2 + rec.z2 + betasq * rec.z32
            rec.z3 = rec.z3 + rec.z3 + betasq * rec.z33
            rec.s3 = cc * xnoi
            rec.s2 = -0.5 * rec.s3 / rec.rtemsq
            rec.s4 = rec.s3 * rec.rtemsq
            rec.s1 = -15.0 * rec.em * rec.s4
            rec.s5 = x1 * x3 + x2 * x4
            rec.s6 = x2 * x3 + x1 * x4
            rec.s7 = x2 * x4 - x1 * x3

            &apos;/* ----------------------- do lunar terms ------------------- */
            if (lsflg = 1) then
            
                rec.ss1 = rec.s1
                rec.ss2 = rec.s2
                rec.ss3 = rec.s3
                rec.ss4 = rec.s4
                rec.ss5 = rec.s5
                rec.ss6 = rec.s6
                rec.ss7 = rec.s7
                rec.sz1 = rec.z1
                rec.sz2 = rec.z2
                rec.sz3 = rec.z3
                rec.sz11 = rec.z11
                rec.sz12 = rec.z12
                rec.sz13 = rec.z13
                rec.sz21 = rec.z21
                rec.sz22 = rec.z22
                rec.sz23 = rec.z23
                rec.sz31 = rec.z31
                rec.sz32 = rec.z32
                rec.sz33 = rec.z33
                zcosg = zcosgl
                zsing = zsingl
                zcosi = zcosil
                zsini = zsinil
                zcosh = zcoshl * rec.cnodm + zsinhl * rec.snodm
                zsinh = rec.snodm * zcoshl - rec.cnodm * zsinhl
                cc = c1l
            end if
            
        Next lsflg

        rec.zmol = fmod(4.7199672 + 0.22997150  * rec.day - rec.gam, twopi)
        rec.zmos = fmod(6.2565837 + 0.017201977 * rec.day, twopi)

       &apos; /* ------------------------ do solar terms ---------------------- */
        rec.se2 = 2.0 * rec.ss1 * rec.ss6
        rec.se3 = 2.0 * rec.ss1 * rec.ss7
        rec.si2 = 2.0 * rec.ss2 * rec.sz12
        rec.si3 = 2.0 * rec.ss2 * (rec.sz13 - rec.sz11)
        rec.sl2 = -2.0 * rec.ss3 * rec.sz2
        rec.sl3 = -2.0 * rec.ss3 * (rec.sz3 - rec.sz1)
        rec.sl4 = -2.0 * rec.ss3 * (-21.0 - 9.0 * rec.emsq) * zes
        rec.sgh2 = 2.0 * rec.ss4 * rec.sz32
        rec.sgh3 = 2.0 * rec.ss4 * (rec.sz33 - rec.sz31)
        rec.sgh4 = -18.0 * rec.ss4 * zes
        rec.sh2 = -2.0 * rec.ss2 * rec.sz22
        rec.sh3 = -2.0 * rec.ss2 * (rec.sz23 - rec.sz21)

       &apos; /* ------------------------ do lunar terms ---------------------- */
        rec.ee2 = 2.0 * rec.s1 * rec.s6
        rec.e3 = 2.0 * rec.s1 * rec.s7
        rec.xi2 = 2.0 * rec.s2 * rec.z12
        rec.xi3 = 2.0 * rec.s2 * (rec.z13 - rec.z11)
        rec.xl2 = -2.0 * rec.s3 * rec.z2
        rec.xl3 = -2.0 * rec.s3 * (rec.z3 - rec.z1)
        rec.xl4 = -2.0 * rec.s3 * (-21.0 - 9.0 * rec.emsq) * zel
        rec.xgh2 = 2.0 * rec.s4 * rec.z32
        rec.xgh3 = 2.0 * rec.s4 * (rec.z33 - rec.z31)
        rec.xgh4 = -18.0 * rec.s4 * zel
        rec.xh2 = -2.0 * rec.s2 * rec.z22
        rec.xh3 = -2.0 * rec.s2 * (rec.z23 - rec.z21)

end sub &apos;  // dscom



sub dsinit( tc as Double, xpidot as Double, rec as ElsetRec)
       Dim ainv2, aonv, cosisq, eoc, f220, f221, f311, _
            f321, f322, f330, f441, f442, f522, f523, _
            f542, f543, g200, g201, g211, g300, g310, _
            g322, g410, g422, g520, g521, g532, g533, _
            ses, sgs, sghl, sghs, shs, shll, sis, _
            sini2, sls, temp, temp1, theta, xno2, q22, _ 
            q31, q33, root22, root44, root54, rptim, root32, _
            root52, x2o3, znl, emo, zns, emsqo as Double

        aonv = 0.0
        
        q22 = 1.7891679e-6
        q31 = 2.1460748e-6
        q33 = 2.2123015e-7
        root22 = 1.7891679e-6
        root44 = 7.3636953e-9
        root54 = 2.1765803e-9
        rptim = 4.37526908801129966e-3 &apos;// this equates to 7.29211514668855e-5 rad/sec
        root32 = 3.7393792e-7
        root52 = 1.1428639e-7
        x2o3 = 2.0 / 3.0
        znl = 1.5835218e-4
        zns = 1.19459e-5

      
        &apos;/* -------------------- deep space initialization ------------ */
        rec.irez = 0
        if ((rec.nm &lt; 0.0052359877) AND (rec.nm &gt; 0.0034906585)) then
            rec.irez = 1
        end if
        if ((rec.nm &gt;= 8.26e-3) AND (rec.nm &lt;= 9.24e-3) AND (rec.em &gt;= 0.5)) then
            rec.irez = 2
        end if

        &apos;/* ------------------------ do solar terms ------------------- */
        ses = rec.ss1 * zns * rec.ss5
        sis = rec.ss2 * zns * (rec.sz11 + rec.sz13)
        sls = -zns * rec.ss3 * (rec.sz1 + rec.sz3 - 14.0 - 6.0 * rec.emsq)
        sghs = rec.ss4 * zns * (rec.sz31 + rec.sz33 - 6.0)
        shs = -zns * rec.ss2 * (rec.sz21 + rec.sz23)
        &apos;// sgp4fix for 180 deg incl
        if ((rec.inclm &lt; 5.2359877e-2) OR (rec.inclm &gt; pi - 5.2359877e-2)) then
            shs = 0.0
        end if
        
        if (rec.sinim &lt;&gt; 0.0) then
            shs = shs / rec.sinim
        end if
        sgs = sghs - rec.cosim * shs

        &apos;/* ------------------------- do lunar terms ------------------ */
        rec.dedt = ses + rec.s1 * znl * rec.s5
        rec.didt = sis + rec.s2 * znl * (rec.z11 + rec.z13)
        rec.dmdt = sls - znl * rec.s3 * (rec.z1 + rec.z3 - 14.0 - 6.0 * rec.emsq)
        sghl = rec.s4 * znl * (rec.z31 + rec.z33 - 6.0)
        shll = -znl * rec.s2 * (rec.z21 + rec.z23)
        &apos;// sgp4fix for 180 deg incl
        if ((rec.inclm &lt; 5.2359877e-2) OR (rec.inclm &gt; pi - 5.2359877e-2)) then
            shll = 0.0
        end if
        
        rec.domdt = sgs + sghl
        rec.dnodt = shs
        
        if (rec.sinim &lt;&gt; 0.0) then
            rec.domdt = rec.domdt - rec.cosim / rec.sinim * shll
            rec.dnodt = rec.dnodt + shll / rec.sinim
        end if

        &apos;/* ----------- calculate deep space resonance effects -------- */
        rec.dndt = 0.0
        theta = fmod(rec.gsto + tc * rptim, twopi)
        rec.em = rec.em + rec.dedt * rec.t
        rec.inclm = rec.inclm + rec.didt * rec.t
        rec.argpm = rec.argpm + rec.domdt * rec.t
        rec.nodem = rec.nodem + rec.dnodt * rec.t
        rec.mm = rec.mm + rec.dmdt * rec.t
        
        &apos;/* -------------- initialize the resonance terms ------------- */
        if (rec.irez &lt;&gt; 0) then
        
            aonv = pow(rec.nm / rec.xke, x2o3)

            &apos;/* ---------- geopotential resonance for 12 hour orbits ------ */
            if (rec.irez = 2) then
            
                cosisq = rec.cosim * rec.cosim
                emo = rec.em
                rec.em = rec.ecco
                emsqo = rec.emsq
                rec.emsq = rec.eccsq
                eoc = rec.em * rec.emsq
                g201 = -0.306 - (rec.em - 0.64) * 0.440

                if (rec.em &lt;= 0.65) then
                
                    g211 = 3.616 - 13.2470 * rec.em + 16.2900 * rec.emsq
                    g310 = -19.302 + 117.3900 * rec.em - 228.4190 * rec.emsq + 156.5910 * eoc
                    g322 = -18.9068 + 109.7927 * rec.em - 214.6334 * rec.emsq + 146.5816 * eoc
                    g410 = -41.122 + 242.6940 * rec.em - 471.0940 * rec.emsq + 313.9530 * eoc
                    g422 = -146.407 + 841.8800 * rec.em - 1629.014 * rec.emsq + 1083.4350 * eoc
                    g520 = -532.114 + 3017.977 * rec.em - 5740.032 * rec.emsq + 3708.2760 * eoc
                
                else
                
                    g211 = -72.099 + 331.819 * rec.em - 508.738 * rec.emsq + 266.724 * eoc
                    g310 = -346.844 + 1582.851 * rec.em - 2415.925 * rec.emsq + 1246.113 * eoc
                    g322 = -342.585 + 1554.908 * rec.em - 2366.899 * rec.emsq + 1215.972 * eoc
                    g410 = -1052.797 + 4758.686 * rec.em - 7193.992 * rec.emsq + 3651.957 * eoc
                    g422 = -3581.690 + 16178.110 * rec.em - 24462.770 * rec.emsq + 12422.520 * eoc
                    if (rec.em &gt; 0.715) then
                        g520 = -5149.66 + 29936.92 * rec.em - 54087.36 * rec.emsq + 31324.56 * eoc
                    else
                        g520 = 1464.74 - 4664.75 * rec.em + 3763.64 * rec.emsq
                    end if
                end if
                
                if (rec.em &lt; 0.7) then
                
                    g533 = -919.22770 + 4988.6100 * rec.em - 9064.7700 * rec.emsq + 5542.21  * eoc
                    g521 = -822.71072 + 4568.6173 * rec.em - 8491.4146 * rec.emsq + 5337.524 * eoc
                    g532 = -853.66600 + 4690.2500 * rec.em - 8624.7700 * rec.emsq + 5341.4  * eoc
                
                else
                
                    g533 = -37995.780 + 161616.52 * rec.em - 229838.20 * rec.emsq + 109377.94 * eoc
                    g521 = -51752.104 + 218913.95 * rec.em - 309468.16 * rec.emsq + 146349.42 * eoc
                    g532 = -40023.880 + 170470.89 * rec.em - 242699.48 * rec.emsq + 115605.82 * eoc
                end if

                sini2 = rec.sinim * rec.sinim
                f220 = 0.75 * (1.0 + 2.0 * rec.cosim + cosisq)
                f221 = 1.5 * sini2
                f321 = 1.875 * rec.sinim  *  (1.0 - 2.0 * rec.cosim - 3.0 * cosisq)
                f322 = -1.875 * rec.sinim  *  (1.0 + 2.0 * rec.cosim - 3.0 * cosisq)
                f441 = 35.0 * sini2 * f220
                f442 = 39.3750 * sini2 * sini2
                f522 = 9.84375 * rec.sinim * (sini2 * (1.0 - 2.0 * rec.cosim - 5.0 * cosisq) + _
                    0.33333333 * (-2.0 + 4.0 * rec.cosim + 6.0 * cosisq))
                f523 = rec.sinim * (4.92187512 * sini2 * (-2.0 - 4.0 * rec.cosim + _
                    10.0 * cosisq) + 6.56250012 * (1.0 + 2.0 * rec.cosim - 3.0 * cosisq))
                f542 = 29.53125 * rec.sinim * (2.0 - 8.0 * rec.cosim + cosisq * _
                    (-12.0 + 8.0 * rec.cosim + 10.0 * cosisq))
                f543 = 29.53125 * rec.sinim * (-2.0 - 8.0 * rec.cosim + cosisq * _
                    (12.0 + 8.0 * rec.cosim - 10.0 * cosisq))
                xno2 = rec.nm * rec.nm
                ainv2 = aonv * aonv
                temp1 = 3.0 * xno2 * ainv2
                temp = temp1 * root22
                rec.d2201 = temp * f220 * g201
                rec.d2211 = temp * f221 * g211
                temp1 = temp1 * aonv
                temp = temp1 * root32
                rec.d3210 = temp * f321 * g310
                rec.d3222 = temp * f322 * g322
                temp1 = temp1 * aonv
                temp = 2.0 * temp1 * root44
                rec.d4410 = temp * f441 * g410
                rec.d4422 = temp * f442 * g422
                temp1 = temp1 * aonv
                temp = temp1 * root52
                rec.d5220 = temp * f522 * g520
                rec.d5232 = temp * f523 * g532
                temp = 2.0 * temp1 * root54
                rec.d5421 = temp * f542 * g521
                rec.d5433 = temp * f543 * g533
                rec.xlamo = fmod(rec.mo + rec.nodeo + rec.nodeo - theta - theta, twopi)
                rec.xfact = rec.mdot + rec.dmdt + 2.0 * (rec.nodedot + rec.dnodt - rptim) - rec.no_unkozai
                rec.em = emo
                rec.emsq = emsqo
            end if

            &apos;/* ---------------- synchronous resonance terms -------------- */
            if (rec.irez = 1) then
            
                g200 = 1.0 + rec.emsq * (-2.5 + 0.8125 * rec.emsq)
                g310 = 1.0 + 2.0 * rec.emsq
                g300 = 1.0 + rec.emsq * (-6.0 + 6.60937 * rec.emsq)
                f220 = 0.75 * (1.0 + rec.cosim) * (1.0 + rec.cosim)
                f311 = 0.9375 * rec.sinim * rec.sinim * (1.0 + 3.0 * rec.cosim) - 0.75 * (1.0 + rec.cosim)
                f330 = 1.0 + rec.cosim
                f330 = 1.875 * f330 * f330 * f330
                rec.del1 = 3.0 * rec.nm * rec.nm * aonv * aonv
                rec.del2 = 2.0 * rec.del1 * f220 * g200 * q22
                rec.del3 = 3.0 * rec.del1 * f330 * g300 * q33 * aonv
                rec.del1 = rec.del1 * f311 * g310 * q31 * aonv
                rec.xlamo = fmod(rec.mo + rec.nodeo + rec.argpo - theta, twopi)
                rec.xfact = rec.mdot + xpidot - rptim + rec.dmdt + rec.domdt + rec.dnodt - rec.no_unkozai
            end if

            &apos;/* ------------ for sgp4, initialize the integrator ---------- */
            rec.xli = rec.xlamo
            rec.xni = rec.no_unkozai
            rec.atime = 0.0
            rec.nm = rec.no_unkozai + rec.dndt
        end if

end sub   &apos;// dsinit

sub dspace(tc as Double, rec as ElsetRec)

        Dim iretn as Integer
        Dim delt, ft, theta, x2li, x2omi, xl, xldot, xnddt, xndt, xomi, g22, g32, _
            g44, g52, g54, fasx2, fasx4, fasx6, rptim, step2, stepn, stepp as Double
        
        xndt = 0
        xnddt = 0
        xldot = 0
        
        fasx2 = 0.13130908
        fasx4 = 2.8843198
        fasx6 = 0.37448087
        g22 = 5.7686396
        g32 = 0.95240898
        g44 = 1.8014998
        g52 = 1.0508330
        g54 = 4.4108898
        rptim = 4.37526908801129966e-3 &apos;// this equates to 7.29211514668855e-5 rad/sec
        stepp = 720.0
        stepn = -720.0
        step2 = 259200.0

        &apos;/* ----------- calculate deep space resonance effects ----------- */
        rec.dndt = 0.0
        theta = fmod(rec.gsto + tc * rptim, twopi)
        rec.em = rec.em + rec.dedt * rec.t

        rec.inclm = rec.inclm + rec.didt * rec.t
        rec.argpm = rec.argpm + rec.domdt * rec.t
        rec.nodem = rec.nodem + rec.dnodt * rec.t
        rec.mm = rec.mm + rec.dmdt * rec.t



        ft = 0.0
        if (rec.irez &lt;&gt; 0) then
        
            &apos;// sgp4fix streamline check
            if ((rec.atime = 0.0) OR (rec.t * rec.atime &lt;= 0.0) OR (abs(rec.t) &lt; abs(rec.atime))) then
            
                rec.atime = 0.0
                rec.xni = rec.no_unkozai
                rec.xli = rec.xlamo
            end if
            &apos;// sgp4fix move check outside loop
            if (rec.t &gt; 0.0) then
                delt = stepp
            else
                delt = stepn
            end if

            iretn = 381 &apos;// added for do loop
            Do while (iretn = 381)
            
                &apos;/* ------------------- dot terms calculated ------------- */
                &apos;/* ----------- near - synchronous resonance terms ------- */
                if (rec.irez &lt;&gt; 2) then
                
                    xndt = rec.del1 * sin(rec.xli - fasx2) + rec.del2 * sin(2.0 * (rec.xli - fasx4)) + _
                            rec.del3 * sin(3.0 * (rec.xli - fasx6))
                    xldot = rec.xni + rec.xfact
                    xnddt = rec.del1 * cos(rec.xli - fasx2) + _
                        2.0 * rec.del2 * cos(2.0 * (rec.xli - fasx4)) + _
                        3.0 * rec.del3 * cos(3.0 * (rec.xli - fasx6))
                    xnddt = xnddt * xldot
                
                else
                
                    &apos;/* --------- near - half-day resonance terms -------- */
                    xomi = rec.argpo + rec.argpdot * rec.atime
                    x2omi = xomi + xomi
                    x2li = rec.xli + rec.xli
                    xndt = rec.d2201 * sin(x2omi + rec.xli - g22) + rec.d2211 * sin(rec.xli - g22) + _
                            rec.d3210 * sin(xomi + rec.xli - g32) + rec.d3222 * sin(-xomi + rec.xli - g32) + _
                            rec.d4410 * sin(x2omi + x2li - g44) + rec.d4422 * sin(x2li - g44) + _
                            rec.d5220 * sin(xomi + rec.xli - g52) + rec.d5232 * sin(-xomi + rec.xli - g52) + _
                            rec.d5421 * sin(xomi + x2li - g54) + rec.d5433 * sin(-xomi + x2li - g54)
                    xldot = rec.xni + rec.xfact
                    xnddt = rec.d2201 * cos(x2omi + rec.xli - g22) + rec.d2211 * cos(rec.xli - g22) + _
                            rec.d3210 * cos(xomi + rec.xli - g32) + rec.d3222 * cos(-xomi + rec.xli - g32) + _
                            rec.d5220 * cos(xomi + rec.xli - g52) + rec.d5232 * cos(-xomi + rec.xli - g52) + _
                        2.0 * (rec.d4410 * cos(x2omi + x2li - g44) + _
                                rec.d4422 * cos(x2li - g44) + rec.d5421 * cos(xomi + x2li - g54) + _
                                rec.d5433 * cos(-xomi + x2li - g54))
                    xnddt = xnddt * xldot
                end if

                &apos;/* ----------------------- integrator ------------------- */
                &apos;// sgp4fix move end checks to end of routine
                if (abs(rec.t - rec.atime) &gt;= stepp) then
                
                    iretn = 381
                
                else &apos;// exit here
                
                    ft = rec.t - rec.atime
                    iretn = 0
                end if

                if (iretn = 381) then
                
                    rec.xli = rec.xli + xldot * delt + xndt * step2
                    rec.xni = rec.xni + xndt * delt + xnddt * step2
                    rec.atime = rec.atime + delt
                end if
            Loop  &apos;// while iretn = 381

            
            rec.nm = rec.xni + xndt * ft + xnddt * ft * ft * 0.5
            xl = rec.xli + xldot * ft + xndt * ft * ft * 0.5
            if (rec.irez &lt;&gt; 1) then
            
                rec.mm = xl - 2.0 * rec.nodem + 2.0 * theta
                rec.dndt = rec.nm - rec.no_unkozai
            
            else
            
                rec.mm = xl - rec.nodem - rec.argpm + theta
                rec.dndt = rec.nm - rec.no_unkozai
            end if
            rec.nm = rec.no_unkozai + rec.dndt
        end if

    end sub  &apos;// dsspace


    sub initl(epoch as Double, rec as ElsetRec)
    
        &apos;/* --------------------- local variables ------------------------ */
        Dim ak, d1, del, adel, po, x2o3, ds70, ts70, tfrac, c1, thgr70, fk5r, c1p2p as Double

        &apos;// sgp4fix use old way of finding gst
        
        &apos;/* ----------------------- earth constants ---------------------- */
        x2o3 = 2.0 / 3.0

        &apos;/* ------------- calculate auxillary epoch quantities ---------- */
        rec.eccsq = rec.ecco * rec.ecco
        rec.omeosq = 1.0 - rec.eccsq
        rec.rteosq = Sqr(rec.omeosq)
        rec.cosio = cos(rec.inclo)
        rec.cosio2 = rec.cosio * rec.cosio

        &apos;/* ------------------ un-kozai the mean motion ----------------- */
        ak = pow(rec.xke / rec.no_kozai, x2o3)
        d1 = 0.75 * rec.j2 * (3.0 * rec.cosio2 - 1.0) / (rec.rteosq * rec.omeosq)
        del = d1 / (ak * ak)
        adel = ak * (1.0 - del * del - del * _
            (1.0 / 3.0 + 134.0 * del * del / 81.0))
        del = d1 / (adel * adel)
        rec.no_unkozai = rec.no_kozai / (1.0 + del)

        rec.ao = pow(rec.xke / (rec.no_unkozai), x2o3)
        rec.sinio = sin(rec.inclo)
        po = rec.ao * rec.omeosq
        rec.con42 = 1.0 - 5.0 * rec.cosio2
        rec.con41 = -rec.con42 - rec.cosio2 - rec.cosio2
        rec.ainv = 1.0 / rec.ao
        rec.posq = po * po
        rec.rp = rec.ao * (1.0 - rec.ecco)
        rec.method = &quot;n&quot;

        ts70 = epoch - 7305.0
        ds70 = flr(ts70 + 1.0e-8)
        tfrac = ts70 - ds70
        &apos;// find greenwich location at epoch
        c1 = 1.72027916940703639e-2
        thgr70 = 1.7321343856509374
        fk5r = 5.07551419432269442e-15
        c1p2p = c1 + twopi
        
        &apos;gsto1 = fmod(thgr70 + c1*ds70 + c1p2p*tfrac + ts70*ts70*fk5r, twopi)
        &apos;if (gsto1 &lt; 0.0) then
        &apos;    gsto1 = gsto1 + twopi
        &apos;end if
        
        rec.gsto = gstime(epoch + 2433281.5)

    end sub &apos;  // initl


    function sgp4init (opsmode as String, satrec as ElsetRec) as Boolean
    
        &apos;/* --------------------- local variables ------------------------ */
        
        Dim cc1sq, cc2, cc3, coef, coef1, cosio4, _
            eeta, etasq, perige, pinvsq, psisq, qzms24, _
            sfour,tc, temp, temp1, temp2, temp3, tsi, xpidot, _
            xhdot1,qzms2t, ss, x2o3, _
            delmotemp, qzms2ttemp, qzms24temp as Double
        Dim r(3) as Double
        Dim v(3) as Double
        Dim epoch as Double
                       
        epoch = (satrec.jdsatepoch + satrec.jdsatepochF) - 2433281.5
        
        &apos;/* ------------------------ initialization --------------------- */
        &apos;// sgp4fix divisor for divide by zero check on inclination
        &apos;// the old check used 1.0 + cos(pi-1.0e-9), but then compared it to
        &apos;// 1.5 e-12, so the threshold was changed to 1.5e-12 for consistency
        Const temp4 as Double = 1.5e-12

        &apos;/* ----------- set all near earth variables to zero ------------ */
        satrec.isimp = 0   
        satrec.method = &quot;n&quot; 
        satrec.aycof = 0.0
        satrec.con41 = 0.0 
        satrec.cc1 = 0.0 
        satrec.cc4 = 0.0
        satrec.cc5 = 0.0 
        satrec.d2 = 0.0
        satrec.d3 = 0.0
        satrec.d4 = 0.0 
        satrec.delmo = 0.0 
        satrec.eta = 0.0
        satrec.argpdot = 0.0 
        satrec.omgcof = 0.0 
        satrec.sinmao = 0.0
        satrec.t = 0.0 
        satrec.t2cof = 0.0 
        satrec.t3cof = 0.0
        satrec.t4cof = 0.0 
        satrec.t5cof = 0.0 
        satrec.x1mth2 = 0.0
        satrec.x7thm1 = 0.0 
        satrec.mdot = 0.0 
        satrec.nodedot = 0.0
        satrec.xlcof = 0.0 
        satrec.xmcof = 0.0 
        satrec.nodecf = 0.0

        &apos;/* ----------- set all deep space variables to zero ------------ */
        satrec.irez = 0   
        satrec.d2201 = 0.0 
        satrec.d2211 = 0.0
        satrec.d3210 = 0.0 
        satrec.d3222 = 0.0 
        satrec.d4410 = 0.0
        satrec.d4422 = 0.0 
        satrec.d5220 = 0.0 
        satrec.d5232 = 0.0
        satrec.d5421 = 0.0 
        satrec.d5433 = 0.0 
        satrec.dedt = 0.0
        satrec.del1 = 0.0 
        satrec.del2 = 0.0 
        satrec.del3 = 0.0
        satrec.didt = 0.0 
        satrec.dmdt = 0.0 
        satrec.dnodt = 0.0
        satrec.domdt = 0.0 
        satrec.e3 = 0.0 
        satrec.ee2 = 0.0
        satrec.peo = 0.0
        satrec.pgho = 0.0 
        satrec.pho = 0.0
        satrec.pinco = 0.0 
        satrec.plo = 0.0 
        satrec.se2 = 0.0
        satrec.se3 = 0.0 
        satrec.sgh2 = 0.0 
        satrec.sgh3 = 0.0
        satrec.sgh4 = 0.0 
        satrec.sh2 = 0.0 
        satrec.sh3 = 0.0
        satrec.si2 = 0.0 
        satrec.si3 = 0.0 
        satrec.sl2 = 0.0
        satrec.sl3 = 0.0 
        satrec.sl4 = 0.0 
        satrec.gsto = 0.0
        satrec.xfact = 0.0
        satrec.xgh2 = 0.0 
        satrec.xgh3 = 0.0
        satrec.xgh4 = 0.0 
        satrec.xh2 = 0.0 
        satrec.xh3 = 0.0
        satrec.xi2 = 0.0 
        satrec.xi3 = 0.0 
        satrec.xl2 = 0.0
        satrec.xl3 = 0.0 
        satrec.xl4 = 0.0 
        satrec.xlamo = 0.0
        satrec.zmol = 0.0 
        satrec.zmos = 0.0 
        satrec.atime = 0.0
        satrec.xli = 0.0 
        satrec.xni = 0.0

        &apos;/* ------------------------ earth constants ----------------------- */
        &apos;// sgp4fix identify constants and allow alternate values
        &apos;// this is now the only call for the constants
        call getgravconst(satrec.whichconst, satrec)

        
        satrec.error = 0
        satrec.operationmode = opsmode


        &apos;// single averaged mean elements
        satrec.am = satrec.em = satrec.im = satrec.Om = satrec.mm = satrec.nm = 0.0

        &apos;/* ------------------------ earth constants ----------------------- */
        ss = 78.0 / satrec.radiusearthkm + 1.0
        &apos;// sgp4fix use multiply for speed instead of pow
        qzms2ttemp = (120.0 - 78.0) / satrec.radiusearthkm
        qzms2t = qzms2ttemp * qzms2ttemp * qzms2ttemp * qzms2ttemp
        x2o3 = 2.0 / 3.0

        satrec.init = &quot;y&quot;
        satrec.t = 0.0

        &apos;// sgp4fix remove satn as it is not needed in initl
        
        call initl(epoch,satrec)
        
        satrec.a = pow(satrec.no_unkozai * satrec.tumin, (-2.0 / 3.0))
        satrec.alta = satrec.a * (1.0 + satrec.ecco) - 1.0
        satrec.altp = satrec.a * (1.0 - satrec.ecco) - 1.0
        satrec.error = 0


        if ((satrec.omeosq &gt;= 0.0) OR (satrec.no_unkozai &gt;= 0.0)) then
        
            satrec.isimp = 0
            if (satrec.rp &lt; (220.0 / satrec.radiusearthkm + 1.0)) then
                satrec.isimp = 1
            end if
            sfour = ss
            qzms24 = qzms2t
            perige = (satrec.rp - 1.0) * satrec.radiusearthkm

            &apos;/* - for perigees below 156 km, s and qoms2t are altered - */
            if (perige &lt; 156.0) then
            
                sfour = perige - 78.0
                if (perige &lt; 98.0) then
                    sfour = 20.0
                end if
                &apos;// sgp4fix use multiply for speed instead of pow
                qzms24temp = (120.0 - sfour) / satrec.radiusearthkm
                qzms24 = qzms24temp * qzms24temp * qzms24temp * qzms24temp
                sfour = sfour / satrec.radiusearthkm + 1.0
            end if
            pinvsq = 1.0 / satrec.posq

            tsi = 1.0 / (satrec.ao - sfour)
            satrec.eta = satrec.ao * satrec.ecco * tsi
            etasq = satrec.eta * satrec.eta
            eeta = satrec.ecco * satrec.eta
            psisq = abs(1.0 - etasq)
            coef = qzms24 * pow(tsi, 4.0)
            coef1 = coef / pow(psisq, 3.5)
            cc2 = coef1 * satrec.no_unkozai * (satrec.ao * (1.0 + 1.5 * etasq + eeta * _
                (4.0 + etasq)) + 0.375 * satrec.j2 * tsi / psisq * satrec.con41 * _
                (8.0 + 3.0 * etasq * (8.0 + etasq)))
            satrec.cc1 = satrec.bstar * cc2
            cc3 = 0.0
            if (satrec.ecco &gt; 1.0e-4) then
                cc3 = -2.0 * coef * tsi * satrec.j3oj2 * satrec.no_unkozai * satrec.sinio / satrec.ecco
            end if
            satrec.x1mth2 = 1.0 - satrec.cosio2
            satrec.cc4 = 2.0* satrec.no_unkozai * coef1 * satrec.ao * satrec.omeosq * _
                (satrec.eta * (2.0 + 0.5 * etasq) + satrec.ecco * _
                (0.5 + 2.0 * etasq) - satrec.j2 * tsi / (satrec.ao * psisq) * _
                (-3.0 * satrec.con41 * (1.0 - 2.0 * eeta + etasq * _
                (1.5 - 0.5 * eeta)) + 0.75 * satrec.x1mth2 * _
                (2.0 * etasq - eeta * (1.0 + etasq)) * cos(2.0 * satrec.argpo)))
            satrec.cc5 = 2.0 * coef1 * satrec.ao * satrec.omeosq * (1.0 + 2.75 * _
                (etasq + eeta) + eeta * etasq)
            cosio4 = satrec.cosio2 * satrec.cosio2
            temp1 = 1.5 * satrec.j2 * pinvsq * satrec.no_unkozai
            temp2 = 0.5 * temp1 * satrec.j2 * pinvsq
            temp3 = -0.46875 * satrec.j4 * pinvsq * pinvsq * satrec.no_unkozai
            satrec.mdot = satrec.no_unkozai + 0.5 * temp1 * satrec.rteosq * satrec.con41 + 0.0625 * _
                temp2 * satrec.rteosq * (13.0 - 78.0 * satrec.cosio2 + 137.0 * cosio4)
            satrec.argpdot = -0.5 * temp1 * satrec.con42 + 0.0625 * temp2 * _
                (7.0 - 114.0 * satrec.cosio2 + 395.0 * cosio4) + _
                temp3 * (3.0 - 36.0 * satrec.cosio2 + 49.0 * cosio4)
            xhdot1 = -temp1 * satrec.cosio
            satrec.nodedot = xhdot1 + (0.5 * temp2 * (4.0 - 19.0 * satrec.cosio2) + _
                2.0 * temp3 * (3.0 - 7.0 * satrec.cosio2)) * satrec.cosio
            xpidot = satrec.argpdot + satrec.nodedot
            satrec.omgcof = satrec.bstar * cc3 * cos(satrec.argpo)
            satrec.xmcof = 0.0
            if (satrec.ecco &gt; 1.0e-4) then
                satrec.xmcof = -x2o3 * coef * satrec.bstar / eeta
            end if
            satrec.nodecf = 3.5 * satrec.omeosq * xhdot1 * satrec.cc1
            satrec.t2cof = 1.5 * satrec.cc1
            &apos;// sgp4fix for divide by zero with xinco = 180 deg
            if (abs(satrec.cosio + 1.0) &gt; 1.5e-12) then
                satrec.xlcof = -0.25 * satrec.j3oj2 * satrec.sinio * (3.0 + 5.0 * satrec.cosio) / (1.0 + satrec.cosio)
            else
                satrec.xlcof = -0.25 * satrec.j3oj2 * satrec.sinio * (3.0 + 5.0 * satrec.cosio) / temp4
            end if
            satrec.aycof = -0.5 * satrec.j3oj2 * satrec.sinio
            &apos;// sgp4fix use multiply for speed instead of pow
            delmotemp = 1.0 + satrec.eta * cos(satrec.mo)
            satrec.delmo = delmotemp * delmotemp * delmotemp
            satrec.sinmao = sin(satrec.mo)
            satrec.x7thm1 = 7.0 * satrec.cosio2 - 1.0

            &apos;/* --------------- deep space initialization ------------- */
            if ((2 * pi / satrec.no_unkozai) &gt;= 225.0) then
            
                satrec.method = &quot;d&quot;
                satrec.isimp = 1
                tc = 0.0
                satrec.inclm = satrec.inclo

                call dscom(epoch, satrec.ecco, satrec.argpo, tc, satrec.inclo, satrec.nodeo, satrec.no_unkozai,satrec)
                
                
                satrec.ep=satrec.ecco
                satrec.inclp=satrec.inclo
                satrec.nodep=satrec.nodeo
                satrec.argpp=satrec.argpo
                satrec.mp=satrec.mo

                
                call dpper(satrec.e3, satrec.ee2, satrec.peo, satrec.pgho, _
                    satrec.pho, satrec.pinco, satrec.plo, satrec.se2, _
                    satrec.se3, satrec.sgh2, satrec.sgh3, satrec.sgh4, _
                    satrec.sh2, satrec.sh3, satrec.si2, satrec.si3, _
                    satrec.sl2, satrec.sl3, satrec.sl4, satrec.t, _
                    satrec.xgh2, satrec.xgh3, satrec.xgh4, satrec.xh2, _
                    satrec.xh3, satrec.xi2, satrec.xi3, satrec.xl2, _
                    satrec.xl3, satrec.xl4, satrec.zmol, satrec.zmos, satrec.init,satrec, _
                    satrec.operationmode)


                satrec.ecco=satrec.ep
                satrec.inclo=satrec.inclp
                satrec.nodeo=satrec.nodep
                satrec.argpo=satrec.argpp
                satrec.mo=satrec.mp


                satrec.argpm = 0.0
                satrec.nodem = 0.0
                satrec.mm = 0.0
                
                call dsinit(tc, xpidot, satrec)
            end if

            &apos;/* ----------- set variables if not deep space ----------- */
            if (satrec.isimp &lt;&gt; 1) then
            
                cc1sq = satrec.cc1 * satrec.cc1
                satrec.d2 = 4.0 * satrec.ao * tsi * cc1sq
                temp = satrec.d2 * tsi * satrec.cc1 / 3.0
                satrec.d3 = (17.0 * satrec.ao + sfour) * temp
                satrec.d4 = 0.5 * temp * satrec.ao * tsi * (221.0 * satrec.ao + 31.0 * sfour) * satrec.cc1
                satrec.t3cof = satrec.d2 + 2.0 * cc1sq
                satrec.t4cof = 0.25 * (3.0 * satrec.d3 + satrec.cc1 *  _
                    (12.0 * satrec.d2 + 10.0 * cc1sq))
                satrec.t5cof = 0.2 * (3.0 * satrec.d4 + _
                    12.0 * satrec.cc1 * satrec.d3 + _
                    6.0 * satrec.d2 * satrec.d2 + _
                    15.0 * cc1sq * (2.0 * satrec.d2 + cc1sq))
            end if
        end if &apos;// if omeosq = 0 ...

        &apos;/* finally propogate to zero epoch to initialize all others. */
        
        
        call sgp4(satrec, 0.0, r, v)

        satrec.init = &quot;n&quot;

        &apos;//sgp4fix return boolean. satrec.error contains any error codes
        sgp4init = TRUE
    end function  &apos;// sgp4init


function sgp4 (satrec as ElsetRec, tsince as Double, r, v) as Boolean
        
        Dim axnl, aynl, betal, cnod, _
            cos2u, coseo1, cosi, cosip, cosisq, cossu, cosu, _
            delm, delomg, ecose, el2, eo1, _
            esine, argpdf, pl, mrt, _
            mvt, rdotl, rl, rvdot, rvdotl, _
            sin2u, sineo1, sini, sinip, sinsu, sinu, _
            snod, su, t2, t3, t4, tem5, temp, _
            temp1, temp2, tempa, tempe, templ, u, ux, _
            uy, uz, vx, vy, vz, _
            xinc, xincp, xl, xlm, _
            xmdf, xmx, xmy, nodedf, xnode, tc, _
            x2o3, vkmpersec, delmtemp as Double
        mrt = 0.0
        Dim ktr as Integer

        &apos;/* ------------------ set mathematical constants --------------- */
        &apos;// sgp4fix divisor for divide by zero check on inclination
        &apos;// the old check used 1.0 + cos(pi-1.0e-9), but then compared it to
        &apos;// 1.5 e-12, so the threshold was changed to 1.5e-12 for consistency
        Const temp4 as Double = 1.5e-12
        x2o3 = 2.0 / 3.0
        vkmpersec = satrec.radiusearthkm * satrec.xke / 60.0

        &apos;/* --------------------- clear sgp4 error flag ----------------- */
        satrec.t = tsince
        satrec.error = 0

        &apos;/* ------- update for secular gravity and atmospheric drag ----- */
        xmdf = satrec.mo + satrec.mdot * satrec.t
        argpdf = satrec.argpo + satrec.argpdot * satrec.t
        nodedf = satrec.nodeo + satrec.nodedot * satrec.t
        satrec.argpm = argpdf
        satrec.mm = xmdf
        t2 = satrec.t * satrec.t
        satrec.nodem = nodedf + satrec.nodecf * t2
        tempa = 1.0 - satrec.cc1 * satrec.t
        tempe = satrec.bstar * satrec.cc4 * satrec.t
        templ = satrec.t2cof * t2

        delomg = 0
        delmtemp = 0
        delm = 0
        temp = 0
        t3 = 0
        t4 = 0
        mrt = 0
        
        if (satrec.isimp &lt;&gt; 1) then
        
            delomg = satrec.omgcof * satrec.t
            &apos;// sgp4fix use mutliply for speed instead of pow
            delmtemp = 1.0 + satrec.eta * cos(xmdf)
            delm = satrec.xmcof * _
                (delmtemp * delmtemp * delmtemp - _
                satrec.delmo)
            temp = delomg + delm
            satrec.mm = xmdf + temp
            satrec.argpm = argpdf - temp
            t3 = t2 * satrec.t
            t4 = t3 * satrec.t
            tempa = tempa - satrec.d2 * t2 - satrec.d3 * t3 - _
                satrec.d4 * t4
            tempe = tempe + satrec.bstar * satrec.cc5 * (sin(satrec.mm) -satrec.sinmao)
            templ = templ + satrec.t3cof * t3 + t4 * (satrec.t4cof + satrec.t * satrec.t5cof)
        end if

        
        tc = 0
        satrec.nm = satrec.no_unkozai
        satrec.em = satrec.ecco
        satrec.inclm = satrec.inclo
        if (satrec.method = &quot;d&quot;) then
        
            tc = satrec.t
            dspace(tc,satrec)        
        end if &apos;// if method = d

        if (satrec.nm &lt;= 0.0) then
        
            satrec.error = 2
            &apos;// sgp4fix add return
            sgp4 =  FALSE
            exit function
        end if
        
        satrec.am = pow((satrec.xke / satrec.nm), x2o3) * tempa * tempa
        satrec.nm = satrec.xke / pow(satrec.am, 1.5)
        satrec.em = satrec.em - tempe

        &apos;// fix tolerance for error recognition
        &apos;// sgp4fix am is fixed from the previous nm check
        if ((satrec.em &gt;= 1.0) OR (satrec.em &lt; -0.001)) then &apos;/* || (am &lt; 0.95)*/)
        
            satrec.error = 1
            &apos;// sgp4fix to return if there is an error in eccentricity
            sgp4 = FALSE
            exit function
        end if
        
        &apos;// sgp4fix fix tolerance to avoid a divide by zero
        if (satrec.em &lt; 1.0e-6) then
            satrec.em = 1.0e-6
        end if
        
        satrec.mm = satrec.mm + satrec.no_unkozai * templ
        xlm = satrec.mm + satrec.argpm + satrec.nodem
        satrec.emsq = satrec.em * satrec.em
        temp = 1.0 - satrec.emsq

        satrec.nodem = fmod(satrec.nodem, twopi)
        satrec.argpm = fmod(satrec.argpm, twopi)
        xlm = fmod(xlm, twopi)
        satrec.mm = fmod(xlm - satrec.argpm - satrec.nodem, twopi)

        &apos;// sgp4fix recover singly averaged mean elements
        satrec.am = satrec.am
        satrec.em = satrec.em
        satrec.im = satrec.inclm
        satrec.Om = satrec.nodem
        satrec.om = satrec.argpm
        satrec.mm = satrec.mm
        satrec.nm = satrec.nm

        &apos;/* ----------------- compute extra mean quantities ------------- */
        satrec.sinim = sin(satrec.inclm)
        satrec.cosim = cos(satrec.inclm)

        &apos;/* -------------------- add lunar-solar periodics -------------- */
        satrec.ep = satrec.em
        xincp = satrec.inclm
        satrec.inclp = satrec.inclm
        satrec.argpp = satrec.argpm
        satrec.nodep = satrec.nodem
        satrec.mp = satrec.mm
        sinip = satrec.sinim
        cosip = satrec.cosim
        if (satrec.method = &quot;d&quot;) then
        
            call dpper(satrec.e3, satrec.ee2, satrec.peo, satrec.pgho, _
                    satrec.pho, satrec.pinco, satrec.plo, satrec.se2, _
                    satrec.se3, satrec.sgh2, satrec.sgh3, satrec.sgh4, _
                    satrec.sh2, satrec.sh3, satrec.si2, satrec.si3, _
                    satrec.sl2, satrec.sl3, satrec.sl4, satrec.t, _
                    satrec.xgh2, satrec.xgh3, satrec.xgh4, satrec.xh2, _
                    satrec.xh3, satrec.xi2, satrec.xi3, satrec.xl2, _
                    satrec.xl3, satrec.xl4, satrec.zmol, satrec.zmos, _ 
                    &quot;n&quot;, satrec, satrec.operationmode)
            
            xincp = satrec.inclp
            if (xincp &lt; 0.0) then
            
                xincp = -xincp
                satrec.nodep = satrec.nodep + pi
                satrec.argpp = satrec.argpp - pi
            end if
            
            if ((satrec.ep &lt; 0.0) OR (satrec.ep &gt; 1.0)) then
            
                satrec.error = 3
                &apos;// sgp4fix add return
                sgp4 = FALSE
                exit function
            end if
        end if &apos; // if method = d

        &apos;/* -------------------- long period periodics ------------------ */
        if (satrec.method = &quot;d&quot;) then
        
            sinip = sin(xincp)
            cosip = cos(xincp)
            satrec.aycof = -0.5*satrec.j3oj2*sinip
            &apos;// sgp4fix for divide by zero for xincp = 180 deg
            if (abs(cosip + 1.0) &gt; 1.5e-12) then
                satrec.xlcof = -0.25 * satrec.j3oj2 * sinip * (3.0 + 5.0 * cosip) / (1.0 + cosip)
            else
                satrec.xlcof = -0.25 * satrec.j3oj2 * sinip * (3.0 + 5.0 * cosip) / temp4
            end if
        end if
        axnl = satrec.ep * cos(satrec.argpp)
        temp = 1.0 / (satrec.am * (1.0 - satrec.ep * satrec.ep))
        aynl = satrec.ep* sin(satrec.argpp) + temp * satrec.aycof
        xl = satrec.mp + satrec.argpp + satrec.nodep + temp * satrec.xlcof * axnl

        &apos;/* --------------------- solve kepler&apos;s equation --------------- */
        u = fmod(xl - satrec.nodep, twopi)
        eo1 = u
        tem5 = 9999.9
        ktr = 1
        sineo1 = 0
        coseo1 = 0
        &apos;//   sgp4fix for kepler iteration
        &apos;//   the following iteration needs better limits on corrections
        Do while (abs(tem5) &gt;= 1.0e-12) AND (ktr &lt;= 10)
        
            sineo1 = sin(eo1)
            coseo1 = cos(eo1)
            tem5 = 1.0 - coseo1 * axnl - sineo1 * aynl
            tem5 = (u - aynl * coseo1 + axnl * sineo1 - eo1) / tem5
            if (abs(tem5) &gt;= 0.95) then
                if(tem5 &gt; 0.0) then
                  tem5 = 0.95
                else
                  tem5 = -0.95
                end if
            end if
            eo1 = eo1 + tem5
            ktr = ktr + 1
        Loop

        &apos;/* ------------- short period preliminary quantities ----------- */
        ecose = axnl*coseo1 + aynl*sineo1
        esine = axnl*sineo1 - aynl*coseo1
        el2 = axnl*axnl + aynl*aynl
        pl = satrec.am*(1.0 - el2)
        if (pl &lt; 0.0) then
            satrec.error = 4
            &apos;// sgp4fix add return
            sgp4 = FALSE
            exit function
        else
        
            rl = satrec.am * (1.0 - ecose)
            rdotl = Sqr(satrec.am) * esine / rl
            rvdotl = Sqr(pl) / rl
            betal = Sqr(1.0 - el2)
            temp = esine / (1.0 + betal)
            sinu = satrec.am / rl * (sineo1 - aynl - axnl * temp)
            cosu = satrec.am / rl * (coseo1 - axnl + aynl * temp)
            su = Atn2(sinu, cosu)
            sin2u = (cosu + cosu) * sinu
            cos2u = 1.0 - 2.0 * sinu * sinu
            temp = 1.0 / pl
            temp1 = 0.5 * satrec.j2 * temp
            temp2 = temp1 * temp

            &apos;/* -------------- update for short period periodics ------------ */
            if (satrec.method = &quot;d&quot;) then
            
                cosisq = cosip * cosip
                satrec.con41 = 3.0*cosisq - 1.0
                satrec.x1mth2 = 1.0 - cosisq
                satrec.x7thm1 = 7.0*cosisq - 1.0
            end if
            mrt = rl * (1.0 - 1.5 * temp2 * betal * satrec.con41) + _
                0.5 * temp1 * satrec.x1mth2 * cos2u
            su = su - 0.25 * temp2 * satrec.x7thm1 * sin2u
            xnode = satrec.nodep + 1.5 * temp2 * cosip * sin2u
            xinc = xincp + 1.5 * temp2 * cosip * sinip * cos2u
            mvt = rdotl - satrec.nm * temp1 * satrec.x1mth2 * sin2u / satrec.xke
            rvdot = rvdotl + satrec.nm * temp1 * (satrec.x1mth2 * cos2u + _
                1.5 * satrec.con41) / satrec.xke

            &apos;/* --------------------- orientation vectors ------------------- */
            sinsu = sin(su)
            cossu = cos(su)
            snod = sin(xnode)
            cnod = cos(xnode)
            sini = sin(xinc)
            cosi = cos(xinc)
            xmx = -snod * cosi
            xmy = cnod * cosi
            ux = xmx * sinsu + cnod * cossu
            uy = xmy * sinsu + snod * cossu
            uz = sini * sinsu
            vx = xmx * cossu - cnod * sinsu
            vy = xmy * cossu - snod * sinsu
            vz = sini * cossu

            &apos;/* --------- position and velocity (in km and km/sec) ---------- */
            r(0) = (mrt * ux)* satrec.radiusearthkm
            r(1) = (mrt * uy)* satrec.radiusearthkm
            r(2) = (mrt * uz)* satrec.radiusearthkm
            v(0) = (mvt * ux + rvdot * vx) * vkmpersec
            v(1) = (mvt * uy + rvdot * vy) * vkmpersec
            v(2) = (mvt * uz + rvdot * vz) * vkmpersec
        end if  &apos;// if pl &gt; 0

        &apos;// sgp4fix for decaying satellites
        if (mrt &lt; 1.0) then
            satrec.error = 6
            sgp4 = FALSE
        else
            sgp4 = TRUE
        end if
end function






function flr( v as Double ) as long
&apos;    flr = v-(v Mod 1)
    flr = Int(v)
end function

sub jday(year as Integer, mon as Integer, dy as Integer, hr as Integer, mn as Integer, sec as Double, jdout)
  Dim jd as Double
  Dim jdFrac as Double
  Dim t1 as Double
  Dim t2 as Double
  Dim t3 as Double
  Dim dtt as Double
  
  t1 = (mon+9)/12.0
  t1 =  flr(t1)
  t2 =  flr((7 * (year + t1)) * 0.25)
  t3 = flr(275 * mon / 9.0)
  jd = 367.0 * year
  jd = jd - t2
  jd = jd + t3 +dy
  jd = jd + 1721013.5
          
  jdFrac = (sec + mn*60.0 +hr*3600.0)/86400.0
  if Abs(jdFrac) &gt; 1 then
    dtt = flr(jdFrac)
    jd = jd+dtt
    jdFrac = jdFrac - dtt
  end if
  
  jdout(0)=jd
  jdout(1)=jdFrac
end sub

sub getgravconst(whichconst, rec)
    rec.whichconst = whichconst

    if whichconst = wgs72old then
      rec.mu = 398600.79964       &apos; #// in km3 / s2
      rec.radiusearthkm = 6378.135  &apos;  # // km
      rec.xke = 0.0743669161     &apos;  # // reciprocal of tumin
      rec.tumin = 1.0 / rec.xke
      rec.j2 = 0.001082616
      rec.j3 = -0.00000253881
      rec.j4 = -0.00000165597
      rec.j3oj2 = rec.j3 / rec.j2
     &apos; #// ------------ wgs-72 constants ------------
    Elseif whichconst = wgs72 then
      rec.mu = 398600.8         &apos;   #// in km3 / s2
      rec.radiusearthkm = 6378.135 &apos;   # // km
      rec.xke = 60.0 / Sqr(rec.radiusearthkm*rec.radiusearthkm*rec.radiusearthkm / rec.mu)
      rec.tumin = 1.0 / rec.xke
      rec.j2 = 0.001082616
      rec.j3 = -0.00000253881
      rec.j4 = -0.00000165597
      rec.j3oj2 = rec.j3 / rec.j2
    else &apos; # wgs84
     &apos; #// ------------ wgs-84 constants ------------
      rec.mu = 398600.5         
      rec.radiusearthkm = 6378.137
      rec.xke = 60.0 / Sqr(rec.radiusearthkm*rec.radiusearthkm*rec.radiusearthkm / rec.mu)
      rec.tumin = 1.0 / rec.xke
      rec.j2 = 0.00108262998905
      rec.j3 = -0.00000253215306
      rec.j4 = -0.00000161098761
      rec.j3oj2 = rec.j3 / rec.j2
    end if
    
end sub


function pow(n1 as Double, n2 as Double) as Double
  pow = n1^n2
end function

function fmod(numer as Double, denom as Double) as Double
  Dim tquot as Double
  
  if(denom &lt;&gt; 0) then
    tquot = flr(numer/denom)
    fmod = numer-tquot*denom
  else
    fmod = 0
  end if
end function

function gstime(jdut1 as Double) as Double
  Dim tut1 as Double
  Dim temp as Double
    tut1 = (jdut1 - 2451545.0) / 36525.0
    temp = (-6.2e-6* tut1 * tut1 * tut1 + 0.093104 * tut1 * tut1 + (876600.0 * 3600 + 8640184.812866) * tut1 + 67310.54841)
    temp = fmod(temp * deg2rad / 240.0, twopi)

    if (temp &lt; 0.0) then
      temp = temp + twopi
    end if

    gstime = temp
end function

Function Atn2(y As Double, x As Double) As Double

  If x &gt; 0 Then

    Atn2 = Atn(y / x)

  ElseIf x &lt; 0 Then

    Atn2 = Sgn(y) * (Pi - Atn(Abs(y / x)))

  ElseIf y = 0 Then

    Atn2 = 0

  Else

    Atn2 = Sgn(y) * Pi / 2

  End If

End Function
