'==========================================================================
'  Peng-Robinson package (+Peneloux) and the isenthalpic temperature march
'  Applies when every component of the stream is in Data_PR (98 pure components)
'==========================================================================
Private prKey() As String, prMW() As Double, prTc() As Double, prPc() As Double, prOm() As Double, prVc() As Double
Private prHo() As Double, prCo(1 To 6) As Variant, prPen() As Double, nPRd As Long
Private prAl1() As String, prAl2() As String, nPRal As Long
Private prKa() As String, prKb() As String, prKv() As Double, nPRk As Long
Private prOK As Boolean, prN As Long, prI() As Long, prKm() As Double, tuneOn As Boolean, isoT As Boolean, useIso As Boolean
Private cSig(1 To 4) As String, cOkv(1 To 4) As Boolean, cIdx(1 To 4, 1 To 60) As Long, cKm(1 To 4, 1 To 60, 1 To 60) As Double, cNext As Long
Private sTmain As Double
Private Const PRG As Double = 83.14462618

Private Sub ReadPR()
    Dim ws As Worksheet, v As Variant, i As Long, j As Long, lr As Long
    nPRd = 0: nPRal = 0: nPRk = 0: cNext = 1
    For i = 1 To 4: cSig(i) = "": Next i
    On Error Resume Next
    Set ws = Worksheets("Data_PR")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 2 Then Exit Sub
    v = ws.Range("A2:O" & lr).Value: nPRd = lr - 1
    ReDim prKey(1 To nPRd): ReDim prMW(1 To nPRd): ReDim prTc(1 To nPRd): ReDim prPc(1 To nPRd): ReDim prOm(1 To nPRd)
    ReDim prVc(1 To nPRd): ReDim prHo(1 To nPRd): ReDim prPen(1 To nPRd)
    For j = 1 To 6: ReDim prCo(j)(1 To nPRd): Next j
    For i = 1 To nPRd
        prKey(i) = UCase$(CStr(v(i, 1))): prMW(i) = v(i, 2): prTc(i) = v(i, 3): prPc(i) = v(i, 4): prOm(i) = v(i, 5)
        prVc(i) = v(i, 6): prHo(i) = v(i, 7)
        For j = 1 To 6: prCo(j)(i) = CDbl(v(i, 7 + j)): Next j
        prPen(i) = v(i, 15)
    Next i
    Set ws = Worksheets("Data_PRAlias")
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    v = ws.Range("A2:B" & lr).Value: nPRal = lr - 1
    ReDim prAl1(1 To nPRal): ReDim prAl2(1 To nPRal)
    For i = 1 To nPRal: prAl1(i) = CStr(v(i, 1)): prAl2(i) = UCase$(CStr(v(i, 2))): Next i
    Set ws = Worksheets("Data_PRKij")
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    v = ws.Range("A2:C" & lr).Value: nPRk = lr - 1
    ReDim prKa(1 To nPRk): ReDim prKb(1 To nPRk): ReDim prKv(1 To nPRk)
    For i = 1 To nPRk: prKa(i) = UCase$(CStr(v(i, 1))): prKb(i) = UCase$(CStr(v(i, 2))): prKv(i) = v(i, 3): Next i
End Sub

Private Function PRIdx(ByVal nm As String) As Long
    Dim u As String, n As String, i As Long, c As String, k As Long
    u = UCase$(Trim$(nm))
    For i = 1 To nPRd
        If prKey(i) = u Then PRIdx = i: Exit Function
    Next i
    For k = 1 To Len(u)
        c = Mid$(u, k, 1)
        If (c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Or c = "=" Then n = n & c
    Next k
    For i = 1 To nPRal
        If prAl1(i) = n Then
            For k = 1 To nPRd
                If prKey(k) = prAl2(i) Then PRIdx = k: Exit Function
            Next k
        End If
    Next i
    For i = 1 To nPRd
        If prKey(i) = n Then PRIdx = i: Exit Function
    Next i
End Function
Private Function KijOf(ByVal a As String, ByVal b As String) As Double
    Dim i As Long, t As String
    If a > b Then t = a: a = b: b = t
    For i = 1 To nPRk
        If prKa(i) = a And prKb(i) = b Then KijOf = prKv(i): Exit Function
    Next i
End Function

' map the current feed components onto the PR databank (cached per component list)
Private Sub PRSetup()
    Dim sig As String, i As Long, j As Long, c As Long, ix As Long
    prOK = False
    If nPRd = 0 Or Not flashOn Or fN < 1 Or fN > 60 Then Exit Sub
    sig = fN & "|"
    For i = 1 To fN: sig = sig & fNm(i) & ",": Next i
    For c = 1 To 4
        If cSig(c) = sig Then
            prOK = cOkv(c)
            If prOK Then
                prN = fN: ReDim prI(1 To prN): ReDim prKm(1 To prN, 1 To prN)
                For i = 1 To prN
                    prI(i) = cIdx(c, i)
                    For j = 1 To prN: prKm(i, j) = cKm(c, i, j): Next j
                Next i
            End If
            Exit Sub
        End If
    Next c
    c = cNext: cNext = cNext + 1: If cNext > 4 Then cNext = 1
    cSig(c) = sig: cOkv(c) = False
    ReDim prI(1 To fN)
    For i = 1 To fN
        ix = PRIdx(fNm(i))
        If ix = 0 Then Exit Sub
        prI(i) = ix
    Next i
    prN = fN: ReDim prKm(1 To prN, 1 To prN)
    For i = 1 To prN
        cIdx(c, i) = prI(i)
        For j = 1 To prN
            If i <> j Then prKm(i, j) = KijOf(prKey(prI(i)), prKey(prI(j)))
            cKm(c, i, j) = prKm(i, j)
        Next j
    Next i
    cOkv(c) = True: prOK = True
End Sub

Private Sub SetTune(y() As Double)
    Dim i As Long, yh As Double
    For i = 1 To prN
        If prKey(prI(i)) = "H2" Then yh = yh + y(i)
    Next i
    tuneOn = (yh >= 0.5)
End Sub
Private Function TnC(ByVal i As Long) As Double
    If tuneOn Then
        If prKey(prI(i)) = "H2" Then TnC = 0.213118 Else TnC = 0.592434
    Else
        TnC = 1#
    End If
End Function
Private Function PRAlpha(ByVal i As Long, ByVal T As Double) As Double
    Dim d As Long, w As Double, m As Double
    d = prI(i): w = prOm(d)
    m = 0.37464 + 1.54226 * w - 0.26992 * w * w
    If tuneOn And prKey(d) = "H2" Then m = 0.258927
    PRAlpha = (1# + m * (1# - Sqr(T / prTc(d)))) ^ 2
End Function
Private Function PRa(ByVal T As Double, y() As Double) As Double
    Dim ai() As Double, i As Long, j As Long, s As Double, d As Long
    ReDim ai(1 To prN)
    For i = 1 To prN
        d = prI(i): ai(i) = 0.45724 * PRG ^ 2 * prTc(d) ^ 2 / prPc(d) * PRAlpha(i, T)
    Next i
    For i = 1 To prN
        For j = 1 To prN
            s = s + y(i) * y(j) * Sqr(ai(i) * ai(j)) * (1# - prKm(i, j))
        Next j
    Next i
    PRa = s
End Function
Private Function PRb(y() As Double) As Double
    Dim i As Long, s As Double, d As Long
    For i = 1 To prN
        d = prI(i): s = s + y(i) * 0.0778 * PRG * prTc(d) / prPc(d)
    Next i
    PRb = s
End Function
Private Function PRc(y() As Double) As Double
    Dim i As Long, s As Double
    For i = 1 To prN
        s = s + y(i) * prPen(prI(i)) * TnC(i)
    Next i
    PRc = s
End Function
Private Function PRM(y() As Double) As Double
    Dim i As Long, s As Double
    For i = 1 To prN: s = s + y(i) * prMW(prI(i)): Next i
    PRM = s
End Function
Private Function PRZ(ByVal T As Double, ByVal P As Double, y() As Double, ByVal liq As Boolean) As Double
    Dim a As Double, pA As Double, pB As Double, c1 As Double, c2 As Double, c3 As Double, z As Double, zn As Double, f As Double, df As Double, it As Long
    a = PRa(T, y): pA = a * P / (PRG * T) ^ 2: pB = PRb(y) * P / (PRG * T)
    c1 = -(1# - pB): c2 = pA - 3# * pB * pB - 2# * pB: c3 = -(pA * pB - pB * pB - pB ^ 3)
    If liq Then z = pB * 1.0001 Else z = 1#
    For it = 1 To 400
        f = ((z + c1) * z + c2) * z + c3: df = (3# * z + 2# * c1) * z + c2
        If df = 0 Then Exit For
        zn = z - f / df
        If liq Then If zn < pB * 1.0001 Then zn = pB * 1.0001
        If Abs(zn - z) < 0.0000000000001 Then z = zn: Exit For
        z = zn
    Next it
    PRZ = z
End Function
Private Function PRHres(ByVal T As Double, ByVal P As Double, y() As Double, ByVal liq As Boolean) As Double
    Dim z As Double, ap As Double, a As Double, B As Double, r2 As Double, bm As Double
    z = PRZ(T, P, y, liq)
    ap = (PRa(T + 0.01, y) - PRa(T - 0.01, y)) / 0.02: a = PRa(T, y): bm = PRb(y)
    B = bm * P / (PRG * T): r2 = Sqr(2#)
    PRHres = 0.1 * (PRG * T * (z - 1#) + (T * ap - a) / (2# * r2 * bm) * Log((z + (1# + r2) * B) / (z + (1# - r2) * B)))
End Function
Private Function PolyV(ByVal d As Long, ByVal T As Double) As Double
    Dim j As Long, s As Double
    For j = 1 To 6: s = s + prCo(j)(d) * T ^ (j - 1): Next j
    PolyV = s
End Function
Private Function DPolyV(ByVal d As Long, ByVal T As Double) As Double
    Dim j As Long, s As Double
    For j = 2 To 6: s = s + (j - 1) * prCo(j)(d) * T ^ (j - 2): Next j
    DPolyV = s
End Function
' ideal-gas Cp (J/mol/K) and H (J/mol) of the mixture
Private Sub PRIdeal(ByVal T As Double, y() As Double, ByRef cp As Double, ByRef h As Double)
    Dim i As Long, d As Long, cpi As Double, hi As Double, T0 As Double, e0 As Double, e1 As Double
    cp = 0: h = 0
    If tuneOn Then e0 = 550.947: e1 = -0.463104
    T0 = (77# + 459.67) / 1.8
    For i = 1 To prN
        d = prI(i)
        If prKey(d) = "H2" Then
            cpi = (3.004605 + 0.0016031 * T - 0.0000013188 * T * T) / 0.238846 * prMW(d)
            hi = 1.8 * (3.004605 * (T - T0) + 0.0016031 * (T * T - T0 * T0) / 2# - 0.0000013188 * (T ^ 3 - T0 ^ 3) / 3#) / 0.429923 * prMW(d)
        Else
            cpi = prMW(d) * DPolyV(d, T) + e1
            hi = prMW(d) * PolyV(d, T) + prHo(d) + e0 + e1 * (T - 298.15)
        End If
        cp = cp + y(i) * cpi: h = h + y(i) * hi
    Next i
End Sub
Private Function PRHonly(ByVal T As Double, ByVal P As Double, y() As Double, ByVal liq As Boolean) As Double
    Dim cp As Double, h As Double
    SetTune y
    PRIdeal T, y, cp, h
    PRHonly = (h + PRHres(T, P, y, liq)) / PRM(y) * 0.429923
End Function

' vapour Z (Peneloux), viscosity (cP), Cp/Cv of the vapour y at T(F), P(psia)
Private Sub PRVapour(y() As Double, ByVal TF As Double, ByVal Ppsia As Double, ByRef zv As Double, ByRef mu As Double, ByRef gam As Double)
    Dim yv() As Double, sy As Double, i As Long, j As Long, T As Double, P As Double, d As Long
    Dim zk As Double, v As Double, a As Double, ap As Double, cpi As Double, hi As Double, dT As Double, cp As Double
    Dim q As Double, dPdT As Double, dPdv As Double, cpmcv As Double, vt As Double, bm As Double, mm As Double, g As Double
    Dim mui() As Double, mu0 As Double, den As Double, phi As Double, tcm As Double, pcm As Double, vcm As Double, rr As Double
    Dim xi As Double, rhs As Double, e As Double, sMu As Double, mHc As Double, mQ As Double
    ReDim yv(1 To prN)
    For i = 1 To prN: sy = sy + y(i): Next i
    For i = 1 To prN: yv(i) = y(i) / sy: Next i
    SetTune yv
    T = (TF + 459.67) / 1.8: P = Ppsia * 0.0689476
    zk = PRZ(T, P, yv, False)
    bm = PRb(yv): mm = PRM(yv)
    v = zk * PRG * T / P: a = PRa(T, yv)
    ap = (PRa(T + 0.01, yv) - PRa(T - 0.01, yv)) / 0.02
    PRIdeal T, yv, cpi, hi
    dT = 0.05
    cp = cpi + (PRHres(T + dT, P, yv, False) - PRHres(T - dT, P, yv, False)) / (2# * dT)
    q = v * v + 2# * bm * v - bm * bm
    dPdT = PRG / (v - bm) - ap / q
    dPdv = -PRG * T / (v - bm) ^ 2 + a * (2# * v + 2# * bm) / q ^ 2
    cpmcv = -T * dPdT ^ 2 / dPdv * 0.1
    g = IIf(tuneOn, 1.0011, 1#)
    vt = v - PRc(yv)
    zv = P * vt / (PRG * T)
    gam = cp / (cp - g * cpmcv)
    ' viscosity: Lucas + Wilke + Jossi-Stiel-Thodos
    ReDim mui(1 To prN)
    mHc = IIf(tuneOn, 0.968744, 1#): mQ = IIf(tuneOn, -0.0108797, 0#): sMu = IIf(tuneOn, 1.34035, 1#)
    For i = 1 To prN
        d = prI(i)
        If prKey(d) = "H2" Then
            mui(i) = 0.009346 * (T / 310.928) ^ 1.5 * (310.928 + 90.91) / (T + 90.91) * 10000#
        Else
            xi = 0.176 * (prTc(d) / (prMW(d) ^ 3 * prPc(d) ^ 4)) ^ (1# / 6#)
            e = (0.807 * (T / prTc(d)) ^ 0.618 - 0.357 * Exp(-0.449 * (T / prTc(d))) + 0.34 * Exp(-4.058 * (T / prTc(d))) + 0.018) / xi
            mui(i) = e * mHc * (T / 373.15) ^ mQ
        End If
    Next i
    For i = 1 To prN
        den = 0
        For j = 1 To prN
            phi = (1# + Sqr(mui(i) / mui(j)) * (prMW(prI(j)) / prMW(prI(i))) ^ 0.25) ^ 2 / Sqr(8# * (1# + prMW(prI(i)) / prMW(prI(j))))
            den = den + yv(j) * phi
        Next j
        mu0 = mu0 + yv(i) * mui(i) / den
    Next i
    For i = 1 To prN
        d = prI(i)
        tcm = tcm + yv(i) * prTc(d): pcm = pcm + yv(i) * prPc(d): vcm = vcm + yv(i) * prVc(d)
    Next i
    rr = P / (zk * PRG * T) * vcm
    mu = mu0
    If rr > 0.1 Then
        xi = tcm ^ (1# / 6#) / (mm ^ 0.5 * (pcm / 1.01325) ^ (2# / 3#))
        rhs = 1.023 + 0.23364 * rr + 0.58533 * rr ^ 2 - 0.40758 * rr ^ 3 + 0.093324 * rr ^ 4
        mu = mu0 + sMu * (rhs ^ 4 - 1#) / xi
    End If
    mu = mu * 0.0001
End Sub

' flash at (p, T) without touching the station state
Private Sub FlashRaw(ByVal pPsia As Double, ByVal tF As Double, ByRef beta As Double, k() As Double, x() As Double, y() As Double, _
                     ByRef vmw As Double, ByRef lmw As Double)
    Dim i As Long, n As Long, ratio As Double, t_r As Double, tref_r As Double, inv As Double, w As Double, sx As Double, sy As Double
    n = fN: ReDim k(1 To n): ReDim x(1 To n): ReDim y(1 To n)
    If pPsia < 0.000001 Then pPsia = 0.000001
    ratio = fPref / pPsia
    t_r = tF + 459.67: tref_r = fTref + 459.67
    If fHasC And Abs(t_r - tref_r) > 0.000000001 And t_r > 0 And tref_r > 0 Then inv = 1# / tref_r - 1# / t_r
    For i = 1 To n
        w = 1#
        If inv <> 0 And fTcR(i) > 0 Then
            w = 5.373 * (1# + fOm(i)) * fTcR(i) * inv
            If w > 50 Then w = 50
            If w < -50 Then w = -50
            w = Exp(w)
        End If
        k(i) = ClampK(fKref(i) * ratio * w, 1E+12)
    Next i
    beta = RR(n, fZ, k)
    For i = 1 To n
        x(i) = fZ(i) / (1# + beta * (k(i) - 1#)): y(i) = k(i) * x(i)
        sx = sx + x(i): sy = sy + y(i)
    Next i
    If sx = 0 Then sx = 1
    If sy = 0 Then sy = 1
    vmw = 0: lmw = 0
    For i = 1 To n
        x(i) = x(i) / sx: y(i) = y(i) / sy
        vmw = vmw + y(i) * fMWc(i): lmw = lmw + x(i) * fMWc(i)
    Next i
End Sub

' total specific enthalpy (BTU/lb) of the flashed stream at (T, P); ok False if it cannot be evaluated
Private Function StreamH(ByVal tF As Double, ByVal pPsia As Double, ByRef ok As Boolean) As Double
    Dim beta As Double, k() As Double, x() As Double, y() As Double, vmw As Double, lmw As Double
    Dim vm As Double, lm As Double, h As Double, T As Double, P As Double, i As Long, sy As Double
    ok = False
    FlashRaw pPsia, tF, beta, k, x, y, vmw, lmw
    vm = beta * fTotMoles * vmw: lm = (1# - beta) * fTotMoles * lmw
    If vm + lm <= 0 Then Exit Function
    T = (tF + 459.67) / 1.8: P = pPsia * 0.0689476
    If vm > 0 Then h = h + vm * PRHonly(T, P, y, False)
    If lm > 0 Then h = h + lm * PRHonly(T, P, x, True)
    StreamH = h / (vm + lm): ok = True
End Function

' outlet temperature of an adiabatic, no-work pressure change
Private Function IsoTOut(ByVal tIn As Double, ByVal pIn As Double, ByVal pOut As Double) As Double
    Dim hIn As Double, ok As Boolean, ok0 As Boolean, ok1 As Boolean, ok2 As Boolean
    Dim t0 As Double, t1 As Double, t2 As Double, f0 As Double, f1 As Double, f2 As Double, it As Long
    IsoTOut = tIn
    If Not prOK Or Abs(pIn - pOut) < 0.000000001 Then Exit Function
    hIn = StreamH(tIn, pIn, ok): If Not ok Then Exit Function
    t0 = tIn: t1 = tIn - 1#
    f0 = StreamH(t0, pOut, ok0) - hIn: f1 = StreamH(t1, pOut, ok1) - hIn
    If Not (ok0 And ok1) Then Exit Function
    For it = 1 To 40
        If f1 = f0 Then Exit For
        t2 = t1 - f1 * (t1 - t0) / (f1 - f0)
        If t2 < tIn - 250# Then t2 = tIn - 250#
        If t2 > tIn + 250# Then t2 = tIn + 250#
        f2 = StreamH(t2, pOut, ok2) - hIn
        If Not ok2 Then Exit Function
        t0 = t1: f0 = f1: t1 = t2: f1 = f2
        If Abs(t1 - t0) < 0.00001 Then Exit For
    Next it
    IsoTOut = t1
End Function

