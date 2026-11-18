Attribute VB_Name = "Procalc"
Option Explicit
'==========================================================================
'  PROCALC HYDRAULICS  -  macro-enabled workbook engine (v1)
'  Sheets used : Control, Input, Streams, Data_K, Data_OD, Data_Pipe, Data_PMS
'  Sheets made : Pressure_Profile, CV_Summary   (rebuilt on every run)
'  Units       : FPS  (psia, degF, lb/hr, ft, in, cP, lb/ft3)  - same as hydraulics.py
'  Model       : Darcy-Weisbach (Swamee-Jain), K-method fittings, homogeneous
'                two-phase (McAdams mu), gas rho = P*MW/(Z R T) at local P (isothermal)
'==========================================================================
Private Const G_SI As Double = 9.80665
Private Const IN_TO_M As Double = 0.0254
Private Const FT_TO_M As Double = 0.3048
Private Const LBHR_TO_KGS As Double = 0.45359237 / 3600#
Private Const LBFT3_TO_KGM3 As Double = 16.018463
Private Const PA_TO_PSI As Double = 1# / 6894.757293
Private Const PSIA_TO_PA As Double = 6894.757293
Private Const ROUGH_M As Double = 0.0000457
Private Const R_UNIV As Double = 8.314462618
Private Const PI As Double = 3.14159265358979

'---- stream properties of the circuit being solved --------------------------
Private sT As Double, sPst As Double, sVap As Double, sLiq As Double
Private sMW As Double, sVrho As Double, sLrho As Double
Private sVmu As Double, sLmu As Double, sZ As Double, sGam As Double
Private sPhase As String, sStream As String

'---- input table ---------------------------------------------------------
Private inp As Variant, nRow As Long
Private hnames() As String, nH As Long
Private kNames As Variant, kVals As Variant, nK As Long
Private odNps As Variant, odNum As Variant, odVal As Variant, nOD As Long
Private pwNps As Variant, pwSch As Variant, pwWall As Variant, nPW As Long
Private pmCls As Variant, pmLo As Variant, pmHi As Variant, pmSch As Variant, pmMin As Variant, nPM As Long

'---- output buffer -----------------------------------------------------------
Private out() As Variant, nOut As Long
Private cvOut() As Variant, nCV As Long
Private gMult As Double, hdrRow As Long

'==========================================================================
'  Entry point - assigned to the RUN button on the Control sheet
'==========================================================================
Public Sub RunHydraulics()
    Dim t0 As Double: t0 = Timer
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    LoadTables
    gMult = Val(Worksheets("Control").Range("C4").Value) / 100#
    If gMult <= 0 Then gMult = 1#
    ReDim out(1 To nRow + 10, 1 To 24): nOut = 0
    ReDim cvOut(1 To nRow + 10, 1 To 14): nCV = 0

    Dim r As Long, r1 As Long, cid As String, nCirc As Long
    r = 1
    Do While r <= nRow
        cid = CStr(GetV(r, "Circuit"))
        r1 = r
        Do While r <= nRow
            If CStr(GetV(r, "Circuit")) <> cid And Len(CStr(GetV(r, "Circuit"))) > 0 Then Exit Do
            r = r + 1
        Loop
        SolveCircuit r1, r - 1, cid
        nCirc = nCirc + 1
    Loop
    WriteOutputs
    Worksheets("Control").Range("C8").Value = nCirc & " circuit(s), " & nOut & " rows  -  " & _
        Format(Timer - t0, "0.00") & " s  -  " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Worksheets("Control").Range("C8").Font.Color = RGB(0, 128, 0)
Done:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Worksheets("Control").Range("C8").Value = "ERROR: " & Err.Description
    Worksheets("Control").Range("C8").Font.Color = RGB(200, 0, 0)
    Resume Done
End Sub

'==========================================================================
'  Table loading
'==========================================================================
Private Sub LoadTables()
    Dim ws As Worksheet, v As Variant, i As Long, lr As Long, lc As Long
    Set ws = Worksheets("Input")
    ' header row = first of rows 1-3 whose column A reads "Circuit"
    hdrRow = 0
    For i = 1 To 3
        If UCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = "CIRCUIT" Then hdrRow = i: Exit For
    Next i
    If hdrRow = 0 Then Err.Raise 513, , "Input sheet: header row (A = Circuit) not found in rows 1-3"
    lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row     ' Line No column drives the last row
    lc = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    If lr <= hdrRow Then Err.Raise 513, , "Input sheet has no data rows"
    inp = ws.Range(ws.Cells(hdrRow, 1), ws.Cells(lr, lc)).Value
    nRow = lr - hdrRow
    nH = lc: ReDim hnames(1 To lc)
    For i = 1 To lc: hnames(i) = NormH(CStr(inp(1, i))): Next i
    ' drop empty rows (no Line No) by compacting later via GetV; keep simple
    Set ws = Worksheets("Data_K")
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    v = ws.Range("A2:B" & lr).Value: nK = lr - 1
    ReDim kNames(1 To nK): ReDim kVals(1 To nK)
    For i = 1 To nK: kNames(i) = LCase$(Trim$(CStr(v(i, 1)))): kVals(i) = CDbl(v(i, 2)): Next i
    Set ws = Worksheets("Data_OD")
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    v = ws.Range("A2:C" & lr).Value: nOD = lr - 1
    ReDim odNps(1 To nOD): ReDim odNum(1 To nOD): ReDim odVal(1 To nOD)
    For i = 1 To nOD: odNps(i) = CStr(v(i, 1)): odNum(i) = CDbl(v(i, 2)): odVal(i) = CDbl(v(i, 3)): Next i
    Set ws = Worksheets("Data_Pipe")
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    v = ws.Range("A2:C" & lr).Value: nPW = lr - 1
    ReDim pwNps(1 To nPW): ReDim pwSch(1 To nPW): ReDim pwWall(1 To nPW)
    For i = 1 To nPW: pwNps(i) = CStr(v(i, 1)): pwSch(i) = UCase$(CStr(v(i, 2))): pwWall(i) = CDbl(v(i, 3)): Next i
    Set ws = Worksheets("Data_PMS")
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    v = ws.Range("A2:E" & lr).Value: nPM = lr - 1
    ReDim pmCls(1 To nPM): ReDim pmLo(1 To nPM): ReDim pmHi(1 To nPM): ReDim pmSch(1 To nPM): ReDim pmMin(1 To nPM)
    For i = 1 To nPM
        pmCls(i) = UCase$(Trim$(CStr(v(i, 1)))): pmLo(i) = CStr(v(i, 2)): pmHi(i) = CStr(v(i, 3))
        pmSch(i) = UCase$(CStr(v(i, 4))): pmMin(i) = UCase$(CStr(v(i, 5)))
    Next i
End Sub

' header key: drop a trailing "(unit)" and case/space differences
Private Function NormH(ByVal s As String) As String
    Dim p As Long: s = Trim$(s)
    p = InStr(s, "(")
    If p > 1 Then s = Trim$(Left$(s, p - 1))
    NormH = UCase$(s)
End Function
Private Function GetV(ByVal r As Long, ByVal nm As String) As Variant
    Dim c As Long, i As Long, key As String
    key = NormH(nm)
    For i = 1 To nH
        If hnames(i) = key Then c = i: Exit For
    Next i
    If c = 0 Then GetV = Empty Else GetV = inp(r + 1, c)
End Function
Private Function GetN(ByVal r As Long, ByVal nm As String, ByRef has As Boolean) As Double
    Dim v As Variant: v = GetV(r, nm)
    has = False
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) And Len(CStr(v)) > 0 Then GetN = CDbl(v): has = True
End Function
Private Function GetS(ByVal r As Long, ByVal nm As String) As String
    Dim v As Variant: v = GetV(r, nm)
    If IsEmpty(v) Or IsError(v) Then GetS = "" Else GetS = Trim$(CStr(v))
End Function

'==========================================================================
'  Pipe ID from PMS class + bore  (mirrors resolve_id)
'==========================================================================
Private Function NpsIndex(ByVal bore As Double) As Long
    Dim i As Long, best As Long, bd As Double: bd = 1E+30
    For i = 1 To nOD
        If Abs(odNum(i) - bore) < bd Then bd = Abs(odNum(i) - bore): best = i
    Next i
    If bd < 0.13 Then NpsIndex = best
End Function
Private Function WallFor(ByVal nps As String, ByVal sch As String) As Double
    Dim i As Long, alts As Variant, a As Variant
    Select Case sch
        Case "40": alts = Array("40", "STD", "40S")
        Case "STD": alts = Array("STD", "40", "40S")
        Case "40S": alts = Array("40S", "STD", "40")
        Case "80": alts = Array("80", "XS", "80S")
        Case "XS": alts = Array("XS", "80", "80S")
        Case "80S": alts = Array("80S", "XS", "80")
        Case "10S": alts = Array("10S", "10")
        Case "10": alts = Array("10", "10S")
        Case Else: alts = Array(sch)
    End Select
    For Each a In alts
        For i = 1 To nPW
            If pwNps(i) = nps And pwSch(i) = a Then WallFor = pwWall(i): Exit Function
        Next i
    Next a
End Function
Private Function ResolveID(ByVal spec As String, ByVal bore As Double, ByRef nps As String, _
                           ByRef sch As String, ByRef basis As String) As Double
    Dim ix As Long, od As Double, wall As Double, ov As String, p As Long, i As Long, j As Long
    Dim jlo As Long, jhi As Long, t As String
    ix = NpsIndex(bore)
    If ix = 0 Then basis = "bore not recognised": Exit Function
    nps = odNps(ix): od = odVal(ix)
    p = InStr(spec, "/")
    If p > 0 Then ov = UCase$(Replace(Trim$(Mid$(spec, p + 1)), "SCH", "")): spec = Trim$(Left$(spec, p - 1))
    sch = ""
    For i = 1 To nPM
        If pmCls(i) = UCase$(spec) Then
            jlo = 0: jhi = 0
            For j = 1 To nOD
                If odNps(j) = pmLo(i) Then jlo = j
                If odNps(j) = pmHi(i) Then jhi = j
            Next j
            If jlo > jhi Then j = jlo: jlo = jhi: jhi = j
            If ix >= jlo And ix <= jhi Then
                sch = pmSch(i)
                If sch = "CAL" Then sch = pmMin(i): If sch = "" Then sch = "STD"
                Exit For
            End If
        End If
    Next i
    basis = spec & " " & nps & """"
    If Len(ov) > 0 Then sch = ov: basis = basis & " Sch " & ov & " (override)"
    If Len(sch) = 0 Then sch = "STD": basis = basis & " (spec n/a -> STD)"
    If IsNumeric(sch) Then
        If CDbl(sch) > 0.01 And CDbl(sch) < 3 And InStr(sch, ".") > 0 Then
            wall = CDbl(sch): ResolveID = od - 2 * wall: Exit Function
        End If
    End If
    wall = WallFor(nps, sch)
    If wall = 0 Then wall = WallFor(nps, "STD"): basis = basis & " (sched n/a -> STD)"
    If wall = 0 Then basis = basis & " (no wall)": Exit Function
    ResolveID = Round(od - 2 * wall, 4)
End Function

Private Function FitK(ByVal nm As String) As Double
    Dim i As Long, s As String: s = LCase$(Trim$(nm))
    For i = 1 To nK
        If kNames(i) = s Then FitK = kVals(i): Exit Function
    Next i
End Function

'==========================================================================
'  Hydraulic kernels
'==========================================================================
Private Function Darcy(ByVal Re As Double, ByVal eod As Double) As Double
    If Re <= 0 Then Exit Function
    If Re < 2000 Then Darcy = 64# / Re: Exit Function
    Dim a As Double: a = Log(eod / 3.7 + 5.74 / Re ^ 0.9) / Log(10#)
    Darcy = 0.25 / (a * a)
End Function

Private Function GasRho(ByVal pPsia As Double) As Double
    If sMW > 0 Then
        GasRho = pPsia * PSIA_TO_PA * (sMW / 1000#) / (sZ * R_UNIV * ((sT + 459.67) / 1.8))
    Else
        GasRho = sVrho * LBFT3_TO_KGM3 * pPsia / sPst
    End If
End Function

' Single/two-phase dP kernel.  Returns rho, v, Re, f, dpF, dpK, dpZ (Pa) through ByRef.
Private Sub DpKernel(ByVal idM As Double, ByVal lenM As Double, ByVal dzM As Double, ByVal k As Double, _
                     ByVal pPsia As Double, ByRef rho As Double, ByRef v As Double, ByRef Re As Double, _
                     ByRef f As Double, ByRef dpF As Double, ByRef dpK As Double, ByRef dpZ As Double)
    Dim area As Double, vm As Double, lm As Double, md As Double, x As Double
    Dim rg As Double, rl As Double, mg As Double, ml As Double, mu As Double, gf As Double, dyn As Double
    area = PI / 4# * idM * idM
    vm = sVap * gMult * LBHR_TO_KGS: lm = sLiq * gMult * LBHR_TO_KGS: md = vm + lm
    rho = 0: v = 0: Re = 0: f = 0: dpF = 0: dpK = 0: dpZ = 0
    If md <= 0 Or area <= 0 Then Exit Sub
    rl = sLrho * LBFT3_TO_KGM3: rg = GasRho(pPsia)
    mg = sVmu * 0.001: ml = sLmu * 0.001: x = vm / md
    Select Case sPhase
        Case "LIQUID": rho = rl: mu = ml
        Case "VAPOR": rho = rg: mu = mg
        Case Else
            If rg <= 0 Or rl <= 0 Then Exit Sub
            rho = 1# / (x / rg + (1# - x) / rl)
            If mg > 0 And ml > 0 Then mu = 1# / (x / mg + (1# - x) / ml) Else mu = mg + ml
    End Select
    If rho <= 0 Then Exit Sub
    gf = md / area: v = gf / rho
    If mu > 0 Then Re = gf * idM / mu
    f = Darcy(Re, ROUGH_M / idM)
    dyn = gf * gf / (2# * rho)
    dpF = f * (lenM / idM) * dyn: dpK = k * dyn: dpZ = rho * G_SI * dzM
End Sub

'==========================================================================
'  Circuit solve
'==========================================================================
Private Function LoadStream(ByVal r1 As Long, ByVal r2 As Long) As Boolean
    Dim r As Long, nm As String, has As Boolean, ws As Worksheet, i As Long, c As Long
    Dim hdr() As String, sv As Variant, lr As Long, lc As Long, rr As Long
    nm = ""
    For r = r1 To r2: If Len(GetS(r, "Stream Lookup")) > 0 Then nm = GetS(r, "Stream Lookup"): Exit For
    Next r
    sStream = nm
    If Len(nm) > 0 Then
        Set ws = Worksheets("Streams")
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row: lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        sv = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value
        ReDim hdr(1 To lc)
        For i = 1 To lc: hdr(i) = LCase$(Trim$(CStr(sv(1, i)))): Next i
        For rr = 2 To lr
            If UCase$(Trim$(CStr(sv(rr, 1)))) = UCase$(nm) Then Exit For
        Next rr
        If rr > lr Then Err.Raise 514, , "Stream '" & nm & "' not found on Streams sheet"
        sT = SV2(sv, rr, hdr, "temp_f"): sPst = SV2(sv, rr, hdr, "pres_psia")
        sVap = SV2(sv, rr, hdr, "vap_mass"): sLiq = SV2(sv, rr, hdr, "liq_mass")
        Dim tm As Double: tm = SV2(sv, rr, hdr, "total_mass")
        sPhase = UCase$(CStr(SVs(sv, rr, hdr, "phase")))
        If sVap = 0 And sLiq = 0 And tm > 0 Then
            If InStr(sPhase, "LIQUID") > 0 And InStr(sPhase, "VAPOR") = 0 Then sLiq = tm Else sVap = tm
        End If
        sMW = SV2(sv, rr, hdr, "mol_weight"): If sMW = 0 Then sMW = SV2(sv, rr, hdr, "vap_mw")
        sVrho = SV2(sv, rr, hdr, "vap_density"): sLrho = SV2(sv, rr, hdr, "liq_density")
        sVmu = SV2(sv, rr, hdr, "vap_visc"): sLmu = SV2(sv, rr, hdr, "liq_visc")
        sZ = SV2(sv, rr, hdr, "vap_z"): If sZ = 0 Then sZ = SV2(sv, rr, hdr, "total_z")
        sGam = SV2(sv, rr, hdr, "vap_cp_cv")
    Else
        ' manual properties - first row of the circuit that carries them
        r = r1
        For i = r1 To r2: GetN i, "Temp (degF)", has: If has Then r = i: Exit For
        Next i
        sT = GetN(r, "Temp (degF)", has): sVap = GetN(r, "Vapor Mass Flow", has): sLiq = GetN(r, "Liq Mass Flow (lb/h)", has)
        sMW = GetN(r, "Vap MW", has): sVrho = GetN(r, "Vap Density (lb/ft3)", has): sLrho = GetN(r, "Liq Density (lb/ft3)", has)
        sVmu = GetN(r, "Vap Visc (cP)", has): sLmu = GetN(r, "Liq Visc (cP)", has)
        sZ = GetN(r, "Vap Z", has): sGam = GetN(r, "Vap Cp/Cv", has): sPst = 0: sPhase = ""
        sStream = "(manual)"
    End If
    If sZ <= 0 Then sZ = 1#
    If sGam <= 0 Then sGam = 1.3
    Dim tot As Double, xq As Double
    tot = sVap + sLiq
    If tot <= 0 Then Err.Raise 515, , "Circuit stream has zero mass flow (check Stream Lookup / Streams sheet)"
    xq = sVap / tot
    If xq > 0.005 And xq < 0.995 And sVrho > 0 And sLrho > 0 Then
        sPhase = "TWO-PHASE"
    ElseIf xq >= 0.5 Then
        sPhase = "VAPOR"
    Else
        sPhase = "LIQUID"
    End If
    LoadStream = True
End Function
Private Function SV2(sv As Variant, ByVal rr As Long, hdr() As String, ByVal key As String) As Double
    Dim c As Long, i As Long
    For i = 1 To UBound(hdr)
        If hdr(i) = key Then c = i: Exit For
    Next i
    If c > 0 Then If IsNumeric(sv(rr, c)) And Not IsEmpty(sv(rr, c)) Then SV2 = CDbl(sv(rr, c))
End Function
Private Function SVs(sv As Variant, ByVal rr As Long, hdr() As String, ByVal key As String) As String
    Dim c As Long, i As Long
    For i = 1 To UBound(hdr)
        If hdr(i) = key Then c = i: Exit For
    Next i
    If c > 0 Then SVs = CStr(sv(rr, c))
End Function

Private Function IsFit(ByVal r As Long, ByVal nm As String) As Boolean
    IsFit = (LCase$(GetS(r, "Fitting Name")) = nm)
End Function

Private Sub SolveCircuit(ByVal r1 As Long, ByVal r2 As Long, ByVal cid As String)
    Dim r As Long, has As Boolean, startP As Double, srcSet As Boolean, dstP As Double, hasDst As Boolean
    Dim p0 As Double, pEnd As Double, it As Long, err1 As Double, guess As Double, note As String
    If Not LoadStream(r1, r2) Then Exit Sub
    startP = 0: srcSet = False: hasDst = False
    For r = r1 To r2
        If IsFit(r, "source") Then
            startP = GetN(r, "Set P (psia)", has): If has Then srcSet = True
            Exit For
        End If
    Next r
    For r = r2 To r1 Step -1
        If IsFit(r, "destination") Then dstP = GetN(r, "Set P (psia)", hasDst): Exit For
    Next r
    Dim typed As Boolean, sp2 As Double
    If Not srcSet Then
        For r = r1 To r2
            sp2 = GetN(r, "Start P (psia)", has): If has Then startP = sp2: typed = True: Exit For
        Next r
        If Not typed Then startP = sPst
    End If
    Dim nStart As Long: nStart = nOut
    If Not srcSet And hasDst Then
        ' ---- back-calculate the source / pump-discharge pressure --------------
        guess = dstP + 25
        For it = 1 To 14
            pEnd = March(r1, r2, cid, guess, False, 0)
            err1 = dstP - pEnd
            If Abs(err1) < 0.001 Then Exit For
            guess = guess + err1: If guess < dstP Then guess = dstP
        Next it
        startP = guess
        note = "Source P back-calculated from Destination Set P = " & Format(dstP, "0.000") & " psia"
    End If
    pEnd = March(r1, r2, cid, startP, True, 0)
    If Len(note) > 0 And nOut > nStart Then out(nStart + 1, 24) = note & " | " & out(nStart + 1, 24)
End Sub

' ---------------------------------------------------------------------------
'  March pressure through rows r1..r2.  write=True records output rows.
'  depth>0 : nested solve (flow-control valves float to the destination)
' ---------------------------------------------------------------------------
Private Function March(ByVal r1 As Long, ByVal r2 As Long, ByVal cid As String, ByVal pStart As Double, _
                       ByVal wr As Boolean, ByVal depth As Long) As Double
    Dim r As Long, p As Double, pOut As Double, cum As Double, lastBore As Double, has As Boolean
    Dim fit As String, fl As String, spec As String, bore As Double, lenFt As Double, dzFt As Double
    Dim fixedDp As Double, hasFD As Boolean, setP As Double, hasSP As Boolean, fixK As Double, hasFK As Boolean
    Dim instDp As Double, hasID As Boolean, idIn As Double, nps As String, sch As String, basis As String
    Dim rho As Double, v As Double, Re As Double, f As Double, eF As Double, eK As Double, eZ As Double
    Dim dPf As Double, dPk As Double, dPz As Double, dPt As Double, k As Double, note As String, d As Double
    Dim ctype As String, dstP As Double, hasDst As Boolean, rr As Long, lo As Double, hi As Double, mdp As Double
    Dim j As Long, pe As Double, isCV As Boolean, bdp As Double, tmp As String
    p = pStart
    For r = r1 To r2
        fit = GetS(r, "Fitting Name"): fl = LCase$(fit)
        If Len(fit) = 0 And Len(GetS(r, "Comp ID")) = 0 Then GoTo NextRow
        spec = GetS(r, "Piping Spec")
        bore = GetN(r, "Bore (in)", has): If has Then lastBore = bore Else bore = lastBore
        lenFt = GetN(r, "Length (ft)", has): dzFt = GetN(r, "Elev Change (ft)", has)
        fixedDp = GetN(r, "Fixed dP (psi)", hasFD): setP = GetN(r, "Set P (psia)", hasSP)
        fixK = GetN(r, "Fixed K", hasFK): instDp = GetN(r, "Instr dP (psi)", hasID)
        idIn = 0: nps = "": sch = "": basis = ""
        If bore > 0 Then idIn = ResolveID(spec, bore, nps, sch, basis)
        rho = 0: v = 0: Re = 0: f = 0: dPf = 0: dPk = 0: dPz = 0: dPt = 0: note = "": isCV = False
        If fl = "fix pressure" Then
            dPt = fixedDp: note = "Fix Pressure"
        ElseIf fl = "orifice" Or UCase$(GetS(r, "Instr Type")) = "FO" Or UCase$(GetS(r, "Instr Type")) = "RO" Then
            dPt = instDp: note = "Orifice " & GetS(r, "Instr Tag") & ": Instr dP"
        ElseIf UCase$(GetS(r, "Instr Type")) = "FT" And hasID Then
            dPt = instDp: note = "Inline FT " & GetS(r, "Instr Tag") & ": rated dP"
        ElseIf fl = "source" Then
            If hasSP Then
                dPt = p - setP: note = "Source: anchor P = " & Format(setP, "0.000") & " psia"
            Else
                note = "Source"
            End If
        ElseIf fl = "destination" Then
            If hasSP Then
                note = "Destination: target " & Format(setP, "0.000") & ", arrival " & Format(p, "0.000")
                note = note & " (margin " & Format(p - setP, "0.000") & ")"
            Else
                note = "Destination"
            End If
        ElseIf fl = "control valve" Then
            isCV = True
            ctype = UCase$(Left$(GetS(r, "Control Valve Type") & "F", 1))
            If InStr("PTFL", ctype) = 0 Then ctype = "F"
            hasDst = False
            If ctype = "F" Then
                For rr = r + 1 To r2
                    If LCase$(GetS(rr, "Fitting Name")) = "destination" Then dstP = GetN(rr, "Set P (psia)", hasDst): Exit For
                Next rr
            End If
            If Not hasFD And hasSP Then
                bdp = p - setP: If bdp < 0 Then bdp = 0
                tmp = "dP from Set P (outlet) " & Format(setP, "0.000")
            ElseIf hasDst And Not hasFD Then
                If depth > 0 Then
                    bdp = p - dstP: If bdp < 0 Then bdp = 0
                Else
                    lo = 0: hi = p: If hi < 1 Then hi = 1
                    For j = 1 To 40
                        mdp = (lo + hi) / 2
                        pe = March(r + 1, r2, cid, p - mdp, False, depth + 1)
                        If pe > dstP Then lo = mdp Else hi = mdp
                        If hi - lo < 0.0005 Then Exit For
                    Next j
                    bdp = (lo + hi) / 2
                End If
                tmp = "flow control: dP solved for Destination " & Format(dstP, "0.0") & " psia"
            ElseIf hasFD Then
                bdp = fixedDp: tmp = "fixed design dP"
            ElseIf hasFK And idIn > 0 Then
                DpKernel idIn * IN_TO_M, 0, 0, fixK, p, rho, v, Re, f, eF, eK, eZ
                bdp = (eF + eK + eZ) * PA_TO_PSI: tmp = "K = " & fixK
            Else
                bdp = 0: tmp = "no dP basis"
            End If
            dPt = bdp: note = "Control Valve [" & ctype & "]: " & tmp
            If wr Then RecordCV r, cid, p, bdp, tmp
        Else
            If idIn <= 0 Then
                note = basis
            Else
                k = FitK(fit): If fl = "fix k" And hasFK Then k = fixK
                d = idIn * IN_TO_M
                DpKernel d, lenFt * FT_TO_M, dzFt * FT_TO_M, k, p, rho, v, Re, f, eF, eK, eZ
                dPf = eF * PA_TO_PSI: dPk = eK * PA_TO_PSI: dPz = eZ * PA_TO_PSI
                dPt = dPf + dPk + dPz: note = basis
            End If
        End If
        pOut = p - dPt: cum = cum + lenFt
        If pOut < 0.5 Then
            note = note & "  | WARNING: P < 0.5 psia - infeasible / choked, clamped"
            pOut = 0.5
        End If
        If wr Then
            nOut = nOut + 1
            out(nOut, 1) = cid
            out(nOut, 2) = GetS(r, "Line No")
            out(nOut, 3) = GetS(r, "Comp ID")
            out(nOut, 4) = fit: out(nOut, 5) = spec: out(nOut, 6) = nps: out(nOut, 7) = sch
            out(nOut, 8) = IIf(idIn > 0, idIn, ""): out(nOut, 9) = lenFt: out(nOut, 10) = dzFt: out(nOut, 11) = cum
            out(nOut, 12) = p: out(nOut, 13) = pOut: out(nOut, 14) = v / FT_TO_M: out(nOut, 15) = Re: out(nOut, 16) = f
            out(nOut, 17) = rho / LBFT3_TO_KGM3: out(nOut, 18) = dPf: out(nOut, 19) = dPk: out(nOut, 20) = dPz
            out(nOut, 21) = dPt: out(nOut, 22) = sPhase: out(nOut, 23) = sStream: out(nOut, 24) = note
        End If
        p = pOut
NextRow:
    Next r
    March = p
End Function

'==========================================================================
'  Control-valve datasheet (ISA 75.01 / IEC 60534-2-1, simplified)
'==========================================================================
Private Sub RecordCV(ByVal r As Long, ByVal cid As String, ByVal p1 As Double, ByVal dp As Double, ByVal basis As String)
    Dim wl As Double, wv As Double, q As Double, sg As Double, cvv As Double, rho1 As Double
    Dim x As Double, fk As Double, y As Double, FL As Double, p2 As Double, tag As String
    FL = 0.9: p2 = p1 - dp: tag = GetS(r, "Instr Tag"): If tag = "" Then tag = GetS(r, "Comp ID")
    wl = sLiq * gMult: wv = sVap * gMult
    nCV = nCV + 1
    cvOut(nCV, 1) = cid: cvOut(nCV, 2) = tag: cvOut(nCV, 3) = sStream: cvOut(nCV, 4) = sPhase
    cvOut(nCV, 5) = p1: cvOut(nCV, 6) = p2: cvOut(nCV, 7) = dp: cvOut(nCV, 8) = wl + wv: cvOut(nCV, 9) = sT
    If dp <= 0 Then cvOut(nCV, 13) = "dP = 0: valve fully open / no dP basis": Exit Sub
    If sPhase = "LIQUID" Then
        q = wl / (sLrho * 7.48052 * 60#): sg = sLrho / 62.428
        cvOut(nCV, 10) = q * Sqr(sg / dp): cvOut(nCV, 11) = sg
        cvOut(nCV, 12) = "gpm = " & Format(q, "0.0")
        cvOut(nCV, 13) = "Liquid Cv = Q*sqrt(SG/dP); FL=0.9 assumed; check choking (Pv) with fluid data"
    ElseIf sPhase = "VAPOR" Then
        rho1 = GasRho(p1) / LBFT3_TO_KGM3
        fk = sGam / 1.4: x = dp / p1
        If x > fk * 0.7 Then x = fk * 0.7: cvOut(nCV, 13) = "CHOKED (x>Fk*xT, xT=0.7 assumed). "
        y = 1# - x / (3# * fk * 0.7)
        cvOut(nCV, 10) = wv / (63.3 * y * Sqr(x * p1 * rho1)): cvOut(nCV, 11) = rho1
        cvOut(nCV, 12) = "rho1 = " & Format(rho1, "0.000") & " lb/ft3"
        cvOut(nCV, 13) = cvOut(nCV, 13) & "Gas Cv (ISA 75.01 W-form), Fk=gamma/1.4, xT=0.7 assumed"
    Else
        cvOut(nCV, 13) = "Two-phase valve: size per vendor / HEM method (not computed in v1)"
    End If
    If IsNumeric(cvOut(nCV, 10)) And Not IsEmpty(cvOut(nCV, 10)) Then cvOut(nCV, 14) = cvOut(nCV, 10) / 0.7
End Sub

'==========================================================================
'  Output sheets
'==========================================================================
Private Sub WriteOutputs()
    Dim ws As Worksheet, hdr As Variant, i As Long, nc As Long
    hdr = Array("Circuit", "Line No", "Comp ID", "Fitting", "Spec", "NPS", "Sch", "ID (in)", "Length (ft)", "Elev (ft)", _
        "Cum L (ft)", "P in (psia)", "P out (psia)", "Velocity (ft/s)", "Re", "f", "Density (lb/ft3)", "dP fric (psi)", _
        "dP fit (psi)", "dP elev (psi)", "dP total (psi)", "Phase", "Stream", "Note")
    Set ws = FreshSheet("Pressure_Profile")
    ws.Range("A1").Resize(1, 24).Value = hdr
    StyleHeader ws.Range("A1").Resize(1, 24)
    If nOut > 0 Then ws.Range("A2").Resize(nOut, 24).Value = out
    ws.Range("I2:K" & nOut + 1).NumberFormat = "0.00"
    ws.Range("L2:M" & nOut + 1).NumberFormat = "0.000"
    ws.Range("N2:N" & nOut + 1).NumberFormat = "0.00"
    ws.Range("O2:O" & nOut + 1).NumberFormat = "#,##0"
    ws.Range("P2:P" & nOut + 1).NumberFormat = "0.0000"
    ws.Range("Q2:Q" & nOut + 1).NumberFormat = "0.000"
    ws.Range("R2:U" & nOut + 1).NumberFormat = "0.000"
    ws.Columns("A:W").ColumnWidth = 12: ws.Columns("D").ColumnWidth = 24: ws.Columns("X").ColumnWidth = 70
    ws.Activate: ActiveWindow.FreezePanes = False: ws.Range("D2").Select: ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    hdr = Array("Circuit", "CV Tag", "Stream", "Phase", "P1 (psia)", "P2 (psia)", "dP (psi)", "Flow (lb/h)", "T (degF)", _
        "Calc Cv", "SG / rho1", "Flow info", "Notes", "Suggested rated Cv (70% travel)")
    Set ws = FreshSheet("CV_Summary")
    ws.Range("A1").Resize(1, 14).Value = hdr
    StyleHeader ws.Range("A1").Resize(1, 14)
    If nCV > 0 Then ws.Range("A2").Resize(nCV, 14).Value = cvOut
    ws.Range("E2:G" & nCV + 2).NumberFormat = "0.000": ws.Range("J2:K" & nCV + 2).NumberFormat = "0.00"
    ws.Range("N2:N" & nCV + 2).NumberFormat = "0.0"
    ws.Columns("A:N").ColumnWidth = 13: ws.Columns("M").ColumnWidth = 60
    Worksheets("Control").Activate
End Sub
Private Function FreshSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        ws.Name = nm
    Else
        ws.Cells.Clear
    End If
    Set FreshSheet = ws
End Function
Private Sub StyleHeader(rg As Range)
    rg.Interior.Color = RGB(0, 23, 66): rg.Font.Color = RGB(255, 255, 255): rg.Font.Bold = True
    rg.WrapText = True: rg.VerticalAlignment = xlCenter
End Sub
