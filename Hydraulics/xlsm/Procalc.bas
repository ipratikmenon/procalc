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
Private Type TPI
    vsg As Double
    vsl As Double
    vm As Double
    rg As Double
    rl As Double
    mug As Double
    mul As Double
    sig As Double
    d As Double
    th As Double
    gvf As Double
End Type

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
Private pmMoc As Variant, stLine() As String
Private pmCls As Variant, pmLo As Variant, pmHi As Variant, pmSch As Variant, pmMin As Variant, nPM As Long

'---- output buffer -----------------------------------------------------------
Private out() As Variant, nOut As Long
Private cvOut() As Variant, nCV As Long
Private gMult As Double, hdrRow As Long
'---- flash (composition) model --------------------------------------------
Private useFlash As Boolean, flashOn As Boolean, fN As Long
Private fNm() As String, fZ() As Double, fKref() As Double, fMWc() As Double
Private fTcR() As Double, fOm() As Double, fHasC As Boolean
Private fPref As Double, fTref As Double, fTotMoles As Double, fBeta0 As Double, fSrc As String
Private cN() As String, cMW() As Double, cTc() As Double, cPc() As Double, cOm() As Double, nCC As Long
Private sVMW As Double, sLMW As Double, sSig As Double
Private eVap As Double, eLiq As Double, eMW As Double, eZ As Double, eMuG As Double, eMuL As Double
Private eK() As Double, eX() As Double, csRow As Long, cdCol As Long, cdNew As Boolean, stnNo As Long
Private ePhase As String, eBeta As Double, eGam As Double, eY() As Double, eNote As String

'==========================================================================
'  Entry point - assigned to the RUN button on the Control sheet
'==========================================================================
'---- module-level state (moved here: declarations must precede procedures) ----
Private stN As Long
Private stSeq() As String, stFit() As String, stCid() As String, stComp() As String, stNps() As String
Private stOD() As Double, stWall() As Double, stID() As Double, stRho() As Double, stV() As Double
Private stRg() As Double, stRl() As Double, stMg() As Double, stMl() As Double, stMu() As Double
Private stMach() As Double, stRv2V() As Double, stRv2L() As Double, stCum() As Double, stPin() As Double, stDp() As Double
Private stHasMom() As Boolean
Private fivOut() As Variant, nFiv As Long, aivOut() As Variant, nAiv As Long, lastOD As Double, lastWall As Double
Private mdotCirc As Double
Private regOut() As Variant, nReg As Long
Private stLen() As Double, stDz() As Double, stSig() As Double, stMuL() As Double
Private idat As Variant, nID As Long, tcls As Variant, nTC As Long
Private stKind() As String, stTag() As String, stSpec() As String, stT() As Double, stGm() As Double
Private instOut() As Variant, nInst As Long, twOut() As Variant, nTw As Long
Private rowMap() As Long, rowSeq() As Long
Private fScale As Double, vCf As Double, lCf As Double, capSplit As Boolean, useInj As Boolean
Private nSplit As Long, spLine(1 To 60) As String, spSeq(1 To 60) As Long, spN(1 To 60) As Long, spV(1 To 60) As Double, spL(1 To 60) As Double
Private nInj As Long, injRow(1 To 20) As Long, injLbl(1 To 20) As String, injVap(1 To 20) As Double, injLiq(1 To 20) As Double
Private injN(1 To 20) As Long, injPref(1 To 20) As Double, injFlash(1 To 20) As Boolean
Private injNm(1 To 20, 1 To 600) As String, injMol(1 To 20, 1 To 600) As Double, injK(1 To 20, 1 To 600) As Double
Private injMW(1 To 20, 1 To 600) As Double, injTc(1 To 20, 1 To 600) As Double, injOm(1 To 20, 1 To 600) As Double
Private bkN(0 To 3) As Long, bkNm(0 To 3, 1 To 600) As String, bkZ(0 To 3, 1 To 600) As Double, bkK(0 To 3, 1 To 600) As Double
Private bkMW(0 To 3, 1 To 600) As Double, bkTc(0 To 3, 1 To 600) As Double, bkOm(0 To 3, 1 To 600) As Double
Private bkHas(0 To 3) As Boolean, bkPref(0 To 3) As Double, bkTref(0 To 3) As Double, bkTot(0 To 3) As Double
Private bkVap(0 To 3) As Double, bkLiq(0 To 3) As Double, bkFlash(0 To 3) As Boolean

Public Sub RunHydraulics()
    Dim t0 As Double: t0 = Timer
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    LoadTables
    gMult = Val(Worksheets("Control").Range("C4").Value) / 100#
    useFlash = (UCase$(Trim$(CStr(Worksheets("Control").Range("C7").Value))) <> "OFF")
    ReadConstants
    ReadInstData
    InitCompSheets
    If gMult <= 0 Then gMult = 1#
    ReDim out(1 To nRow + 10, 1 To 28): nOut = 0
    ReDim cvOut(1 To nRow + 10, 1 To 20): nCV = 0
    InitStations

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
    WriteFivAiv
    WriteRegime
    WriteInstr
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
    v = ws.Range("A2:F" & lr).Value: nPM = lr - 1
    ReDim pmMoc(1 To nPM): ReDim pmCls(1 To nPM): ReDim pmLo(1 To nPM): ReDim pmHi(1 To nPM): ReDim pmSch(1 To nPM): ReDim pmMin(1 To nPM)
    For i = 1 To nPM
        pmCls(i) = UCase$(Trim$(CStr(v(i, 1)))): pmLo(i) = CStr(v(i, 2)): pmHi(i) = CStr(v(i, 3))
        pmSch(i) = UCase$(CStr(v(i, 4))): pmMin(i) = UCase$(CStr(v(i, 5))): pmMoc(i) = CStr(v(i, 6))
    Next i
    ComputeSeq
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
    lastOD = 0: lastWall = 0
    ix = NpsIndex(bore)
    If ix = 0 Then basis = "bore not recognised": Exit Function
    nps = odNps(ix): od = odVal(ix): lastOD = od
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
            wall = CDbl(sch): lastWall = wall: ResolveID = od - 2 * wall: Exit Function
        End If
    End If
    wall = WallFor(nps, sch)
    If wall = 0 Then wall = WallFor(nps, "STD"): basis = basis & " (sched n/a -> STD)"
    If wall = 0 Then basis = basis & " (no wall)": Exit Function
    lastWall = wall
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
    If eMW > 0 Then
        GasRho = pPsia * PSIA_TO_PA * (eMW / 1000#) / (eZ * R_UNIV * ((sT + 459.67) / 1.8))
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
    vm = eVap * LBHR_TO_KGS: lm = eLiq * LBHR_TO_KGS: md = vm + lm
    rho = 0: v = 0: Re = 0: f = 0: dpF = 0: dpK = 0: dpZ = 0
    If md <= 0 Or area <= 0 Then Exit Sub
    rl = sLrho * LBFT3_TO_KGM3: rg = GasRho(pPsia)
    mg = eMuG * 0.001: ml = eMuL * 0.001: x = vm / md
    If flashOn Then
        If x <= 0.00001 Or rg <= 0 Then
            rho = rl: mu = ml
        ElseIf x >= 0.99999 Or rl <= 0 Then
            rho = rg: mu = mg
        Else
            rho = 1# / (x / rg + (1# - x) / rl)
            If mg > 0 And ml > 0 Then mu = 1# / (x / mg + (1# - x) / ml) Else mu = mg + ml
        End If
    Else
        Select Case ePhase
            Case "LIQUID": rho = rl: mu = ml
            Case "VAPOR": rho = rg: mu = mg
            Case Else
                If rg <= 0 Or rl <= 0 Then Exit Sub
                rho = 1# / (x / rg + (1# - x) / rl)
                If mg > 0 And ml > 0 Then mu = 1# / (x / mg + (1# - x) / ml) Else mu = mg + ml
        End Select
    End If
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
'  PHASE 2 - composition + flash  (ports of _build_feed / flash / rachford_rice)
'==========================================================================
Private Function NumOf(ByVal v As Variant, ByRef ok As Boolean) As Double
    ok = False
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If IsNumeric(v) Then NumOf = CDbl(v): ok = True
End Function

' Component_Constants sheet = COMP_CONSTANTS layout: B Name, E MW, I Tc(F), J Pc(psia), M omega
Private Sub ReadConstants()
    Dim ws As Worksheet, v As Variant, i As Long, r0 As Long, lr As Long, ok As Boolean
    nCC = 0
    On Error Resume Next
    Set ws = Worksheets("Component_Constants")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    r0 = 0
    For i = 1 To 6
        If UCase$(Trim$(CStr(ws.Cells(i, 2).Value))) = "NAME" Then r0 = i: Exit For
    Next i
    If r0 = 0 Or lr <= r0 Then Exit Sub
    v = ws.Range(ws.Cells(r0 + 1, 2), ws.Cells(lr, 13)).Value
    nCC = UBound(v, 1)
    ReDim cN(1 To nCC): ReDim cMW(1 To nCC): ReDim cTc(1 To nCC): ReDim cPc(1 To nCC): ReDim cOm(1 To nCC)
    For i = 1 To nCC
        cN(i) = UCase$(Trim$(CStr(v(i, 1))))
        cMW(i) = NumOf(v(i, 4), ok): If Not ok Then cMW(i) = 0
        cTc(i) = NumOf(v(i, 8), ok): If Not ok Then cTc(i) = -9999
        cPc(i) = NumOf(v(i, 9), ok): If Not ok Then cPc(i) = 0
        cOm(i) = NumOf(v(i, 12), ok): If Not ok Then cOm(i) = -9999
    Next i
End Sub

Private Function CIdx(ByVal nm As String) As Long
    Dim i As Long: nm = UCase$(Trim$(nm))
    For i = 1 To nCC
        If cN(i) = nm Then CIdx = i: Exit Function
    Next i
End Function

Private Function FindStreamSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet, t As String, p As Long, q As Long
    For Each ws In Worksheets
        If UCase$(ws.Name) = UCase$(nm) Then Set FindStreamSheet = ws: Exit Function
    Next ws
    For Each ws In Worksheets
        t = CStr(ws.Cells(1, 2).Value)
        p = InStr(UCase$(t), "STREAM:")
        If p > 0 Then
            q = InStr(p, t, "|")
            If q = 0 Then q = Len(t) + 1
            If UCase$(Trim$(Mid$(t, p + 7, q - p - 7))) = UCase$(nm) Then Set FindStreamSheet = ws: Exit Function
        End If
    Next ws
End Function

Private Function IsSecHdr(ByVal lab As String) As Boolean
    Dim p As Long: p = InStr(lab, ".")
    If p < 2 Then Exit Function
    IsSecHdr = IsNumeric(Left$(lab, p - 1))
End Function

Private Function PickCol(v As Variant, ByVal i As Long, ByVal c1 As Long, ByVal c2 As Long, ByRef ok As Boolean) As Double
    Dim x As Double
    PickCol = NumOf(v(i, c1), ok)
    If ok Then Exit Function
    If c2 > 0 Then PickCol = NumOf(v(i, c2), ok)
End Function

' Reads an HMB per-stream sheet: properties -> s* globals, compositions -> BuildFeed
Private Function ReadStreamSheet(ByVal nm As String) As Boolean
    Dim ws As Worksheet, v As Variant, lr As Long, i As Long, j As Long, ok As Boolean, lab As String, ll As String
    Dim unit As String, sec As String, comp As Long, cT As Long, cV As Long, cL As Long, t As String, p As Long, q As Long
    Dim tot As Double, vp As Double, lq As Double, gotV As Boolean, gotL As Boolean, gotZ As Boolean, x As Double
    Dim secN(1 To 5) As Long, secNm(1 To 5, 1 To 600) As String, secVal(1 To 5, 1 To 600) As Double
    Set ws = FindStreamSheet(nm)
    If ws Is Nothing Then Exit Function
    lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lr < 4 Then Exit Function
    v = ws.Range(ws.Cells(1, 2), ws.Cells(lr, 6)).Value
    t = CStr(v(1, 1)): sPhase = ""
    p = InStr(t, "Phase:")
    If p > 0 Then
        q = InStr(p, t, "|"): If q = 0 Then q = Len(t) + 1
        sPhase = UCase$(Trim$(Mid$(t, p + 6, q - p - 6)))
    End If
    cT = 3: cV = 4: cL = 5
    For j = 3 To 5
        Select Case UCase$(Trim$(CStr(v(2, j))))
            Case "TOTAL": cT = j
            Case "VAPOR": cV = j
            Case "LIQUID": cL = j
        End Select
    Next j
    sT = 0: sPst = 0: sVap = 0: sLiq = 0: sMW = 0: sVMW = 0: sLMW = 0: sVrho = 0: sLrho = 0
    sVmu = 0: sLmu = 0: sZ = 0: sGam = 0: sSig = 0
    Dim tm As Double, hvT As Boolean, hvV As Boolean, hvL As Boolean
    For i = 3 To lr
        lab = Trim$(CStr(v(i, 1)))
        If Len(lab) = 0 Then GoTo NextI
        unit = UCase$(Trim$(CStr(v(i, 2))))
        If IsSecHdr(lab) Then
            sec = UCase$(lab): comp = 0
            If InStr(sec, "VAPOR MOLE FRACTIONS") > 0 Then
                comp = 3
            ElseIf InStr(sec, "LIQUID MOLE FRACTIONS") > 0 Then
                comp = 2
            ElseIf InStr(sec, "MOLE FLOW") > 0 Then
                comp = 4
            ElseIf InStr(sec, "MASS FLOW") > 0 Then
                comp = 5
            ElseIf InStr(sec, "MOLE FRACTIONS") > 0 Then
                comp = 1
            ElseIf InStr(sec, "MASS FRACTIONS") > 0 Then
                comp = -1
            End If
            GoTo NextI
        End If
        If comp <> 0 Then
            If comp > 0 And LCase$(lab) <> "component" Then
                x = NumOf(v(i, 3), ok)
                If ok And secN(comp) < 600 Then secN(comp) = secN(comp) + 1: secNm(comp, secN(comp)) = UCase$(lab): secVal(comp, secN(comp)) = x
            End If
            GoTo NextI
        End If
        ll = LCase$(lab)
        If InStr(sec, "FLOW RATE") > 0 Then
            If ll = "mass rate" Then
                tm = NumOf(v(i, cT), hvT): vp = NumOf(v(i, cV), hvV): lq = NumOf(v(i, cL), hvL)
                If hvT Then sVap = 0: sLiq = 0: tot = tm
                If hvV Then sVap = vp: gotV = True
                If hvL Then sLiq = lq: gotL = True
            End If
        ElseIf InStr(sec, "CONDITION") > 0 Then
            If InStr(ll, "temperature") > 0 And sT = 0 Then sT = NumOf(v(i, cT), ok)
            If InStr(ll, "pressure") > 0 And sPst = 0 Then sPst = NumOf(v(i, cT), ok)
            If InStr(ll, "molecular weight") > 0 Then
                If sMW = 0 Then sMW = PickCol(v, i, cT, cV, ok)
                If sVMW = 0 Then sVMW = NumOf(v(i, cV), ok)
                If sLMW = 0 Then sLMW = NumOf(v(i, cL), ok)
            End If
        ElseIf InStr(sec, "VAPOR PHASE") > 0 Then
            If InStr(ll, "actual density") > 0 Then sVrho = PickCol(v, i, cV, cT, ok)
            If InStr(ll, "viscosity") > 0 And InStr(ll, "kin") = 0 Then sVmu = PickCol(v, i, cV, cT, ok)
            If (InStr(ll, "z-factor") > 0 Or InStr(ll, "z factor") > 0) And sZ = 0 Then sZ = PickCol(v, i, cV, cT, ok)
            If InStr(ll, "cp/cv") > 0 And sGam = 0 Then sGam = PickCol(v, i, cV, cT, ok)
        ElseIf InStr(sec, "LIQUID PHASE") > 0 Then
            If InStr(ll, "actual density") > 0 Then sLrho = PickCol(v, i, cL, cT, ok)
            If InStr(ll, "viscosity") > 0 And InStr(ll, "kin") = 0 Then sLmu = PickCol(v, i, cL, cT, ok)
            If InStr(ll, "surface tension") > 0 Then sSig = PickCol(v, i, cL, cT, ok)
        End If
NextI:
    Next i
    ' split total mass by phase if only the total is printed (as extract_stream)
    If Not gotV And Not gotL Then
        If InStr(sPhase, "LIQUID") > 0 And InStr(sPhase, "VAPOR") = 0 Then sLiq = tot Else sVap = tot
    ElseIf gotV And Not gotL Then
        sLiq = tot - sVap: If sLiq < 0 Then sLiq = 0
    ElseIf gotL And Not gotV Then
        sVap = tot - sLiq: If sVap < 0 Then sVap = 0
    End If
    If sPst <= 0 Then Exit Function
    sStream = nm & " (sheet)"
    ReadStreamSheet = True
    ' ---- composition -> feed ------------------------------------------------
    flashOn = False
    If secN(1) >= 2 Then flashOn = BuildFeed(secN, secNm, secVal)
End Function

Private Function SecGet(secN() As Long, secNm() As String, secVal() As Double, ByVal sec As Long, ByVal nm As String) As Double
    Dim i As Long
    For i = 1 To secN(sec)
        If secNm(sec, i) = nm Then SecGet = secVal(sec, i): Exit Function
    Next i
End Function

Private Function RR(ByVal n As Long, z() As Double, k() As Double) As Double
    Dim lo As Double, hi As Double, mdp As Double, gm As Double, i As Long, it As Long, g0 As Double, g1 As Double
    For i = 1 To n
        g0 = g0 + z(i) * (k(i) - 1#)
        g1 = g1 + z(i) * (k(i) - 1#) / k(i)
    Next i
    If g0 <= 0 Then RR = 0: Exit Function
    If g1 >= 0 Then RR = 1: Exit Function
    lo = 0: hi = 1
    For it = 1 To 200
        mdp = 0.5 * (lo + hi): gm = 0
        For i = 1 To n
            gm = gm + z(i) * (k(i) - 1#) / (1# + mdp * (k(i) - 1#))
        Next i
        If Abs(gm) < 0.000000000001 Then RR = mdp: Exit Function
        If gm > 0 Then lo = mdp Else hi = mdp
    Next it
    RR = 0.5 * (lo + hi)
End Function

Private Function SolveLambda(ByVal n As Long, z() As Double, kr() As Double, ByVal target As Double) As Double
    Dim lo As Double, hi As Double, mdp As Double, b As Double, it As Long, i As Long
    Dim kk() As Double: ReDim kk(1 To n)
    lo = 0.000000001: hi = 1000000000#
    For i = 1 To n: kk(i) = ClampK(lo * kr(i), 1E+12): Next i
    If RR(n, z, kk) >= target Then SolveLambda = lo: Exit Function
    For i = 1 To n: kk(i) = ClampK(hi * kr(i), 1E+12): Next i
    If RR(n, z, kk) <= target Then SolveLambda = hi: Exit Function
    For it = 1 To 200
        mdp = Sqr(lo * hi)
        For i = 1 To n: kk(i) = ClampK(mdp * kr(i), 1E+12): Next i
        b = RR(n, z, kk)
        If Abs(b - target) < 0.000000001 Then SolveLambda = mdp: Exit Function
        If b < target Then lo = mdp Else hi = mdp
    Next it
    SolveLambda = Sqr(lo * hi)
End Function
Private Function ClampK(ByVal k As Double, ByVal hi As Double) As Double
    If k > hi Then k = hi
    If k < 1 / hi Then k = 1 / hi
    ClampK = k
End Function

Private Function BuildFeed(secN() As Long, secNm() As String, secVal() As Double) As Boolean
    Dim i As Long, n As Long, nm As String, zi As Double, xi As Double, yi As Double, mi As Double, mo As Double
    Dim yOk As Long, useYX As Boolean, vmol As Double, lmol As Double, b0 As Double, vmw As Double, lmw As Double
    Dim kraw() As Double, sz As Double, lam As Double, avg As Double, nk As Long, tmw As Double, tmass As Double, ci As Long
    Dim zz() As Double, kk() As Double
    ' reference molar vapour fraction from the HMB phase masses (_ref_beta)
    vmw = sVMW: If vmw = 0 Then vmw = sMW
    lmw = sLMW: If lmw = 0 Then lmw = sMW
    If vmw > 0 Then vmol = sVap / vmw
    If lmw > 0 Then lmol = sLiq / lmw
    If vmol + lmol > 0 Then b0 = vmol / (vmol + lmol)
    If b0 < 0.0001 Then b0 = 0.0001
    If b0 > 0.9999 Then b0 = 0.9999
    n = 0
    ReDim fNm(1 To secN(1)): ReDim fZ(1 To secN(1)): ReDim kraw(1 To secN(1)): ReDim fMWc(1 To secN(1))
    ReDim fTcR(1 To secN(1)): ReDim fOm(1 To secN(1))
    For i = 1 To secN(1)
        If secVal(1, i) > 0.000000000001 Then
            zi = secVal(1, i): nm = secNm(1, i)
            xi = SecGet(secN, secNm, secVal, 2, nm): yi = SecGet(secN, secNm, secVal, 3, nm)
            If yi > 0 And xi > 0.000000000001 Then yOk = yOk + 1
        End If
    Next i
    useYX = (yOk >= 2)
    For i = 1 To secN(1)
        If secVal(1, i) > 0.000000000001 Then
            n = n + 1
            nm = secNm(1, i): zi = secVal(1, i)
            xi = SecGet(secN, secNm, secVal, 2, nm): yi = SecGet(secN, secNm, secVal, 3, nm)
            mi = SecGet(secN, secNm, secVal, 5, nm): mo = SecGet(secN, secNm, secVal, 4, nm)
            fNm(n) = nm: fZ(n) = zi
            If useYX And yi > 0 And xi > 0.000000000001 Then
                kraw(n) = yi / xi
            ElseIf xi > 0.000000000001 Then
                kraw(n) = zi / xi
            Else
                kraw(n) = 1000000#
            End If
            If mi > 0 And mo > 0 Then fMWc(n) = mi / mo Else fMWc(n) = 0
            ci = CIdx(nm)
            fTcR(n) = -9999: fOm(n) = -9999
            If ci > 0 Then
                If cTc(ci) > -9000 And cOm(ci) > -9000 Then fTcR(n) = cTc(ci) + 459.67: fOm(n) = cOm(ci)
                If fMWc(n) = 0 And cMW(ci) > 0 Then fMWc(n) = cMW(ci)
            End If
        End If
    Next i
    If n < 1 Then Exit Function
    If n = 1 And (sLiq > 0.000000001 Or sVap <= 0) Then Exit Function
    For i = 1 To n
        If fMWc(i) > 0 Then avg = avg + fMWc(i): nk = nk + 1
    Next i
    If nk > 0 Then avg = avg / nk Else avg = IIf(sMW > 0, sMW, 50)
    For i = 1 To n
        If fMWc(i) = 0 Then fMWc(i) = avg
        sz = sz + fZ(i)
    Next i
    ReDim zz(1 To n): ReDim kk(1 To n): ReDim fKref(1 To n)
    For i = 1 To n
        fZ(i) = fZ(i) / sz: zz(i) = fZ(i)
        tmw = tmw + fZ(i) * fMWc(i)
    Next i
    lam = SolveLambda(n, zz, kraw, b0)
    For i = 1 To n
        fKref(i) = ClampK(lam * kraw(i), 10000000000#)
    Next i
    If n = 1 Then fKref(1) = 10000000000#
    tmass = sVap + sLiq
    If tmass <= 0 Or tmw <= 0 Then Exit Function
    fTotMoles = tmass / tmw
    fPref = sPst: fTref = sT: fBeta0 = b0: fN = n
    fHasC = False
    For i = 1 To n
        If fTcR(i) > -9000 Then fHasC = True: Exit For
    Next i
    fSrc = n & " comps, beta0=" & Format(b0, "0.0000") & ", " & IIf(useYX, "y/x", "z/x-lambda") & " K basis"
    BuildFeed = True
End Function

' Isothermal flash at pPsia (optionally Wilson T-corrected): fills eVap/eLiq/eMW/eBeta and eY()
Private Sub FlashAt(ByVal pPsia As Double)
    Dim i As Long, n As Long, ratio As Double, t_r As Double, tref_r As Double, inv As Double, w As Double
    Dim k() As Double, beta As Double, sx As Double, sy As Double, vmw As Double, lmw As Double
    Dim x() As Double, y() As Double, vmol As Double, lmol As Double
    n = fN: ReDim k(1 To n): ReDim x(1 To n): ReDim y(1 To n): ReDim eY(1 To n): ReDim eX(1 To n): ReDim eK(1 To n)
    If pPsia < 0.000001 Then pPsia = 0.000001
    ratio = fPref / pPsia
    t_r = sT + 459.67: tref_r = fTref + 459.67
    inv = 0
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
    For i = 1 To n
        x(i) = x(i) / sx: y(i) = y(i) / sy: eY(i) = y(i): eX(i) = x(i): eK(i) = k(i)
        vmw = vmw + y(i) * fMWc(i): lmw = lmw + x(i) * fMWc(i)
    Next i
    vmol = beta * fTotMoles: lmol = (1# - beta) * fTotMoles
    eVap = vmol * vmw * gMult: eLiq = lmol * lmw * gMult
    eMW = vmw: eBeta = beta
End Sub

' Vapour mu / Z / Cp-Cv for flash-generated vapour when the stream sheet has none (Lucas + JST, Peng-Robinson Z)
Private Sub FillVapour(ByVal pPsia As Double)
    Dim i As Long, n As Long, ci As Long, T As Double, P As Double, ok As Boolean, sy As Double
    Dim idx() As Long, yv() As Double, m As Long
    eNote = ""
    If eVap <= 0 Then Exit Sub
    If sVmu > 0 And sZ >= 0.2 Then Exit Sub
    T = (sT + 459.67) / 1.8: P = pPsia / 14.5038
    n = fN: ReDim idx(1 To n): ReDim yv(1 To n)
    ok = True
    For i = 1 To n
        If eY(i) > 0.000000000001 Then
            ci = CIdx(fNm(i))
            If ci = 0 Then ok = False: Exit For
            If cMW(ci) <= 0 Or cTc(ci) <= -9000 Or cPc(ci) <= 0 Or cOm(ci) <= -9000 Then ok = False: Exit For
            m = m + 1: idx(m) = ci: yv(m) = eY(i): sy = sy + eY(i)
        End If
    Next i
    If ok And m > 0 And sy > 0 Then
        Dim a() As Double, bb() As Double, tc() As Double, pc() As Double, w() As Double, mw() As Double
        Dim j As Long, k As Double, al As Double, am As Double, bm As Double, pA As Double, pB As Double
        Dim Z As Double, zn As Double, f As Double, df As Double, it As Long, R As Double
        Dim eta0 As Double, nu As Double, de As Double, tr As Double, xi As Double, Mv As Double
        ReDim a(1 To m): ReDim bb(1 To m): ReDim tc(1 To m): ReDim pc(1 To m): ReDim w(1 To m): ReDim mw(1 To m)
        R = 83.14462618
        For i = 1 To m
            yv(i) = yv(i) / sy: ci = idx(i)
            tc(i) = (cTc(ci) + 459.67) / 1.8: pc(i) = cPc(ci) / 14.5038: w(i) = cOm(ci): mw(i) = cMW(ci)
            Mv = Mv + yv(i) * mw(i)
            k = 0.37464 + 1.54226 * w(i) - 0.26992 * w(i) ^ 2
            al = (1# + k * (1# - Sqr(T / tc(i)))) ^ 2
            a(i) = 0.45724 * R * R * tc(i) ^ 2 / pc(i) * al
            bb(i) = 0.0778 * R * tc(i) / pc(i)
        Next i
        For i = 1 To m
            For j = 1 To m
                am = am + yv(i) * yv(j) * Sqr(a(i) * a(j))
            Next j
            bm = bm + yv(i) * bb(i)
        Next i
        pA = am * P / (R * T) ^ 2: pB = bm * P / (R * T)
        Z = 1#
        For it = 1 To 100
            f = Z ^ 3 - (1# - pB) * Z ^ 2 + (pA - 3# * pB * pB - 2# * pB) * Z - (pA * pB - pB * pB - pB ^ 3)
            df = 3# * Z * Z - 2# * (1# - pB) * Z + (pA - 3# * pB * pB - 2# * pB)
            If df = 0 Then Exit For
            zn = Z - f / df
            If Abs(zn - Z) < 0.0000000001 Then Z = zn: Exit For
            Z = zn
        Next it
        If Z <= pB Then Z = 1#
        ' dilute-gas viscosity (Lucas), Wilke-type sqrt(MW) mixing, JST density term
        Dim nn As Double, dd As Double, tcm As Double, pcm As Double, vcm As Double, rhor As Double, rhs As Double
        For i = 1 To m
            tr = T / tc(i)
            xi = 0.176 * (tc(i) / (mw(i) ^ 3 * pc(i) ^ 4)) ^ (1# / 6#)
            eta0 = (0.807 * tr ^ 0.618 - 0.357 * Exp(-0.449 * tr) + 0.34 * Exp(-4.058 * tr) + 0.018) / xi
            nn = nn + yv(i) * eta0 * Sqr(mw(i)): dd = dd + yv(i) * Sqr(mw(i))
            tcm = tcm + yv(i) * tc(i): pcm = pcm + yv(i) * pc(i)
            vcm = vcm + yv(i) * (0.291 - 0.08 * w(i)) * R * tc(i) / pc(i)
        Next i
        eta0 = nn / dd
        rhor = (P / (Z * R * T)) * vcm
        If rhor > 0.1 Then
            xi = tcm ^ (1# / 6#) / (Mv ^ 0.5 * (pcm / 1.01325) ^ (2# / 3#))
            rhs = 1.023 + 0.23364 * rhor + 0.58533 * rhor ^ 2 - 0.40758 * rhor ^ 3 + 0.093324 * rhor ^ 4
            eta0 = eta0 + (rhs ^ 4 - 1#) / xi
        End If
        If sVmu <= 0 Then eMuG = eta0 * 0.0001
        If sZ < 0.2 Then eZ = Z
        eNote = " | vapour mu/Z from component correlations (Lucas / Peng-Robinson)"
    Else
        If sVmu <= 0 Then eMuG = 0.0111 * (IIf(eMW > 0, eMW, 30) / 16#) ^ -0.3 * (T / 300#) ^ 0.8
        If sZ < 0.2 Then eZ = 1#
        eNote = " | vapour mu/Z: generic gas fallback (component constants incomplete)"
    End If
    If sGam <= 0 Then eGam = 1.1
End Sub

' Set the phase state used by DpKernel for the row about to be evaluated at pPsia
Private Sub SetState(ByVal pPsia As Double)
    Dim q As Double
    eMuG = sVmu: eMuL = sLmu: eZ = sZ: eGam = sGam: eNote = ""
    If flashOn Then
        FlashAt IIf(pPsia < 0.05, 0.05, pPsia)
        FillVapour pPsia
        eVap = eVap * fScale * vCf: eLiq = eLiq * fScale * lCf
        If eVap + eLiq > 0 Then q = eVap / (eVap + eLiq)
        If q <= 0.00001 Then
            ePhase = "LIQUID"
        ElseIf q >= 0.99999 Then
            ePhase = "VAPOR"
        Else
            ePhase = "TWO-PHASE"
        End If
        If eZ < 0.2 Then eZ = 1#
    Else
        eVap = sVap * gMult * fScale * vCf: eLiq = sLiq * gMult * fScale * lCf: eMW = sMW: ePhase = sPhase: eBeta = 0
        If eZ < 0.2 Then eZ = 1#
    End If
End Sub


'==========================================================================
Private Function LoadStream(ByVal r1 As Long, ByVal r2 As Long) As Boolean
    Dim r As Long, nm As String, has As Boolean, ws As Worksheet, i As Long, c As Long
    Dim hdr() As String, sv As Variant, lr As Long, lc As Long, rr As Long
    nm = ""
    For r = r1 To r2: If Len(GetS(r, "Stream Lookup")) > 0 Then nm = GetS(r, "Stream Lookup"): Exit For
    Next r
    sStream = nm: flashOn = False: fSrc = ""
    Dim gotSheet As Boolean
    If Len(nm) > 0 And useFlash Then gotSheet = ReadStreamSheet(nm)
    If gotSheet Then
        sStream = nm & " (sheet)"
    ElseIf Len(nm) > 0 Then
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
        sGam = SV2(sv, rr, hdr, "vap_cp_cv"): sSig = SV2(sv, rr, hdr, "liq_surf_tens")
    Else
        ' manual properties - first row of the circuit that carries them
        r = r1
        For i = r1 To r2: GetN i, "Temp (degF)", has: If has Then r = i: Exit For
        Next i
        sT = GetN(r, "Temp (degF)", has): sVap = GetN(r, "Vapor Mass Flow", has): sLiq = GetN(r, "Liq Mass Flow (lb/h)", has)
        sMW = GetN(r, "Vap MW", has): sVrho = GetN(r, "Vap Density (lb/ft3)", has): sLrho = GetN(r, "Liq Density (lb/ft3)", has)
        sVmu = GetN(r, "Vap Visc (cP)", has): sLmu = GetN(r, "Liq Visc (cP)", has)
        sZ = GetN(r, "Vap Z", has): sGam = GetN(r, "Vap Cp/Cv", has): sPst = 0: sPhase = ""
        sSig = GetN(r, "Liq Surf Tension (N/m)", has) * 1000#
        sStream = "(manual)"
    End If
    If Len(nm) > 0 Then
        For r = r1 To r2
            Dim tOv As Double: tOv = GetN(r, "Temp (degF)", has)
            If has Then sT = tOv: Exit For
        Next r
    End If
    If sZ <= 0 And Not flashOn Then sZ = 1#
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
    Dim r As Long, k As Long, has As Boolean, startP As Double, srcSet As Boolean, dstP As Double, hasDst As Boolean
    Dim p0 As Double, pEnd As Double, it As Long, err1 As Double, guess As Double, note As String
    Dim mainMap() As Long, brMap() As Long, bS() As Long, bE() As Long, nM As Long, nBr As Long, nB As Long
    ReDim mainMap(1 To r2 - r1 + 2): ReDim brMap(1 To r2 - r1 + 2): ReDim bS(1 To 40): ReDim bE(1 To 40)
    For r = r1 To r2
        If IsBranchRow(r) Then
            nBr = nBr + 1: brMap(nBr) = r
            If nB = 0 Or Len(GetS(r, "Start P (psia)")) > 0 Then
                If nB < 40 Then nB = nB + 1: bS(nB) = nBr
            End If
            bE(nB) = nBr
        Else
            nM = nM + 1: mainMap(nM) = r
        End If
    Next r
    If nM = 0 Then Exit Sub
    If Not LoadStream(r1, r2) Then Exit Sub
    startP = 0: srcSet = False: hasDst = False
    For k = 1 To nM
        r = mainMap(k)
        If IsFit(r, "source") Then
            startP = GetN(r, "Set P (psia)", has): If has Then srcSet = True
            Exit For
        End If
    Next k
    For k = nM To 1 Step -1
        r = mainMap(k)
        If IsFit(r, "destination") Then dstP = GetN(r, "Set P (psia)", hasDst): Exit For
    Next k
    Dim typed As Boolean, sp2 As Double
    If Not srcSet Then
        For k = 1 To nM
            sp2 = GetN(mainMap(k), "Start P (psia)", has): If has Then startP = sp2: typed = True: Exit For
        Next k
        If Not typed Then startP = sPst
    End If
    Dim nStart As Long: nStart = nOut
    cdNew = True: stnNo = 0
    SaveFeed 0
    nSplit = 0: nInj = 0
    If nB > 0 Then
        rowMap = mainMap: capSplit = True: useInj = False: fScale = 1: vCf = 1: lCf = 1
        pEnd = March(1, nM, cid, IIf(startP > 0, startP, sPst), False, 0)
        capSplit = False
        RunBranches cid, nB, bS, bE, brMap, IIf(startP > 0, startP, sPst), False, True, mainMap, nM, r1, r2
        RestoreFeed 0
    End If
    If Not srcSet And hasDst Then
        ' ---- back-calculate the source / pump-discharge pressure --------------
        guess = dstP + 25
        For it = 1 To 14
            ResetMain mainMap
            pEnd = March(1, nM, cid, guess, False, 0)
            err1 = dstP - pEnd
            If Abs(err1) < 0.001 Then Exit For
            guess = guess + err1: If guess < dstP Then guess = dstP
        Next it
        startP = guess
        note = "Source P back-calculated from Destination Set P = " & Format(dstP, "0.000") & " psia"
    End If
    ResetMain mainMap
    pEnd = March(1, nM, cid, startP, True, 0)
    FlushStations cid
    If flashOn Then note = note & IIf(Len(note) > 0, " | ", "") & "Flash model: " & fSrc & " (" & sStream & ")"
    If nB > 0 Then note = note & IIf(Len(note) > 0, " | ", "") & nB & " branch run(s), " & nInj & " join(s)"
    If Len(note) > 0 And nOut > nStart Then out(nStart + 1, 24) = note & " | " & out(nStart + 1, 24)
    If nB > 0 Then
        useInj = False
        RunBranches cid, nB, bS, bE, brMap, IIf(startP > 0, startP, sPst), True, False, mainMap, nM, r1, r2
        RestoreFeed 0
    End If
End Sub
Private Sub ResetMain(mainMap() As Long)
    rowMap = mainMap: useInj = True: capSplit = False
    fScale = 1: vCf = 1: lCf = 1
    RestoreFeed 0
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
    Dim rho As Double, v As Double, Re As Double, f As Double, qF As Double, qK As Double, qZ As Double
    Dim dPf As Double, dPk As Double, dPz As Double, dPt As Double, k As Double, note As String, d As Double
    Dim ctype As String, dstP As Double, hasDst As Boolean, rr As Long, lo As Double, hi As Double, mdp As Double
    Dim kp As Long, kk As Long, frac As Double, hasF As Boolean, vc As Double, hasC As Boolean, slot As Long, pendS As Double
    Dim sc0 As Double, sv0 As Double, sl0 As Double, sFd As Boolean
    Dim j As Long, pe As Double, isCV As Boolean, bdp As Double, tmp As String, stdRow As Boolean
    p = pStart
    sc0 = fScale: sv0 = vCf: sl0 = lCf: sFd = False
    If depth > 0 And nInj > 0 Then SaveFeed depth: sFd = True
    For kp = r1 To r2
        r = rowMap(kp)
        fit = GetS(r, "Fitting Name"): fl = LCase$(fit)
        If Len(fit) = 0 And Len(GetS(r, "Comp ID")) = 0 Then GoTo NextRow
        ' ---- tee split / merge, carry-over, branch join -------------------------
        pendS = 0
        frac = GetN(r, "Flow Fraction from Main", hasF)
        If hasF And IsTeeF(fl) Then
            If frac > 0 Then
                pendS = fScale: fScale = frac: pendS = 0
            Else
                pendS = fScale - IIf(fScale + frac > 0, fScale + frac, 0)
                fScale = IIf(fScale + frac > 0, fScale + frac, 0)
            End If
        End If
        vc = GetN(r, "Mass Vapor Fraction Carry Over", hasC)
        If hasC Then vCf = vCf * IIf(vc > 0, vc, 0)
        vc = GetN(r, "Mass Liquid Fraction Carry Over", hasC)
        If hasC Then lCf = lCf * IIf(vc > 0, vc, 0)
        If capSplit And pendS > 0 And TeeSplitN(fl) > 0 Then AddSplit GetS(r, "Line No"), rowSeq(r), TeeSplitN(fl), pendS * vCf, pendS * lCf
        If useInj And nInj > 0 Then
            slot = InjSlot(r)
            If slot > 0 Then DoJoin slot, p
        End If
        SetState p
        spec = GetS(r, "Piping Spec")
        bore = GetN(r, "Bore (in)", has): If has Then lastBore = bore Else bore = lastBore
        lenFt = GetN(r, "Length (ft)", has): dzFt = GetN(r, "Elev Change (ft)", has)
        fixedDp = GetN(r, "Fixed dP (psi)", hasFD): setP = GetN(r, "Set P (psia)", hasSP)
        fixK = GetN(r, "Fixed K", hasFK): instDp = GetN(r, "Instr dP (psi)", hasID)
        idIn = 0: nps = "": sch = "": basis = ""
        If bore > 0 Then idIn = ResolveID(spec, bore, nps, sch, basis)
        rho = 0: v = 0: Re = 0: f = 0: dPf = 0: dPk = 0: dPz = 0: dPt = 0: note = "": isCV = False: stdRow = False
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
                For kk = kp + 1 To r2
                    rr = rowMap(kk)
                    If LCase$(GetS(rr, "Fitting Name")) = "destination" Then dstP = GetN(rr, "Set P (psia)", hasDst): Exit For
                Next kk
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
                        pe = March(kp + 1, r2, cid, p - mdp, False, depth + 1)
                        If pe > dstP Then lo = mdp Else hi = mdp
                        If hi - lo < 0.0005 Then Exit For
                    Next j
                    bdp = (lo + hi) / 2
                End If
                tmp = "flow control: dP solved for Destination " & Format(dstP, "0.0") & " psia"
            ElseIf hasFD Then
                bdp = fixedDp: tmp = "fixed design dP"
            ElseIf hasFK And idIn > 0 Then
                DpKernel idIn * IN_TO_M, 0, 0, fixK, p, rho, v, Re, f, qF, qK, qZ
                bdp = (qF + qK + qZ) * PA_TO_PSI: tmp = "K = " & fixK
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
                DpKernel d, lenFt * FT_TO_M, dzFt * FT_TO_M, k, p, rho, v, Re, f, qF, qK, qZ
                dPf = qF * PA_TO_PSI: dPk = qK * PA_TO_PSI: dPz = qZ * PA_TO_PSI
                dPt = dPf + dPk + dPz: note = basis & eNote: stdRow = True
            End If
        End If
        pOut = p - dPt: cum = cum + lenFt
        If pOut < 0.05 Then
            note = note & "  | CHOKED (dP capped, P floor 0.05 psia)"
            dPt = IIf(p > 0.05, p - 0.05, 0): pOut = p - dPt
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
            out(nOut, 21) = dPt: out(nOut, 22) = ePhase: out(nOut, 23) = sStream: out(nOut, 24) = note
            out(nOut, 25) = IIf(eVap + eLiq > 0, eVap / (eVap + eLiq), 0)
            out(nOut, 26) = eBeta: out(nOut, 27) = eVap: out(nOut, 28) = eLiq
            If flashOn Then WriteComp cid, fit, p, pOut
            RecStn cid, r, fit, nps, idIn, cum, p, dPt, rho, v, stdRow, lenFt, dzFt
        End If
        p = pOut
NextRow:
    Next kp
    If depth > 0 Then
        fScale = sc0: vCf = sv0: lCf = sl0
        If sFd Then RestoreFeed depth
    End If
    March = p
End Function

'==========================================================================
'  Control-valve datasheet (ISA 75.01 / IEC 60534-2-1, simplified)
'==========================================================================


'==========================================================================
'  Output sheets

'==========================================================================
'  Composition_Splits (long list) and Component_Detail (components x stations)
'==========================================================================
Private Sub InitCompSheets()
    Dim ws As Worksheet
    Set ws = FreshSheet("Composition_Splits")
    ws.Range("A1").Resize(1, 11).Value = Array("Circuit", "Stn", "Fitting", "P in (psia)", "Component", "z", "K", "y (vap)", "x (liq)", "Vapour (lbmol/h)", "Liquid (lbmol/h)")
    StyleHeader ws.Range("A1").Resize(1, 11)
    Set ws = FreshSheet("Component_Detail")
    ws.Range("A1").Value = "Component mole flows (lbmol/h): Vapour | Liquid per station, one block per circuit"
    csRow = 2: cdCol = 1
End Sub

Private Sub WriteComp(ByVal cid As String, ByVal fit As String, ByVal pIn As Double, ByVal pOut As Double)
    Dim wsS As Worksheet, wsD As Worksheet, i As Long, a() As Variant, d() As Variant
    Dim fm As Double, fl As Double
    Set wsS = Worksheets("Composition_Splits"): Set wsD = Worksheets("Component_Detail")
    stnNo = stnNo + 1
    fm = eBeta * fTotMoles * gMult * fScale * vCf: fl = (1# - eBeta) * fTotMoles * gMult * fScale * lCf
    ReDim a(1 To fN, 1 To 11): ReDim d(1 To fN, 1 To 2)
    For i = 1 To fN
        a(i, 1) = cid: a(i, 2) = stnNo: a(i, 3) = fit: a(i, 4) = pIn: a(i, 5) = fNm(i)
        a(i, 6) = fZ(i): a(i, 7) = eK(i): a(i, 8) = eY(i): a(i, 9) = eX(i)
        a(i, 10) = eY(i) * fm: a(i, 11) = eX(i) * fl
        d(i, 1) = a(i, 10): d(i, 2) = a(i, 11)
    Next i
    wsS.Range(wsS.Cells(csRow, 1), wsS.Cells(csRow + fN - 1, 11)).Value = a
    csRow = csRow + fN
    If cdNew Then
        Dim nmv() As Variant: ReDim nmv(1 To fN, 1 To 1)
        For i = 1 To fN: nmv(i, 1) = fNm(i): Next i
        If cdCol > 1 Then cdCol = cdCol + 1
        wsD.Cells(3, cdCol).Value = "Component (" & cid & ")"
        wsD.Range(wsD.Cells(4, cdCol), wsD.Cells(3 + fN, cdCol)).Value = nmv
        cdCol = cdCol + 1: cdNew = False
    End If
    wsD.Cells(2, cdCol).Value = stnNo & " " & fit
    wsD.Cells(3, cdCol).Value = "Vapour": wsD.Cells(3, cdCol + 1).Value = "Liquid"
    wsD.Range(wsD.Cells(4, cdCol), wsD.Cells(3 + fN, cdCol + 1)).Value = d
    cdCol = cdCol + 2
End Sub

'==========================================================================
'  FIV (EI T2.2) and AIV (source attenuation) - ports of build_fiv_sheet / build_aiv_sheet
'==========================================================================

Private Sub InitStations()
    Dim n As Long: n = nRow + 10
    ReDim stSeq(1 To n): ReDim stFit(1 To n): ReDim stCid(1 To n): ReDim stComp(1 To n): ReDim stNps(1 To n)
    ReDim stOD(1 To n): ReDim stWall(1 To n): ReDim stID(1 To n): ReDim stRho(1 To n): ReDim stV(1 To n)
    ReDim stRg(1 To n): ReDim stRl(1 To n): ReDim stMg(1 To n): ReDim stMl(1 To n): ReDim stMu(1 To n)
    ReDim stMach(1 To n): ReDim stRv2V(1 To n): ReDim stRv2L(1 To n): ReDim stCum(1 To n): ReDim stPin(1 To n)
    ReDim stDp(1 To n): ReDim stHasMom(1 To n): ReDim stLen(1 To n): ReDim stDz(1 To n): ReDim stSig(1 To n): ReDim stMuL(1 To n)
    ReDim regOut(1 To n, 1 To 19): nReg = 0
    ReDim stKind(1 To n): ReDim stTag(1 To n): ReDim stSpec(1 To n): ReDim stT(1 To n): ReDim stGm(1 To n): ReDim stLine(1 To n)
    ReDim instOut(1 To n, 1 To 24): nInst = 0: ReDim twOut(1 To n, 1 To 21): nTw = 0
    ReDim fivOut(1 To n, 1 To 32): nFiv = 0
    ReDim aivOut(1 To n + 20, 1 To 14): nAiv = 0
    stN = 0
End Sub

' called from March on the write pass
Private Sub RecStn(ByVal cid As String, ByVal r As Long, ByVal fit As String, ByVal nps As String, ByVal idIn As Double, _
                   ByVal cum As Double, ByVal pIn As Double, ByVal dPt As Double, ByVal rho As Double, ByVal v As Double, ByVal std As Boolean, ByVal lenFt As Double, ByVal dzFt As Double)
    Dim area As Double, vm As Double, lm As Double, rg As Double, rl As Double, c As Double, k As Double
    stN = stN + 1
    stCid(stN) = cid: stFit(stN) = fit: stNps(stN) = nps: stComp(stN) = GetS(r, "Comp ID")
    stSeq(stN) = CStr(rowSeq(r))
    stOD(stN) = lastOD * 25.4: stWall(stN) = lastWall * 25.4: stID(stN) = idIn
    stRho(stN) = rho / LBFT3_TO_KGM3: stV(stN) = v / FT_TO_M
    stCum(stN) = cum: stPin(stN) = pIn: stDp(stN) = dPt
    stMg(stN) = eVap: stMl(stN) = eLiq: stMu(stN) = eMuG: stMuL(stN) = eMuL: stSig(stN) = sSig
    stLen(stN) = lenFt: stDz(stN) = dzFt
    stKind(stN) = InstKindOf(fit, GetS(r, "Instr Type")): stTag(stN) = GetS(r, "Instr Tag"): If stTag(stN) = "" Then stTag(stN) = stComp(stN)
    stSpec(stN) = GetS(r, "Piping Spec"): stT(stN) = sT: stGm(stN) = eGam: stLine(stN) = GetS(r, "Line No")
    rg = GasRho(pIn) / LBFT3_TO_KGM3: rl = sLrho
    stRg(stN) = rg: stRl(stN) = rl
    stHasMom(stN) = False
    If std And idIn > 0 Then
        area = PI / 4# * (idIn * IN_TO_M) ^ 2
        vm = eVap * LBHR_TO_KGS: lm = eLiq * LBHR_TO_KGS
        stHasMom(stN) = True
        If vm > 0 And rg > 0 Then
            c = vm / (rg * LBFT3_TO_KGM3) / area
            stRv2V(stN) = rg * LBFT3_TO_KGM3 * c * c
            k = eGam
            If k > 0 And eMW > 0 Then
                c = Sqr(k * eZ * R_UNIV * ((sT + 459.67) / 1.8) / (eMW / 1000#))
                If c > 0 Then stMach(stN) = v / c
            End If
        End If
        If lm > 0 And rl > 0 Then
            c = lm / (rl * LBFT3_TO_KGM3) / area
            stRv2L(stN) = rl * LBFT3_TO_KGM3 * c * c
        End If
    End If
End Sub

Private Function FcfF(ByVal g As Double, ByVal mu As Double) As Double
    If g < 0.2 Then FcfF = 0.2 + 4# * g: Exit Function
    If g < 0.88 Then FcfF = 1#: Exit Function
    If g < 0.99 Then FcfF = -27.882 * g ^ 2 + 45.545 * g - 17.495: Exit Function
    If mu <= 0 Then mu = 0.0012
    FcfF = Sqr(mu / 0.0012)
End Function
Private Function Ln10(ByVal x As Double) As Double
    Ln10 = Log(x) / Log(10#)
End Function
Private Function FvStiff(ByVal D As Double, ByVal T As Double) As Double
    Dim a As Double, b As Double
    If D < 812 Then
        a = 446187# + 646 * D + 0.000917 * D ^ 3: b = 0.1 * Log(D) - 1.374
    Else
        a = 419373.9 + 859.4 * D + 0.4803 * D ^ 2 - 0.0001625 * D ^ 3: b = 0.1001 * Log(D) - 0.00005309 * D - 1.358
    End If
    FvStiff = a * (D / T) ^ b
End Function
Private Function FvMedStiff(ByVal D As Double, ByVal T As Double) As Double
    Dim a As Double, b As Double
    If D < 812 Then
        a = 283921# + 370 * D: b = 0.1106 * Log(D) - 1.501
    Else
        a = 246340# + 502.2 * D + 0.1218 * D ^ 2 - 0.0000332 * D ^ 3: b = 0.0916 * Log(D) - 0.0000125 * D - 1.38
    End If
    FvMedStiff = a * (D / T) ^ b
End Function
Private Function FvMedium(ByVal D As Double, ByVal T As Double) As Double
    Dim a As Double, b As Double
    If D < 812 Then
        If D <= 219 Then
            FvMedium = Exp((13.1 - 0.00475 * D + 0.0000141 * D ^ 2) * (D / T) ^ (-0.132 + 0.000228 * D - 0.000000372 * D ^ 2))
            Exit Function
        End If
        a = 150412# + 209 * D: b = 0.0815 * Log(D) - 1.327
    Else
        a = 179.1 * D + 195015#: b = 0.1533 * Log(D) - 0.00007 * D - 1.768
    End If
    FvMedium = a * (D / T) ^ b
End Function
Private Function FvFlex(ByVal D As Double, ByVal T As Double) As Double
    Dim a As Double, b As Double
    If D < 812 Then
        If D <= 219 Then
            FvFlex = Exp((0.0000132 * D ^ 2 - 0.00442 * D + 12.22) * (D / T) ^ (0.000284 * D - 0.000000462 * D ^ 2 - 0.164))
            Exit Function
        End If
        a = 41.21 * D + 49397#: b = 0.0815 * Log(D) - 1.384
    Else
        a = 42.21 * D + 52778#: b = 0.0697 * Log(D) - 1.323
    End If
    FvFlex = a * (D / T) ^ b
End Function
Private Function SpanOf(ByVal arr As Long, ByVal D As Double) As Double
    Select Case arr
        Case 2   ' Medium
            If D <= 762 Then SpanOf = -0.00001597 * D ^ 2 + 0.0336 * D + 4.429 Else SpanOf = 0.7443 * D ^ 0.5033
        Case 3   ' Medium stiff
            If D <= 762 Then SpanOf = -0.00001189 * D ^ 2 + 0.0253 * D + 3.36 Else SpanOf = 0.5623 * D ^ 0.5034
        Case 4   ' Stiff
            If D <= 762 Then SpanOf = -0.00001235 * D ^ 2 + 0.02 * D + 2.056 Else SpanOf = 0.4056 * D ^ 0.4927
    End Select
End Function
Private Function FvOf(ByVal arr As Long, ByVal D As Double, ByVal T As Double) As Double
    Select Case arr
        Case 1: FvOf = FvFlex(D, T)
        Case 2: FvOf = FvMedium(D, T)
        Case 3: FvOf = FvMedStiff(D, T)
        Case Else: FvOf = FvStiff(D, T)
    End Select
End Function

' assess_span: returns text in sup, span (m) in sp (-1 = no limit, -2 = not achievable)
Private Sub AssessSpan(ByVal rv2 As Double, ByVal fcfV As Double, ByVal D As Double, ByVal T As Double, ByVal lim As Double, _
                       ByRef sup As String, ByRef sp As Double)
    Dim fvReq As Double, a As Long
    fvReq = rv2 * fcfV / lim
    For a = 1 To 4
        If FvOf(a, D, T) >= fvReq Then
            sup = Choose(a, "Flexible", "Medium", "Medium Stiff", "Stiff")
            If a = 1 Then sp = -1 Else sp = SpanOf(a, D)
            Exit Sub
        End If
    Next a
    sup = "Not achievable": sp = -2
End Sub

Private Sub FlushStations(ByVal cid As String)
    Dim i As Long, j As Long, D As Double, T As Double, rv2 As Double, qg As Double, ql As Double, g As Double, fc As Double
    Dim lims As Variant, sup As String, sp As Double, c As Long, worst As Long, wr2 As Double
    Dim srcI As Long, dmax As Double, mdot As Double, rhoSrc As Double, dpPa As Double, wac As Double, lws As Double
    Dim cumAtt As Double, prevCum As Double, started As Boolean, dist As Double, lwl As Double, cmp As Double, eil As Double, dyn As Double
    Dim cmS As String, eiS As String, k As Long
    If stN = 0 Then Exit Sub
    lims = Array(0.15, 0.25, 0.35, 0.45)
    ' ---- FIV
    For i = 1 To stN
        D = stOD(i): T = stWall(i)
        If D <= 0 Or T <= 0 Then GoTo NextF
        rv2 = stRho(i) * LBFT3_TO_KGM3 * (stV(i) * FT_TO_M) ^ 2
        qg = 0: ql = 0
        If stRg(i) > 0 Then qg = stMg(i) * LBHR_TO_KGS / (stRg(i) * LBFT3_TO_KGM3)
        If stRl(i) > 0 Then ql = stMl(i) * LBHR_TO_KGS / (stRl(i) * LBFT3_TO_KGM3)
        If qg + ql > 0 Then g = qg / (qg + ql) Else g = 1#
        fc = FcfF(g, stMu(i) * 0.001)
        nFiv = nFiv + 1
        fivOut(nFiv, 1) = cid: fivOut(nFiv, 2) = stSeq(i): fivOut(nFiv, 3) = stComp(i): fivOut(nFiv, 4) = stFit(i)
        fivOut(nFiv, 5) = stNps(i): fivOut(nFiv, 6) = D: fivOut(nFiv, 7) = T: fivOut(nFiv, 8) = stID(i)
        fivOut(nFiv, 9) = stV(i): fivOut(nFiv, 10) = stRho(i): fivOut(nFiv, 11) = rv2: fivOut(nFiv, 12) = g: fivOut(nFiv, 13) = fc
        fivOut(nFiv, 14) = FvStiff(D, T): fivOut(nFiv, 15) = FvMedStiff(D, T): fivOut(nFiv, 16) = FvMedium(D, T): fivOut(nFiv, 17) = FvFlex(D, T)
        If stHasMom(i) Then
            fivOut(nFiv, 18) = stMach(i): fivOut(nFiv, 19) = stRv2V(i): fivOut(nFiv, 20) = stRv2L(i)
        End If
        c = 21
        For j = 0 To 3
            AssessSpan rv2, fc, D, T, CDbl(lims(j)), sup, sp
            fivOut(nFiv, c) = sup
            If sp = -1 Then
                fivOut(nFiv, c + 1) = "No EI span limit"
            ElseIf sp = -2 Then
                fivOut(nFiv, c + 1) = "Not achievable"
            Else
                fivOut(nFiv, c + 1) = sp: fivOut(nFiv, c + 2) = sp / FT_TO_M
            End If
            c = c + 3
        Next j
NextF:
    Next i
    ' ---- AIV: source = largest total dP
    dmax = -1E+30
    For i = 1 To stN
        If stDp(i) > dmax Then dmax = stDp(i): srcI = i
    Next i
    mdot = (sVap + sLiq) * gMult * LBHR_TO_KGS
    rhoSrc = stRho(srcI) * LBFT3_TO_KGM3
    If rhoSrc <= 0 Then            ' CV / orifice rows carry no density: use the nearest upstream flowing station
        For k = srcI - 1 To 1 Step -1
            If stRho(k) > 0 Then rhoSrc = stRho(k) * LBFT3_TO_KGM3: Exit For
        Next k
    End If
    dpPa = stDp(srcI) * PSIA_TO_PA
    If rhoSrc > 0 Then wac = 0.0001 * mdot * dpPa / rhoSrc
    If wac > 0 Then lws = 10# * Ln10(wac / 1E-12)
    nAiv = nAiv + 1
    aivOut(nAiv, 1) = cid: aivOut(nAiv, 2) = "SOURCE": aivOut(nAiv, 3) = stComp(srcI) & " / " & stFit(srcI)
    aivOut(nAiv, 4) = "dP (psi)": aivOut(nAiv, 5) = stDp(srcI): aivOut(nAiv, 6) = "W_ac (W)": aivOut(nAiv, 7) = wac
    aivOut(nAiv, 8) = "PWL (dB)": aivOut(nAiv, 9) = lws
    For i = srcI To stN
        D = stOD(i): T = stWall(i)
        If i = srcI Then prevCum = stCum(i): cumAtt = 0
        cumAtt = cumAtt + 0.2 * IIf(stCum(i) - prevCum > 0, stCum(i) - prevCum, 0) * FT_TO_M + FitLoss(stFit(i))
        prevCum = stCum(i)
        dist = IIf(stCum(i) - stCum(srcI) > 0, stCum(i) - stCum(srcI), 0) * FT_TO_M
        lwl = lws - cumAtt
        If D > 0 Then cmp = lwl + 10# * Ln10(D) Else cmp = lwl
        If D > 0 And T > 0 Then eil = lwl + 20# * Ln10(D / T) Else eil = lwl
        If cmp < 145 Then
            cmS = "LOW"
        ElseIf cmp < 155 Then
            cmS = "MEDIUM"
        Else
            cmS = "HIGH"
        End If
        If eil < 160 Then
            eiS = "LOW"
        ElseIf eil < 170 Then
            eiS = "MEDIUM"
        Else
            eiS = "HIGH"
        End If
        If T > 0 Then dyn = (stPin(i) * PSIA_TO_PA * (D / 1000#) / (2# * (T / 1000#))) / 1000000# Else dyn = 0
        nAiv = nAiv + 1
        aivOut(nAiv, 1) = cid: aivOut(nAiv, 2) = stSeq(i): aivOut(nAiv, 3) = stComp(i) & " / " & stFit(i)
        aivOut(nAiv, 4) = D: aivOut(nAiv, 5) = T: aivOut(nAiv, 6) = dist: aivOut(nAiv, 7) = cumAtt
        aivOut(nAiv, 8) = lwl: aivOut(nAiv, 9) = cmp: aivOut(nAiv, 10) = cmS: aivOut(nAiv, 11) = eil
        aivOut(nAiv, 12) = eiS: aivOut(nAiv, 13) = dyn
    Next i
    RegimeRows cid
    InstrRows cid
    stN = 0
End Sub
Private Function FitLoss(ByVal fit As String) As Double
    Dim f As String: f = LCase$(fit)
    If InStr(f, "elbow") > 0 Or InStr(f, "bend") > 0 Then
        FitLoss = 2#
    ElseIf InStr(f, "tee") > 0 Or InStr(f, "stub") > 0 Or InStr(f, "branch") > 0 Then
        FitLoss = 5#
    ElseIf InStr(f, "reducer") > 0 Or InStr(f, "swage") > 0 Then
        FitLoss = 2#
    End If
End Function

Private Sub WriteFivAiv()
    Dim ws As Worksheet, h As Variant, i As Long, c As Long, c2 As Variant
    h = Array("Circuit", "Seq", "Comp ID", "Fitting", "NPS", "OD (mm)", "Wall (mm)", "ID (in)", "Velocity (ft/s)", "Rho (lb/ft3)", _
        "rho v2 (kg/m s2)", "GVF", "FCF", "Fv Stiff", "Fv Med-Stiff", "Fv Medium", "Fv Flexible", "Mach (vapor)", "rho v2 Vapor", "rho v2 Liquid", _
        "Support (LOF<=0.15)", "Span @0.15 (m)", "Span @0.15 (ft)", "Support (LOF<=0.25)", "Span @0.25 (m)", "Span @0.25 (ft)", _
        "Support (LOF<=0.35)", "Span @0.35 (m)", "Span @0.35 (ft)", "Support (LOF<=0.45)", "Span @0.45 (m)", "Span @0.45 (ft)")
    Set ws = FreshSheet("FIV_EI_T2.2")
    ws.Range("A1").Resize(1, 32).Value = h: StyleHeader ws.Range("A1").Resize(1, 32)
    If nFiv > 0 Then
        ws.Range("A2").Resize(nFiv, 32).Value = fivOut
        ws.Range("F2:S" & nFiv + 1).NumberFormat = "0.000"
        For i = 2 To nFiv + 1
            For c = 21 To 30 Step 3
                Select Case CStr(ws.Cells(i, c).Value)
                    Case "Stiff", "Not achievable": ws.Cells(i, c).Interior.Color = RGB(255, 199, 206)
                    Case "Medium Stiff": ws.Cells(i, c).Interior.Color = RGB(255, 235, 156)
                    Case Else: ws.Cells(i, c).Interior.Color = RGB(198, 239, 206)
                End Select
            Next c
        Next i
    End If
    ws.Columns("A:AF").ColumnWidth = 12: ws.Columns("D").ColumnWidth = 22
    ws.Range("A" & nFiv + 4).Value = "Energy Institute T2.2 screening (SI): LOF = rho v2 * FCF / Fv; least-stiff support meeting each LOF limit. VERIFY against the contracted EI edition."
    h = Array("Circuit", "Seq", "Comp / Fitting", "OD (mm)", "Wall (mm)", "Dist from src (m)", "Cum. atten (dB)", "Local PWL (dB)", _
        "C-M param", "C-M screen", "EI likelihood", "EI risk", "Dyn. stress idx (MPa)", "")
    Set ws = FreshSheet("AIV")
    ws.Range("A1").Resize(1, 13).Value = h: StyleHeader ws.Range("A1").Resize(1, 13)
    If nAiv > 0 Then
        ws.Range("A2").Resize(nAiv, 13).Value = aivOut
        ws.Range("D2:I" & nAiv + 1).NumberFormat = "0.00": ws.Range("K2:K" & nAiv + 1).NumberFormat = "0.0": ws.Range("M2:M" & nAiv + 1).NumberFormat = "0.0"
        For i = 2 To nAiv + 1
            If CStr(ws.Cells(i, 2).Value) = "SOURCE" Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, 13)).Interior.Color = RGB(160, 238, 192)
            Else
                For Each c2 In Array(10, 12)
                    Select Case CStr(ws.Cells(i, c2).Value)
                        Case "LOW": ws.Cells(i, c2).Interior.Color = RGB(198, 239, 206)
                        Case "MEDIUM": ws.Cells(i, c2).Interior.Color = RGB(255, 235, 156)
                        Case "HIGH": ws.Cells(i, c2).Interior.Color = RGB(255, 199, 206)
                    End Select
                Next c2
            End If
        Next i
    End If
    ws.Columns("A:M").ColumnWidth = 13: ws.Columns("C").ColumnWidth = 26
    ws.Range("A" & nAiv + 4).Value = "AIV: source = largest dP; W = 1e-4*mdot*dP/rho; PWL = 10log10(W/1e-12); 0.2 dB/m + 2 dB elbow / 5 tee / 2 reducer; Carucci-Mueller and EI likelihood screens. Valve rows use the nearest upstream density. VERIFY eta and source."
End Sub


'==========================================================================
'  Two-phase flow-regime maps (ports of the six correlations in hydraulics.py)
'==========================================================================

Private Sub TdGeom(ByVal hd As Double, ByRef al As Double, ByRef ag As Double, ByRef sl As Double, ByRef sg As Double, ByRef si As Double)
    Dim g As Double, ac As Double, rt As Double
    g = 2# * hd - 1#
    If g > 0.999999 Then g = 0.999999
    If g < -0.999999 Then g = -0.999999
    ac = ACosF(g): rt = Sqr(1# - g * g)
    al = 0.25 * (PI - ac + g * rt): ag = 0.25 * (ac - g * rt)
    sl = PI - ac: sg = ac: si = rt
End Sub
Private Function ACosF(ByVal x As Double) As Double
    ACosF = Atn(-x / Sqr(-x * x + 1#)) + 2# * Atn(1#)
End Function
Private Function TdF(ByVal hd As Double, ByVal x2 As Double, ByRef ok As Boolean) As Double
    Dim al As Double, ag As Double, sl As Double, sg As Double, si As Double, A As Double
    Dim ul As Double, ug As Double, dl As Double, dg As Double, liq As Double, gas As Double
    TdGeom hd, al, ag, sl, sg, si
    ok = False
    If al < 0.000000001 Or ag < 0.000000001 Then Exit Function
    A = PI / 4#: ul = A / al: ug = A / ag: dl = 4# * al / sl: dg = 4# * ag / (sg + si)
    liq = x2 * (ul * dl) ^ (-0.2) * ul ^ 2 * (sl / al)
    gas = (ug * dg) ^ (-0.2) * ug ^ 2 * (sg / ag + si / al + si / ag)
    TdF = liq - gas: ok = True
End Function
Private Function TdEquil(ByVal x2 As Double) As Double
    Dim lo As Double, hi As Double, flo As Double, fhi As Double, fm As Double, mdp As Double
    Dim o1 As Boolean, o2 As Boolean, o3 As Boolean, i As Long, best As Double, ba As Double, v As Double
    lo = 0.0001: hi = 0.9999
    flo = TdF(lo, x2, o1): fhi = TdF(hi, x2, o2)
    If (Not o1) Or (Not o2) Or flo * fhi > 0 Then
        best = 0.5: ba = 1E+30
        For i = 1 To 999
            v = TdF(i / 1000#, x2, o3)
            If o3 Then
                If Abs(v) < ba Then ba = Abs(v): best = i / 1000#
            End If
        Next i
        TdEquil = best: Exit Function
    End If
    For i = 1 To 80
        mdp = 0.5 * (lo + hi)
        fm = TdF(mdp, x2, o3)
        If Not o3 Then
            lo = mdp
        ElseIf flo * fm <= 0 Then
            hi = mdp: fhi = fm
        Else
            lo = mdp: flo = fm
        End If
    Next i
    TdEquil = 0.5 * (lo + hi)
End Function
Private Function DpdlF(ByRef t As TPI, ByVal v As Double, ByVal rho As Double, ByVal mu As Double) As Double
    Dim Re As Double, ff As Double
    If v <= 0 Then DpdlF = 1E-12: Exit Function
    Re = rho * Abs(v) * t.d / mu
    If Re > 0 Then ff = 0.046 * Re ^ -0.2 Else ff = 0.046
    DpdlF = 4# * ff * rho * v * v / (2# * t.d)
End Function
Private Function TaitelDukler(ByRef t As TPI) As String
    Dim cs As Double, dpl As Double, dpg As Double, X As Double, F As Double, nul As Double, resl As Double, K As Double, tg As Double
    Dim hd As Double, al As Double, ag As Double, sl As Double, sg As Double, si As Double, A As Double, ug As Double, ul As Double
    Dim f2c As Double, kc As Double, tc As Double, dr As Double
    cs = Cos(t.th * PI / 180#): If cs < 0.05 Then cs = 0.05
    dpl = DpdlF(t, t.vsl, t.rl, t.mul): dpg = DpdlF(t, t.vsg, t.rg, t.mug)
    If dpg > 0 Then X = Sqr(dpl / dpg) Else X = 1000000#
    dr = t.rl - t.rg: If dr < 0.000000001 Then dr = 0.000000001
    F = Sqr(t.rg / dr) * t.vsg / Sqr(t.d * G_SI * cs)
    nul = t.mul / t.rl
    If nul > 0 Then resl = t.vsl * t.d / nul
    If resl > 0 Then K = F * Sqr(resl)
    tg = Sqr(Abs(dpl) / (dr * G_SI * cs))
    hd = TdEquil(X * X)
    TdGeom hd, al, ag, sl, sg, si
    A = PI / 4#: ug = A / ag: ul = A / al
    f2c = (1# - hd) ^ 2 * ag / (ug ^ 2 * (si / A))
    If F * F <= f2c Then
        If t.vsg > 0 Then
            kc = Sqr(4# * nul * dr * G_SI * cs / (0.01 * t.rg * IIf(t.vsg > 0.000000001, t.vsg, 0.000000001) ^ 2))
        Else
            kc = 1000000000#
        End If
        If K >= kc Then TaitelDukler = "Stratified Wavy" Else TaitelDukler = "Stratified Smooth"
    ElseIf hd >= 0.35 Then
        tc = Sqr(8# * ag / (si * ul ^ 2 * (ul * 4# * al / sl) ^ -0.2))
        If tg >= tc Then TaitelDukler = "Dispersed Bubble" Else TaitelDukler = "Intermittent"
    Else
        TaitelDukler = "Annular"
    End If
End Function
Private Function TaitelBarnea(ByRef t As TPI) As String
    Dim dr As Double, u0 As Double, va As Double, vdb As Double, vbs As Double
    dr = t.rl - t.rg: If dr < 0.000000001 Then dr = 0.000000001
    u0 = 1.53 * (G_SI * t.sig * dr / t.rl ^ 2) ^ 0.25
    va = 3.1 * (G_SI * t.sig * dr / t.rg ^ 2) ^ 0.25
    If t.vsg >= va Then TaitelBarnea = "Annular": Exit Function
    If t.d > 0 Then vdb = 4# * (t.d * G_SI * dr / t.rl) ^ 0.5 * (t.sig / (dr * G_SI * t.d ^ 2)) ^ 0.1 Else vdb = 1000000000#
    If t.vm >= vdb And t.gvf < 0.52 Then TaitelBarnea = "Dispersed Bubble": Exit Function
    vbs = 3# * t.vsg - 1.15 * u0
    If t.vsl > vbs And t.gvf < 0.25 Then TaitelBarnea = "Bubble": Exit Function
    If t.vsg > 0.5 And t.gvf > 0.35 Then TaitelBarnea = "Churn": Exit Function
    TaitelBarnea = "Slug"
End Function
Private Function BakerF(ByRef t As TPI) As String
    Dim gg As Double, gl As Double, lam As Double, psi As Double, by As Double, bx As Double
    gg = t.rg * t.vsg: gl = t.rl * t.vsl
    lam = Sqr((t.rg / 1.225) * (t.rl / 1000#))
    psi = (0.072 / t.sig) * ((t.mul / 0.001) * (1000# / t.rl) ^ 2) ^ (1# / 3#)
    If lam > 0 Then by = gg / lam
    If gg > 0 Then bx = gl / gg * lam * psi Else bx = 1000000000#
    If by > 80 Then
        If bx < 1 Then BakerF = "Dispersed" Else BakerF = "Annular"
    ElseIf by > 3 Then
        If bx < 0.3 Then
            BakerF = "Annular"
        ElseIf bx < 6 Then
            BakerF = "Slug"
        Else
            BakerF = "Bubble"
        End If
    Else
        If bx < 0.5 Then
            BakerF = "Stratified"
        ElseIf bx < 5 Then
            BakerF = "Wave"
        ElseIf bx < 50 Then
            BakerF = "Plug"
        Else
            BakerF = "Bubble"
        End If
    End If
End Function
Private Function HewittRoberts(ByRef t As TPI) As String
    Dim jl As Double, jg As Double
    jl = t.rl * t.vsl ^ 2: jg = t.rg * t.vsg ^ 2
    If jg > 1000 And jl > 10000 Then
        HewittRoberts = "Wispy Annular"
    ElseIf jg > 1000 Then
        HewittRoberts = "Annular"
    ElseIf jg > 10 Then
        If jl < 1000 Then HewittRoberts = "Churn" Else HewittRoberts = "Bubbly-Slug"
    ElseIf jl > 1000 Then
        HewittRoberts = "Bubbly"
    Else
        HewittRoberts = "Slug"
    End If
End Function
Private Function BarneaUnified(ByRef t As TPI) As String
    Dim ang As Double, dr As Double, ff As Double, vdb As Double, re As Double
    ang = Abs(t.th)
    dr = t.rl - t.rg: If dr < 0.000000001 Then dr = 0.000000001
    re = t.rl * t.vm * t.d / t.mul: If re < 1 Then re = 1
    ff = 0.046 * re ^ -0.2
    If t.d > 0 Then
        vdb = 2# * (0.4 * t.sig / (dr * G_SI)) ^ 0.5 * (t.rl / t.sig) ^ 0.6 * (2# * ff / t.d) ^ 0.4
    Else
        vdb = 1000000000#
    End If
    If t.vm >= vdb And t.gvf < 0.52 Then BarneaUnified = "Dispersed Bubble": Exit Function
    If ang >= 60 Then BarneaUnified = TaitelBarnea(t) Else BarneaUnified = TaitelDukler(t)
End Function
Private Function ShellBB(ByRef t As TPI) As String
    Dim lam As Double, fr As Double, l1 As Double, l2 As Double, l3 As Double, l4 As Double
    If t.vm > 0 Then lam = t.vsl / t.vm
    If lam > 1 Then lam = 1
    If lam < 0.000001 Then lam = 0.000001
    If t.d > 0 Then fr = t.vm ^ 2 / (G_SI * t.d)
    l1 = 316# * lam ^ 0.302: l2 = 0.0009252 * lam ^ -2.4684: l3 = 0.1 * lam ^ -1.4516: l4 = 0.5 * lam ^ -6.738
    If (lam < 0.01 And fr < l1) Or (lam >= 0.01 And fr < l2) Then
        ShellBB = "Stratified"
    ElseIf lam >= 0.01 And fr >= l2 And fr <= l3 Then
        ShellBB = "Transition"
    ElseIf (lam >= 0.01 And lam < 0.4 And fr > l3 And fr <= l1) Or (lam >= 0.4 And fr > l3 And fr <= l4) Then
        ShellBB = "Intermittent"
    Else
        ShellBB = "Distributed"
    End If
End Function

Private Sub RegimeRows(ByVal cid As String)
    Dim i As Long, t As TPI, qg As Double, ql As Double, area As Double, sp1 As String, ln As Double
    Dim fz As Double, hls As Double, rs As Double, force As Double, inner As Double
    For i = 1 To stN
        If stID(i) <= 0 Then GoTo NextR
        nReg = nReg + 1
        t.d = stID(i) * IN_TO_M: area = PI * 0.25 * t.d * t.d
        t.rg = stRg(i) * LBFT3_TO_KGM3: t.rl = stRl(i) * LBFT3_TO_KGM3
        t.mug = IIf(stMu(i) > 0, stMu(i), 0.01) * 0.001: t.mul = IIf(stMuL(i) > 0, stMuL(i), 0.5) * 0.001
        t.sig = IIf(stSig(i) > 0, stSig(i), 20#) * 0.001
        qg = 0: ql = 0
        If t.rg > 0 Then qg = stMg(i) * LBHR_TO_KGS / t.rg
        If t.rl > 0 Then ql = stMl(i) * LBHR_TO_KGS / t.rl
        t.vsg = qg / area: t.vsl = ql / area: t.vm = t.vsg + t.vsl
        If qg + ql > 0 Then t.gvf = qg / (qg + ql) Else t.gvf = 1#
        ln = stLen(i)
        If ln > 0.000000001 Then
            fz = stDz(i) / ln
            If fz >= 1 Then
                t.th = 90
            ElseIf fz <= -1 Then
                t.th = -90
            Else
                t.th = Atn(fz / Sqr(1# - fz * fz)) * 180# / PI
            End If
        Else
            t.th = 0
        End If
        regOut(nReg, 1) = cid: regOut(nReg, 2) = stSeq(i): regOut(nReg, 3) = stComp(i): regOut(nReg, 4) = stFit(i)
        regOut(nReg, 5) = stID(i): regOut(nReg, 6) = t.th: regOut(nReg, 7) = t.vsg: regOut(nReg, 8) = t.vsl
        regOut(nReg, 9) = t.vm: regOut(nReg, 10) = t.gvf
        sp1 = ""
        If t.vsl <= 0.000000001 Or t.gvf >= 0.99999 Then
            sp1 = "Single-phase gas"
        ElseIf t.vsg <= 0.000000001 Or t.gvf <= 0.00001 Then
            sp1 = "Single-phase liquid"
        End If
        If Len(sp1) > 0 Then
            regOut(nReg, 11) = sp1
        Else
            regOut(nReg, 11) = TaitelDukler(t): regOut(nReg, 12) = TaitelBarnea(t): regOut(nReg, 13) = BakerF(t)
            regOut(nReg, 14) = HewittRoberts(t): regOut(nReg, 15) = BarneaUnified(t): regOut(nReg, 16) = ShellBB(t)
            If t.d > 0 And t.vm > 0 Then
                inner = t.vsl / (G_SI * t.d) * (19.75 / t.vm + t.vm)
                If inner > 0 Then fz = 0.0226 * inner ^ 1.2 Else fz = 0
            Else
                fz = 0
            End If
            hls = 1# / (1# + (t.vm / 8.66) ^ 1.39)
            rs = t.rl * hls + t.rg * (1# - hls)
            force = Sqr(2#) * rs * t.vm * t.vm * area
            regOut(nReg, 17) = fz * 60#: regOut(nReg, 18) = hls: regOut(nReg, 19) = force
        End If
NextR:
    Next i
End Sub

Private Sub WriteRegime()
    Dim ws As Worksheet, h As Variant
    h = Array("Circuit", "Seq", "Comp ID", "Fitting", "ID (in)", "Incl (deg)", "Vsg (m/s)", "Vsl (m/s)", "Vm (m/s)", "GVF", _
        "Taitel-Dukler (horiz)", "Taitel-Barnea (vert)", "Baker (horiz)", "Hewitt-Roberts (vert)", "Barnea unified", "Shell DEP / Beggs-Brill", _
        "Slug freq (/min)", "Slug holdup", "Slug force @bend (N)")
    Set ws = FreshSheet("Two_Phase_Regime")
    ws.Range("A1").Resize(1, 19).Value = h: StyleHeader ws.Range("A1").Resize(1, 19)
    If nReg > 0 Then
        ws.Range("A2").Resize(nReg, 19).Value = regOut
        ws.Range("E2:J" & nReg + 1).NumberFormat = "0.000"
    End If
    ws.Columns("A:S").ColumnWidth = 14: ws.Columns("D").ColumnWidth = 22: ws.Columns("K:P").ColumnWidth = 20
    ws.Range("A" & nReg + 3).Value = "Six flow-pattern correlations per station (Taitel-Dukler, Taitel-Barnea, Baker, Hewitt-Roberts, Barnea unified, Shell/Beggs-Brill) + Gregory slug frequency / holdup / bend force. VERIFY against a multiphase simulator before design use."
End Sub


'==========================================================================
'  Instrument_Data, control-valve travel, inline instruments and thermowell wake check
'==========================================================================

Private Sub ReadInstData()
    Dim ws As Worksheet, lr As Long
    nID = 0: nTC = 0
    On Error Resume Next
    Set ws = Worksheets("Instrument_Data")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lr >= 3 Then idat = ws.Range("A3:O" & lr).Value: nID = lr - 2
    End If
    Set ws = Nothing
    On Error Resume Next
    Set ws = Worksheets("Data_TWClass")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lr >= 2 Then tcls = ws.Range("A2:F" & lr).Value: nTC = lr - 1
    End If
End Sub
Private Function IDRow(ByVal tag As String) As Long
    Dim i As Long
    tag = UCase$(Trim$(tag))
    If Len(tag) = 0 Then Exit Function
    For i = 1 To nID
        If UCase$(Trim$(CStr(idat(i, 1)))) = tag Then IDRow = i: Exit Function
    Next i
End Function
Private Function IDNum(ByVal row As Long, ByVal col As Long) As Double
    Dim ok As Boolean
    If row > 0 Then IDNum = NumOf(idat(row, col), ok)
End Function
Private Function IDHas(ByVal row As Long, ByVal col As Long) As Boolean
    Dim ok As Boolean, x As Double
    If row > 0 Then x = NumOf(idat(row, col), ok): IDHas = ok
End Function

Private Function ParseCurve(ByVal txt As String, xs() As Double, cs() As Double) As Long
    Dim parts() As String, i As Long, j As Long, p As Long, n As Long, a As Double, b As Double, tx As Double, tc As Double
    Dim tok As String
    ReDim xs(1 To 60): ReDim cs(1 To 60)
    txt = Replace(txt, vbLf, ";")
    parts = Split(txt, ";")
    For i = 0 To UBound(parts)
        tok = Trim$(parts(i))
        p = InStr(tok, ":"): If p = 0 Then p = InStr(tok, "=")
        If p > 1 Then
            If IsNumeric(Trim$(Left$(tok, p - 1))) And IsNumeric(Trim$(Mid$(tok, p + 1))) And n < 60 Then
                n = n + 1: xs(n) = Val(Trim$(Left$(tok, p - 1))): cs(n) = Val(Trim$(Mid$(tok, p + 1)))
            End If
        End If
    Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If xs(j) < xs(i) Then
                tx = xs(i): xs(i) = xs(j): xs(j) = tx: tc = cs(i): cs(i) = cs(j): cs(j) = tc
            End If
        Next j
    Next i
    If n >= 2 Then ParseCurve = n
End Function

' valve travel (%) for the required Cv; returns -1 when no rated Cv / curve
Private Function CvTravel(ByVal cvReq As Double, ByVal rated As Double, ByVal ch As String, ByVal rng As Double, _
                          ByVal curve As String, ByRef basis As String, ByRef status As String, ByRef ratedOut As Double) As Double
    Dim xs() As Double, cs() As Double, n As Long, i As Long, x As Double, ratio As Double
    CvTravel = -1: basis = "": status = "": ratedOut = rated
    n = ParseCurve(curve, xs, cs)
    If n >= 2 Then
        If rated <= 0 Then ratedOut = cs(n)
        If cvReq <= cs(1) Then
            If cs(1) > 0 Then x = xs(1) * cvReq / cs(1) Else x = xs(1)
        ElseIf cvReq >= cs(n) Then
            x = xs(n) + (cvReq - cs(n)) / IIf(cs(n) - cs(n - 1) > 0.000000001, cs(n) - cs(n - 1), 0.000000001) * (xs(n) - xs(n - 1))
        Else
            x = xs(n)
            For i = 1 To n - 1
                If cs(i) <= cvReq And cvReq <= cs(i + 1) Then
                    If cs(i + 1) > cs(i) Then x = xs(i) + (cvReq - cs(i)) / (cs(i + 1) - cs(i)) * (xs(i + 1) - xs(i)) Else x = xs(i)
                    Exit For
                End If
            Next i
        End If
        basis = "measured curve (" & n & " points)"
    ElseIf rated > 0 Then
        If Len(ch) = 0 Then ch = "equal %"
        If rng <= 0 Then rng = 50
        ratio = cvReq / rated
        If ratio <= 0 Then
            x = 0
        ElseIf LCase$(Left$(ch, 3)) = "lin" Then
            x = ratio * 100#
        ElseIf LCase$(Left$(ch, 5)) = "quick" Then
            x = ratio * ratio * 100#
        Else
            If rng > 1 Then x = (1# + Log(ratio) / Log(rng)) * 100# Else x = ratio * 100#
        End If
        basis = ch & IIf(LCase$(Left$(ch, 3)) <> "lin" And LCase$(Left$(ch, 5)) <> "quick", ", R = " & rng, "")
    Else
        Exit Function
    End If
    CvTravel = x
    If x > 100 Then
        status = "UNDERSIZED - needs > 100 % travel"
    ElseIf x > 90 Then
        status = "Too open (> 90 %)"
    ElseIf x < 10 Then
        status = "Too closed (< 10 %)"
    Else
        status = "OK (10-90 %)"
    End If
End Function

' bubble-point pressure at the line temperature for a liquid-only inlet (0 if none)
Private Function BubbleP(ByVal p1 As Double) As Double
    Dim lo As Double, hi As Double, i As Long, mdp As Double
    If Not flashOn Then Exit Function
    FlashAt 0.05
    If eVap <= 0 Then FlashAt IIf(p1 < 0.05, 0.05, p1): Exit Function
    lo = 0.05: hi = IIf(p1 > 0.05, p1, 0.05)
    For i = 1 To 24
        mdp = 0.5 * (lo + hi)
        FlashAt mdp
        If eVap > 0 Then lo = mdp Else hi = mdp
    Next i
    BubbleP = 0.5 * (lo + hi)
    FlashAt IIf(p1 < 0.05, 0.05, p1)
End Function

Private Sub RecordCV(ByVal r As Long, ByVal cid As String, ByVal p1 As Double, ByVal dp As Double, ByVal basis As String)
    Dim wl As Double, wv As Double, q As Double, sg As Double, rho1 As Double, cvg As Double, cvl As Double
    Dim x As Double, fk As Double, y As Double, FL As Double, xT As Double, p2 As Double, tag As String, xmax As Double, xe As Double
    Dim dpe As Double, pv As Double, ff As Double, dpmax As Double, ir As Long, notes As String, chk As String
    Dim trv As Double, tb As String, ts As String, rated As Double, rOut As Double
    p2 = p1 - dp: tag = GetS(r, "Instr Tag"): If tag = "" Then tag = GetS(r, "Comp ID")
    ir = IDRow(tag)
    FL = 0.9: xT = 0.7
    If IDHas(ir, 10) Then FL = IDNum(ir, 10)
    If IDHas(ir, 11) Then xT = IDNum(ir, 11)
    wl = eLiq: wv = eVap
    nCV = nCV + 1
    cvOut(nCV, 1) = cid: cvOut(nCV, 2) = tag: cvOut(nCV, 3) = sStream: cvOut(nCV, 4) = ePhase
    cvOut(nCV, 5) = p1: cvOut(nCV, 6) = p2: cvOut(nCV, 7) = dp: cvOut(nCV, 8) = wl + wv: cvOut(nCV, 9) = sT
    cvOut(nCV, 15) = FL: cvOut(nCV, 16) = xT
    If dp <= 0 Then cvOut(nCV, 13) = "dP = 0: valve fully open / no dP basis": Exit Sub
    If wv > 0 And p1 > 0 Then
        rho1 = GasRho(p1) / LBFT3_TO_KGM3
        fk = IIf(eGam > 0, eGam, 1.3) / 1.4: x = dp / p1: xmax = fk * xT
        xe = x: If xe > xmax Then xe = xmax: chk = "CHOKED (gas, x >= Fk*xT)"
        y = 1# - xe / (3# * fk * xT)
        If xe > 0 And rho1 > 0 Then cvg = wv / (63.3 * y * Sqr(xe * p1 * rho1))
        cvOut(nCV, 12) = "rho1 = " & Format(rho1, "0.000") & " lb/ft3"
        cvOut(nCV, 11) = rho1
    End If
    If wl > 0 And sLrho > 0 Then
        q = wl / sLrho * 0.124675: sg = sLrho / 62.428: dpe = dp
        If wv <= 0 Then
            pv = BubbleP(p1)
            If pv > 0 Then
                ff = 0.96 - 0.28 * Sqr(pv / 3208#)
                dpmax = FL * FL * (p1 - ff * pv)
                If dp > dpmax And dpmax > 0 Then dpe = dpmax: chk = "CHOKED (liquid, dP > FL^2 (P1 - FF Pv))"
                notes = "Pv = " & Format(pv, "0.000") & " psia"
            End If
        End If
        If dpe > 0 Then cvl = q * Sqr(sg / dpe)
        If wv <= 0 Then cvOut(nCV, 11) = sg
        cvOut(nCV, 12) = cvOut(nCV, 12) & IIf(Len(cvOut(nCV, 12)) > 0, "; ", "") & "gpm = " & Format(q, "0.0")
    End If
    cvOut(nCV, 10) = cvg + cvl
    notes = notes & IIf(Len(notes) > 0, "; ", "") & "ISA 75.01 / IEC 60534-2-1 screening" & IIf(wv > 0 And wl > 0, "; two-phase inlet: vapour part + liquid part (additive) - vendor to confirm", "")
    cvOut(nCV, 13) = notes: cvOut(nCV, 20) = chk
    If cvg + cvl > 0 Then cvOut(nCV, 14) = (cvg + cvl) / 0.7
    rated = IDNum(ir, 9)
    If ir > 0 Then
        trv = CvTravel(cvg + cvl, rated, CStr(idat(ir, 12)), IDNum(ir, 13), CStr(idat(ir, 14)), tb, ts, rOut)
        If trv >= 0 Then
            cvOut(nCV, 17) = rOut: cvOut(nCV, 18) = trv: cvOut(nCV, 19) = ts & " [" & tb & "]"
        End If
    End If
End Sub

'---- thermowell materials: E70 (psi), dE/dT (psi/F), density (lb/in3) -------
Private Sub TwMat(ByVal m As String, ByRef e0 As Double, ByRef de As Double, ByRef rm As Double, ByRef ok As Boolean)
    ok = True
    Select Case UCase$(m)
        Case "SS316", "SS316L": e0 = 28300000#: de = -5660: rm = 0.29
        Case "SS304": e0 = 28300000#: de = -5660: rm = 0.289
        Case "CS": e0 = 29500000#: de = -4500: rm = 0.284
        Case "C276": e0 = 29800000#: de = -5000: rm = 0.321
        Case "MONEL": e0 = 26000000#: de = -4000: rm = 0.319
        Case "DUPLEX": e0 = 29000000#: de = -5200: rm = 0.28
        Case "INC625": e0 = 30100000#: de = -5300: rm = 0.305
        Case "1.25CR", "2.25CR": e0 = 30000000#: de = -4800: rm = 0.284
        Case "5CR", "9CR": e0 = 30000000#: de = -4800: rm = 0.28
        Case "ALLOY825": e0 = 28300000#: de = -5000: rm = 0.294
        Case "ALLOY600": e0 = 31000000#: de = -5000: rm = 0.304
        Case "ALLOY20": e0 = 28000000#: de = -5000: rm = 0.292
        Case Else: ok = False
    End Select
End Sub
Private Function MocMat(ByVal moc As String) As String
    Dim a As Variant, i As Long, m As String
    m = UCase$(moc)
    a = Array("C-276", "C276", "ALLOY 400", "MONEL", "ALLOY 825", "ALLOY825", "ALLOY 600", "ALLOY600", "ALLOY 20", "ALLOY20", _
              "DUPLEX", "DUPLEX", "2205", "DUPLEX", "316", "SS316", "304", "SS304", "321", "SS304", "347", "SS304", "317", "SS316", _
              "9% CR", "9CR", "5% CR", "5CR", "2 1/4% CR", "2.25CR", "1 1/4% CR", "1.25CR", "1 1/4%CR", "1.25CR", "CARBON STEEL", "CS")
    For i = 0 To UBound(a) Step 2
        If InStr(m, a(i)) > 0 Then MocMat = a(i + 1): Exit Function
    Next i
End Function

Private Function InstKindOf(ByVal fit As String, ByVal it As String) As String
    Dim f As String, u As String: f = LCase$(Trim$(fit)): u = UCase$(Trim$(it))
    If f = "control valve" Then InstKindOf = "CV": Exit Function
    Select Case u
        Case "PT", "PI", "PRESSURE TRANSMITTER": InstKindOf = "PT"
        Case "TT", "TE", "TW", "TEMPERATURE TRANSMITTER", "THERMOWELL": InstKindOf = "TT"
        Case "FT", "FT-I", "FT INLINE", "FT (INLINE)", "INLINE FT", "FLOW TRANSMITTER", "MAG", "CORIOLIS", "VORTEX": InstKindOf = "FT"
        Case "FO", "FT-O", "FT ORIFICE", "FT (ORIFICE)", "ORIFICE", "FLOW ORIFICE", "FE": InstKindOf = "FO"
        Case "RO", "RESTRICTION ORIFICE": InstKindOf = "RO"
        Case Else
            If f = "orifice" Then InstKindOf = "FO"
    End Select
End Function

Private Sub InstrRows(ByVal cid As String)
    Dim i As Long, k As String, ir As Long, area As Double, qv As Double, ql As Double, vm As Double, mt As Double, rhoh As Double
    Dim P1 As Double, dP As Double, dor As Double, Din As Double, crit As Double, ratio As Double, chk As String, kk As Double
    For i = 1 To stN
        k = stKind(i)
        If Len(k) = 0 Or k = "CV" Then GoTo NextI
        ir = IDRow(stTag(i))
        area = PI / 4# * (stID(i) / 12#) ^ 2
        qv = 0: ql = 0
        If stRg(i) > 0 Then qv = stMg(i) / stRg(i)
        If stRl(i) > 0 Then ql = stMl(i) / stRl(i)
        vm = 0: If area > 0 Then vm = (qv + ql) / 3600# / area
        mt = stMg(i) + stMl(i)
        rhoh = 0: If qv + ql > 0 Then rhoh = mt / (qv + ql)
        P1 = stPin(i): dP = stDp(i): dor = IDNum(ir, 3): Din = stID(i)
        crit = 0: ratio = 0: chk = ""
        kk = IIf(stGm(i) > 0, stGm(i), 1.3)
        If stMg(i) > 0 Then
            If kk > 1 Then crit = (2# / (kk + 1#)) ^ (kk / (kk - 1#)) Else crit = 0.5283
        End If
        If P1 > 0 Then ratio = (P1 - dP) / P1
        If (k = "FO" Or k = "RO") And crit > 0 Then chk = IIf(ratio < crit, "CHOKED (P2/P1 < critical)", "No")
        nInst = nInst + 1
        instOut(nInst, 1) = cid: instOut(nInst, 2) = stTag(i): instOut(nInst, 3) = k & IIf(k = "FT", " (inline)", IIf(k = "FO", " (orifice FT)", ""))
        instOut(nInst, 4) = stLine(i): instOut(nInst, 5) = stSeq(i): instOut(nInst, 6) = stComp(i): instOut(nInst, 7) = stNps(i): instOut(nInst, 8) = stID(i)
        instOut(nInst, 9) = P1: instOut(nInst, 10) = stT(i)
        instOut(nInst, 11) = IIf(stMg(i) > 0 And stMl(i) > 0, "TWO-PHASE", IIf(stMg(i) > 0, "VAPOR", "LIQUID"))
        instOut(nInst, 12) = mt: instOut(nInst, 13) = stMg(i): instOut(nInst, 14) = stMl(i)
        instOut(nInst, 15) = vm: instOut(nInst, 16) = rhoh * LBFT3_TO_KGM3 * (vm * FT_TO_M) ^ 2
        instOut(nInst, 17) = dP: If P1 > 0 Then instOut(nInst, 18) = dP / P1 * 100#
        instOut(nInst, 19) = P1 - dP
        If dor > 0 Then
            instOut(nInst, 20) = dor
            If Din > 0 Then instOut(nInst, 21) = dor / Din
        End If
        If crit > 0 Then instOut(nInst, 22) = crit
        instOut(nInst, 23) = ratio: instOut(nInst, 24) = chk
        If k = "TT" Then TwRow cid, i, ir
NextI:
    Next i
End Sub

Private Sub TwRow(ByVal cid As String, ByVal i As Long, ByVal ir As Long)
    Dim root As Double, tip As Double, bore As Double, U As Double, mat As String, e0 As Double, de As Double, rm As Double, ok As Boolean
    Dim cl As String, j As Long, moc As String, dflt As String, flags As String, dAvg As Double
    Dim qv As Double, ql As Double, rhof As Double, mu As Double, phase As String, area As Double, V As Double
    Dim Re As Double, S As Double, fw As Double, Iin As Double, Am As Double, mpl As Double, wn As Double, fn As Double
    Dim ratio As Double, vlim As Double, verdict As String, E As Double
    cl = UCase$(Trim$(Split(stSpec(i) & "/", "/")(0)))
    root = 0: tip = 0: bore = 0: U = -1: mat = ""
    For j = 1 To nTC
        If UCase$(Trim$(CStr(tcls(j, 1)))) = cl Then
            If IsNumeric(tcls(j, 2)) And Not IsEmpty(tcls(j, 2)) Then root = tcls(j, 2)
            If IsNumeric(tcls(j, 3)) And Not IsEmpty(tcls(j, 3)) Then tip = tcls(j, 3)
            If IsNumeric(tcls(j, 4)) And Not IsEmpty(tcls(j, 4)) Then bore = tcls(j, 4)
            If IsNumeric(tcls(j, 5)) And Not IsEmpty(tcls(j, 5)) Then U = tcls(j, 5)
            If Len(Trim$(CStr(tcls(j, 6)))) > 0 Then mat = Trim$(CStr(tcls(j, 6)))
            flags = "class defaults from Data_TWClass[" & cl & "]"
            Exit For
        End If
    Next j
    If Len(mat) = 0 Then
        moc = ClassMoc(cl)
        If Len(moc) > 0 Then
            mat = MocMat(moc)
            If Len(mat) > 0 Then flags = "material from piping class " & cl & " MOC '" & moc & "'"
        End If
    End If
    If IDHas(ir, 5) Then root = IDNum(ir, 5)
    If IDHas(ir, 6) Then tip = IDNum(ir, 6)
    If IDHas(ir, 7) Then bore = IDNum(ir, 7)
    If IDHas(ir, 4) Then U = IDNum(ir, 4)
    If ir > 0 Then If Len(Trim$(CStr(idat(ir, 8)))) > 0 Then mat = Trim$(CStr(idat(ir, 8)))
    If root <= 0 Then root = 1#: dflt = dflt & "root OD, "
    If tip <= 0 Then tip = 0.75: dflt = dflt & "tip OD, "
    If bore <= 0 Then bore = 0.26: dflt = dflt & "bore, "
    If U < 0 Then
        U = 0.5 * stID(i): If U < 3 Then U = 3
        If U > 12 Then U = 12
        dflt = dflt & "U, "
    End If
    If Len(mat) = 0 Then dflt = dflt & "material, ": mat = "SS316"
    TwMat mat, e0, de, rm, ok
    If Not ok Then mat = "SS316": TwMat mat, e0, de, rm, ok: flags = flags & "; unknown material -> SS316"
    E = e0 + de * (stT(i) - 70#): If E < 10000000# Then E = 10000000#
    dAvg = 0.5 * (root + tip)
    qv = 0: ql = 0
    If stRg(i) > 0 Then qv = stMg(i) / stRg(i)
    If stRl(i) > 0 Then ql = stMl(i) / stRl(i)
    If qv >= ql And stRg(i) > 0 Then
        rhof = stRg(i): mu = stMu(i): phase = "vapour"
    ElseIf stRl(i) > 0 Then
        rhof = stRl(i): mu = stMuL(i): phase = "liquid"
    Else
        rhof = stRg(i): mu = stMu(i): phase = "vapour"
    End If
    If qv > 0 And ql > 0 Then flags = flags & "; two-phase at TW: wake check uses the larger-volume phase"
    area = PI / 4# * (stID(i) / 12#) ^ 2
    If area > 0 Then V = (qv + ql) / 3600# / area
    nTw = nTw + 1
    twOut(nTw, 1) = cid: twOut(nTw, 2) = stTag(i): twOut(nTw, 3) = stLine(i): twOut(nTw, 4) = phase
    If V <= 0 Or rhof <= 0 Or mu <= 0 Then
        twOut(nTw, 19) = "no velocity / fluid properties at this station": twOut(nTw, 20) = flags
        Exit Sub
    End If
    twOut(nTw, 5) = V * FT_TO_M: twOut(nTw, 6) = rhof * LBFT3_TO_KGM3: twOut(nTw, 7) = mu
    twOut(nTw, 8) = mat: twOut(nTw, 9) = E: twOut(nTw, 10) = root: twOut(nTw, 11) = tip: twOut(nTw, 12) = bore: twOut(nTw, 13) = U
    Re = rhof * LBFT3_TO_KGM3 * (V * FT_TO_M) * (dAvg * IN_TO_M) / (mu * 0.001)
    If Re <= 200000# Then
        S = 0.2
    ElseIf Re >= 3500000# Then
        S = 0.27
    Else
        S = 0.2 + 0.07 * Ln10(Re / 200000#) / Ln10(3500000# / 200000#)
    End If
    fw = S * (V * FT_TO_M) / (tip * IN_TO_M)
    Iin = PI / 64# * (dAvg ^ 4 - bore ^ 4): Am = PI / 4# * (dAvg ^ 2 - bore ^ 2)
    mpl = rm * Am + (rhof / 1728#) * PI / 4# * dAvg ^ 2
    wn = 1.875 ^ 2 * Sqr(E * Iin / ((mpl / 386.088) * U ^ 4))
    fn = wn / (2# * PI): ratio = fw / fn
    vlim = 0.8 * fn * (tip * IN_TO_M) / S / FT_TO_M
    If ratio < 0.4 Then
        verdict = "PASS (fs/fn < 0.4)"
    ElseIf ratio < 0.8 Then
        verdict = "CHECK - in-line drag resonance band (0.4-0.8): ASME PTC 19.3 TW stress check"
    Else
        verdict = "FAIL (fs/fn >= 0.8) - shorten U / thicken stem / lower velocity"
    End If
    twOut(nTw, 14) = Re: twOut(nTw, 15) = S: twOut(nTw, 16) = fw: twOut(nTw, 17) = fn: twOut(nTw, 18) = ratio
    twOut(nTw, 19) = vlim * FT_TO_M: twOut(nTw, 20) = verdict
    If Len(dflt) > 0 Then flags = flags & "; DEFAULT used for: " & Left$(dflt, Len(dflt) - 2) & " - enter real data in Instrument_Data / Data_TWClass"
    twOut(nTw, 21) = flags
End Sub
Private Function ClassMoc(ByVal cl As String) As String
    Dim i As Long
    For i = 1 To nPM
        If pmCls(i) = cl Then ClassMoc = pmMoc(i): Exit Function
    Next i
End Function

Private Sub WriteInstr()
    Dim ws As Worksheet, h As Variant, i As Long
    h = Array("Circuit", "Tag", "Type", "Line", "Seq", "Comp ID", "NPS", "ID (in)", "P in (psia)", "T (F)", "Phase", "Mass flow (lb/h)", _
        "Vapour (lb/h)", "Liquid (lb/h)", "Vm (ft/s)", "rho v2 (kg/m s2)", "dP (psi)", "dP % of P1", "P out (psia)", "Orifice d (in)", _
        "beta (d/D)", "Critical P2/P1", "P2/P1", "Choked?")
    Set ws = FreshSheet("Inline_Instruments")
    ws.Range("A1").Resize(1, 24).Value = h: StyleHeader ws.Range("A1").Resize(1, 24)
    If nInst > 0 Then ws.Range("A2").Resize(nInst, 24).Value = instOut: ws.Range("H2:W" & nInst + 1).NumberFormat = "0.000"
    ws.Columns("A:X").ColumnWidth = 13
    i = nInst + 4
    If nTw > 0 Then
        ws.Cells(i, 1).Value = "THERMOWELL WAKE-FREQUENCY CHECK (TT) - simplified Euler-Bernoulli cantilever, average stem diameter, fluid added mass | final check: ASME PTC 19.3 TW by the vendor"
        ws.Cells(i, 1).Font.Bold = True
        h = Array("Circuit", "Tag", "Line", "Fluid used", "Velocity (m/s)", "Fluid rho (kg/m3)", "Fluid mu (cP)", "Material", "E at T (psi)", "Root OD (in)", _
            "Tip OD (in)", "Bore (in)", "Insertion U (in)", "Reynolds", "Strouhal S", "Wake fs (Hz)", "Natural fn (Hz)", "fs / fn", "Max V for fs/fn=0.8 (m/s)", "Verdict", "Notes")
        ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 21)).Value = h: StyleHeader ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 1, 21))
        ws.Range(ws.Cells(i + 2, 1), ws.Cells(i + 1 + nTw, 21)).Value = twOut
        ws.Range(ws.Cells(i + 2, 5), ws.Cells(i + 1 + nTw, 19)).NumberFormat = "0.000"
    End If
End Sub


'==========================================================================
'  Phase 2: tees and branches (split / merge / carry-over / branch runs / join re-flash)
'==========================================================================

Private Function IsTeeF(ByVal f As String) As Boolean
    IsTeeF = (InStr(f, "tee") > 0 Or InStr(f, "stub") > 0 Or InStr(f, "branch") > 0)
End Function
Private Function TrailNum(ByVal f As String) As Long
    Dim i As Long, s As String
    f = Trim$(f)
    i = Len(f)
    Do While i > 0
        If Mid$(f, i, 1) >= "0" And Mid$(f, i, 1) <= "9" Then s = Mid$(f, i, 1) & s: i = i - 1 Else Exit Do
    Loop
    If Len(s) > 0 Then TrailNum = CLng(s) Else TrailNum = 1
End Function
Private Function TeeSplitN(ByVal f As String) As Long
    If InStr(f, "tee") > 0 And InStr(f, "split") > 0 And InStr(f, "branch") > 0 Then TeeSplitN = TrailNum(f)
End Function
Private Function TeeJoinN(ByVal f As String) As Long
    If InStr(f, "tee") > 0 And InStr(f, "join") > 0 And InStr(f, "branch") > 0 Then TeeJoinN = TrailNum(f)
End Function

Private Sub ComputeSeq()
    Dim r As Long, circ As String, ln As String, rt As String, key As String, has As Boolean, q As Double, i As Long
    Dim keys(1 To 400) As String, cnt(1 To 400) As Long, nk As Long
    ReDim rowSeq(1 To nRow)
    For r = 1 To nRow
        If Len(GetS(r, "Circuit")) > 0 Then circ = GetS(r, "Circuit")
        If Len(GetS(r, "Fitting Name")) = 0 And Len(GetS(r, "Comp ID")) = 0 Then GoTo NextS
        q = GetN(r, "Seq", has)
        If has Then
            rowSeq(r) = CLng(q)
        Else
            rt = IIf(IsBranchRow(r), "B", "M")
            key = circ & "|" & GetS(r, "Line No") & "|" & rt
            For i = 1 To nk
                If keys(i) = key Then Exit For
            Next i
            If i > nk And nk < 400 Then nk = nk + 1: keys(nk) = key: cnt(nk) = 0: i = nk
            cnt(i) = cnt(i) + 1: rowSeq(r) = cnt(i)
        End If
NextS:
    Next r
End Sub
Private Function IsBranchRow(ByVal r As Long) As Boolean
    IsBranchRow = (UCase$(Left$(GetS(r, "Run Type"), 1)) = "B")
End Function

Private Sub AddSplit(ByVal ln As String, ByVal sq As Long, ByVal n As Long, ByVal v As Double, ByVal l As Double)
    Dim i As Long
    For i = 1 To nSplit
        If spLine(i) = ln And spSeq(i) = sq Then spN(i) = n: spV(i) = v: spL(i) = l: Exit Sub
    Next i
    If nSplit >= 60 Then Exit Sub
    nSplit = nSplit + 1
    spLine(nSplit) = ln: spSeq(nSplit) = sq: spN(nSplit) = n: spV(nSplit) = v: spL(nSplit) = l
End Sub

Private Sub SaveFeed(ByVal sl As Long)
    Dim i As Long
    bkN(sl) = fN: bkHas(sl) = fHasC: bkPref(sl) = fPref: bkTref(sl) = fTref: bkTot(sl) = fTotMoles
    bkVap(sl) = sVap: bkLiq(sl) = sLiq: bkFlash(sl) = flashOn
    If flashOn Then
        For i = 1 To fN
            bkNm(sl, i) = fNm(i): bkZ(sl, i) = fZ(i): bkK(sl, i) = fKref(i): bkMW(sl, i) = fMWc(i): bkTc(sl, i) = fTcR(i): bkOm(sl, i) = fOm(i)
        Next i
    End If
End Sub
Private Sub RestoreFeed(ByVal sl As Long)
    Dim i As Long
    fN = bkN(sl): fHasC = bkHas(sl): fPref = bkPref(sl): fTref = bkTref(sl): fTotMoles = bkTot(sl)
    sVap = bkVap(sl): sLiq = bkLiq(sl): flashOn = bkFlash(sl)
    If flashOn Then
        ReDim fNm(1 To fN): ReDim fZ(1 To fN): ReDim fKref(1 To fN): ReDim fMWc(1 To fN): ReDim fTcR(1 To fN): ReDim fOm(1 To fN)
        For i = 1 To fN
            fNm(i) = bkNm(sl, i): fZ(i) = bkZ(sl, i): fKref(i) = bkK(sl, i): fMWc(i) = bkMW(sl, i): fTcR(i) = bkTc(sl, i): fOm(i) = bkOm(sl, i)
        Next i
    End If
End Sub

Private Function InjSlot(ByVal r As Long) As Long
    Dim i As Long
    For i = 1 To nInj
        If injRow(i) = r Then InjSlot = i: Exit Function
    Next i
End Function

' keep the state at the end of a branch run for a later join into the Main
Private Sub CaptureInj(ByVal row As Long, ByVal lbl As String)
    Dim i As Long, s As Long, fm As Double, fl As Double
    If nInj >= 20 Then Exit Sub
    nInj = nInj + 1: s = nInj
    injRow(s) = row: injLbl(s) = lbl: injVap(s) = eVap: injLiq(s) = eLiq: injFlash(s) = flashOn
    If flashOn And fN <= 600 Then
        injN(s) = fN: injPref(s) = fPref
        fm = eBeta * fTotMoles * gMult * fScale * vCf: fl = (1# - eBeta) * fTotMoles * gMult * fScale * lCf
        For i = 1 To fN
            injNm(s, i) = fNm(i): injMol(s, i) = eY(i) * fm + eX(i) * fl: injK(s, i) = fKref(i)
            injMW(s, i) = fMWc(i): injTc(s, i) = fTcR(i): injOm(s, i) = fOm(i)
        Next i
    End If
End Sub

' blend the branch outlet into the running Main (new feed, re-flashed from here on)
Private Sub DoJoin(ByVal sl As Long, ByVal p As Double)
    Dim i As Long, j As Long, n As Long, fm As Double, fl As Double, tot As Double, wa As Double, wb As Double, fac As Double
    Dim nm() As String, mol() As Double, kr() As Double, mw() As Double, tc() As Double, om() As Double, a As Long, F As Double
    Dim inA() As Boolean
    If Not flashOn Or Not injFlash(sl) Then
        ' frozen split (manual / no composition): add the phase flows
        sVap = sVap * fScale * vCf + injVap(sl) / gMult: sLiq = sLiq * fScale * lCf + injLiq(sl) / gMult
        fScale = 1: vCf = 1: lCf = 1
        Exit Sub
    End If
    FlashAt IIf(p < 0.05, 0.05, p)
    fm = eBeta * fTotMoles * gMult * fScale * vCf: fl = (1# - eBeta) * fTotMoles * gMult * fScale * lCf
    n = fN + injN(sl)
    ReDim nm(1 To n): ReDim mol(1 To n): ReDim kr(1 To n): ReDim mw(1 To n): ReDim tc(1 To n): ReDim om(1 To n): ReDim inA(1 To n)
    For i = 1 To fN
        nm(i) = fNm(i): mol(i) = eY(i) * fm + eX(i) * fl: kr(i) = fKref(i): mw(i) = fMWc(i): tc(i) = fTcR(i): om(i) = fOm(i)
    Next i
    a = fN
    fac = injPref(sl) / fPref
    For j = 1 To injN(sl)
        For i = 1 To fN
            If nm(i) = injNm(sl, j) Then Exit For
        Next i
        If i <= fN Then
            wa = mol(i): wb = injMol(sl, j)
            If wa + wb > 0 Then kr(i) = (wa * kr(i) + wb * injK(sl, j) * fac) / (wa + wb)
            mol(i) = mol(i) + injMol(sl, j)
            If tc(i) < -9000 Then tc(i) = injTc(sl, j): om(i) = injOm(sl, j)
        Else
            a = a + 1: nm(a) = injNm(sl, j): mol(a) = injMol(sl, j): kr(a) = injK(sl, j) * fac
            mw(a) = injMW(sl, j): tc(a) = injTc(sl, j): om(a) = injOm(sl, j)
        End If
    Next j
    n = 0: F = 0
    For i = 1 To a
        If mol(i) > 0 Then F = F + mol(i)
    Next i
    If F <= 0 Then Exit Sub
    ReDim fNm(1 To a): ReDim fZ(1 To a): ReDim fKref(1 To a): ReDim fMWc(1 To a): ReDim fTcR(1 To a): ReDim fOm(1 To a)
    fHasC = False
    For i = 1 To a
        If mol(i) > 0 Then
            n = n + 1
            fNm(n) = nm(i): fZ(n) = mol(i) / F: fKref(n) = ClampK(kr(i), 10000000000#)
            fMWc(n) = IIf(mw(i) > 0, mw(i), 50): fTcR(n) = tc(i): fOm(n) = om(i)
            If tc(i) > -9000 Then fHasC = True
        End If
    Next i
    fN = n: fTotMoles = F / gMult
    fScale = 1: vCf = 1: lCf = 1
End Sub

' ---- circuit driver: main + branch blocks --------------------------------
Private Sub RunBranches(ByVal cid As String, ByVal nB As Long, bS() As Long, bE() As Long, brMap() As Long, ByVal pMain As Double, _
                        ByVal wr As Boolean, ByVal capture As Boolean, mainMap() As Long, ByVal nM As Long, ByVal cr1 As Long, ByVal cr2 As Long)
    Dim b As Long, k As Long, i As Long, bp As Double, has As Boolean, upLine As String, upSeq As Double, hasUS As Boolean
    Dim capI As Long, cnt As Long, n As Long, r0 As Long, bline As String, lbl As String, jIdx As Long, pe As Double
    Dim own As String, standalone As Boolean, saved As Boolean, kp As Long, rr As Long
    Dim lnCnt(1 To 30) As String, lnN(1 To 30) As Long, nl As Long, jLine(1 To 30) As String, jN(1 To 30) As Long, nj As Long
    For b = 1 To nB
        r0 = brMap(bS(b)): bline = GetS(r0, "Line No")
        lbl = cid & " [Branch " & b & "]"
        bp = pMain
        For k = bS(b) To bE(b)
            upSeq = GetN(brMap(k), "Start P (psia)", has)
            If has Then bp = upSeq: Exit For
        Next k
        ' ---- seed from a Main 'Tee Split Branch Flow N' capture
        upLine = GetS(r0, "Upstream Line No"): upSeq = GetN(r0, "Upstream Seq", hasUS): capI = 0
        If Len(upLine) > 0 Then
            If hasUS Then
                For i = 1 To nSplit
                    If spLine(i) = upLine And spSeq(i) = CLng(upSeq) Then capI = i: Exit For
                Next i
            Else
                cnt = 0
                For i = 1 To nl
                    If lnCnt(i) = upLine Then cnt = lnN(i) + 1: lnN(i) = cnt: Exit For
                Next i
                If cnt = 0 Then nl = nl + 1: lnCnt(nl) = upLine: lnN(nl) = 1: cnt = 1
                For i = 1 To nSplit
                    If spLine(i) = upLine And spN(i) = cnt Then capI = i: Exit For
                Next i
            End If
        End If
        fScale = 1: vCf = 1: lCf = 1
        standalone = (capI = 0)
        saved = False
        If capI > 0 Then
            vCf = spV(capI): lCf = spL(capI)
        ElseIf Len(GetS(r0, "Stream Lookup")) > 0 Then
            saved = True
            If Not LoadStream(brMap(bS(b)), brMap(bE(b))) Then GoTo EndB
        End If
        rowMap = brMap
        useInj = False: cdNew = True: stnNo = 0
        pe = March(bS(b), bE(b), lbl, bp, wr, 0)
        If wr Then FlushStations lbl
        If capture Then
            ' join row on the Main: explicit Upstream Line No = this branch line, else Nth join with this line
            jIdx = 0: cnt = 1
            For i = 1 To nj
                If jLine(i) = bline Then jN(i) = jN(i) + 1: cnt = jN(i): Exit For
            Next i
            If i > nj Then nj = nj + 1: jLine(nj) = bline: jN(nj) = 1: cnt = 1
            For k = 1 To nM
                rr = mainMap(k)
                If TeeJoinN(LCase$(GetS(rr, "Fitting Name"))) > 0 And GetS(rr, "Upstream Line No") = bline And Not IsBranchRow(rr) Then jIdx = rr: Exit For
            Next k
            If jIdx = 0 Then
                For k = 1 To nM
                    rr = mainMap(k)
                    If GetS(rr, "Line No") = bline And TeeJoinN(LCase$(GetS(rr, "Fitting Name"))) = cnt Then jIdx = rr: Exit For
                Next k
            End If
            If jIdx > 0 Then CaptureInj jIdx, lbl
        End If
EndB:
        If saved Then If Not LoadStream(cr1, cr2) Then Exit Sub
    Next b
End Sub

'==========================================================================
Private Sub WriteOutputs()
    Dim ws As Worksheet, hdr As Variant, i As Long, nc As Long
    hdr = Array("Circuit", "Line No", "Comp ID", "Fitting", "Spec", "NPS", "Sch", "ID (in)", "Length (ft)", "Elev (ft)", _
        "Cum L (ft)", "P in (psia)", "P out (psia)", "Velocity (ft/s)", "Re", "f", "Density (lb/ft3)", "dP fric (psi)", _
        "dP fit (psi)", "dP elev (psi)", "dP total (psi)", "Phase", "Stream", "Note", "Vap frac (mass)", "Beta (mol)", "Vapour (lb/h)", "Liquid (lb/h)")
    Set ws = FreshSheet("Pressure_Profile")
    ws.Range("A1").Resize(1, 28).Value = hdr
    StyleHeader ws.Range("A1").Resize(1, 28)
    If nOut > 0 Then ws.Range("A2").Resize(nOut, 28).Value = out
    ws.Range("I2:K" & nOut + 1).NumberFormat = "0.00"
    ws.Range("L2:M" & nOut + 1).NumberFormat = "0.000"
    ws.Range("N2:N" & nOut + 1).NumberFormat = "0.00"
    ws.Range("O2:O" & nOut + 1).NumberFormat = "#,##0"
    ws.Range("P2:P" & nOut + 1).NumberFormat = "0.0000"
    ws.Range("Q2:Q" & nOut + 1).NumberFormat = "0.000"
    ws.Range("R2:U" & nOut + 1).NumberFormat = "0.000"
    ws.Columns("A:AB").ColumnWidth = 12: ws.Columns("D").ColumnWidth = 24: ws.Columns("X").ColumnWidth = 70
    ws.Range("Y2:Z" & nOut + 1).NumberFormat = "0.0000": ws.Range("AA2:AB" & nOut + 1).NumberFormat = "#,##0"
    ws.Activate: ActiveWindow.FreezePanes = False: ws.Range("D2").Select: ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    hdr = Array("Circuit", "CV Tag", "Stream", "Phase", "P1 (psia)", "P2 (psia)", "dP (psi)", "Flow (lb/h)", "T (degF)", _
        "Calc Cv", "SG / rho1", "Flow info", "Notes", "Suggested rated Cv (70% travel)", "FL", "xT", "Rated Cv", "Travel (%)", "Travel status", "Choked?")
    Set ws = FreshSheet("CV_Summary")
    ws.Range("A1").Resize(1, 20).Value = hdr
    StyleHeader ws.Range("A1").Resize(1, 20)
    If nCV > 0 Then ws.Range("A2").Resize(nCV, 20).Value = cvOut
    ws.Range("E2:G" & nCV + 2).NumberFormat = "0.000": ws.Range("J2:K" & nCV + 2).NumberFormat = "0.00"
    ws.Range("N2:N" & nCV + 2).NumberFormat = "0.0"
    ws.Columns("A:T").ColumnWidth = 13: ws.Columns("M").ColumnWidth = 60: ws.Columns("S").ColumnWidth = 40
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
