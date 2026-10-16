Option VBASupport 1
Option Explicit
' LibreOffice에서 TotalAssets.bas의 계산부를 돌려 보는 점검용 모듈 (엑셀에 넣는 것이 아님)

Private Function Utf8Read(ByVal path As String) As String
    Dim sfa As Object, ins As Object, tis As Object, s As String
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    ins = sfa.openFileRead(ConvertToURL(path))
    tis = CreateUnoService("com.sun.star.io.TextInputStream")
    tis.setInputStream(ins)
    tis.setEncoding("UTF-8")
    Do While Not tis.isEOF()
        s = s & tis.readLine() & Chr(10)
    Loop
    tis.closeInput()
    Utf8Read = s
End Function

Private Sub Utf8Write(ByVal path As String, ByVal text As String)
    Dim sfa As Object, outs As Object, tos As Object
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If sfa.exists(ConvertToURL(path)) Then sfa.kill(ConvertToURL(path))
    outs = sfa.openFileWrite(ConvertToURL(path))
    tos = CreateUnoService("com.sun.star.io.TextOutputStream")
    tos.setOutputStream(outs)
    tos.setEncoding("UTF-8")
    tos.writeString(text)
    tos.closeOutput()
End Sub

' 엑셀 파일의 모든 시트를 엑셀 UsedRange.Value2처럼 1부터 시작하는 2차원 배열로 읽는다
Private Function LoadSheets(ByVal path As String, ByRef nS As Long, Optional ByRef names As Variant) As Variant
    Dim args(1) As New com.sun.star.beans.PropertyValue, doc As Object, sh As Object, cur As Object, data As Variant
    Dim k As Long, r As Long, c As Long, nr As Long, nc As Long, out() As Variant
    args(0).Name = "Hidden": args(0).Value = True
    args(1).Name = "ReadOnly": args(1).Value = True
    doc = StarDesktop.loadComponentFromURL(ConvertToURL(path), "_blank", 0, args())
    nS = doc.Sheets.Count
    ReDim out(nS - 1)
    Dim nm() As String
    ReDim nm(nS - 1)
    For k = 0 To nS - 1
        nm(k) = doc.Sheets.getByIndex(k).Name
    Next
    names = nm
    For k = 0 To nS - 1
        sh = doc.Sheets.getByIndex(k)
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        nr = cur.RangeAddress.EndRow + 1: nc = cur.RangeAddress.EndColumn + 1
        data = sh.getCellRangeByPosition(0, 0, nc - 1, nr - 1).getDataArray()
        Dim a() As Variant
        ReDim a(1 To nr, 1 To nc)
        For r = 1 To nr
            For c = 1 To nc
                If VarType(data(r - 1)(c - 1)) = 8 And data(r - 1)(c - 1) = "" Then
                    a(r, c) = Empty
                Else
                    a(r, c) = data(r - 1)(c - 1)
                End If
            Next
        Next
        out(k) = a
    Next
    doc.close(True)
    LoadSheets = out
End Function

Private Function Fmt6(ByVal x As Double) As String
    Fmt6 = Format(x, "0.000000")
End Function

Private Function ValText(v As Variant) As String
    If IsEmpty(v) Then
        ValText = ""
    Else
        ValText = Format(Int(CDbl(v) + 0.5), "0")
    End If
End Function

' 실제 파일로 끝까지: 대상 엑셀 + 장부 폴더 → 행별 판정
Function Ping(ByVal s As String) As String
    On Error GoTo Fail
    Ping = NormName(s) & "|" & CoreName(s) & "|" & NumsKey(s)
    Exit Function
Fail:
    Ping = "ERROR " & Err & " " & Error$ & " @" & Erl
End Function

Function Full(ByVal target As String, ByVal dir_ As String, ByVal baseYmd As String) As String
    On Error GoTo Fail
    Dim outPath As String, nS As Long, sheets As Variant, k As Long, ok As Boolean
    Dim f As String, files() As String, nF As Long, lds() As Variant, i As Long, txt As String, ld As Variant
    sheets = LoadSheets(target, nS)
    For k = 0 To nS - 1
        If ParseTargetArray(sheets(k)) Then ok = True: Exit For
    Next
    If Not ok Then Full = "ERROR target": Exit Function
    If baseYmd = "" Then baseYmd = tBaseDate
    If baseYmd = "-" Then baseYmd = ""
    nF = 0
    ReDim files(0)
    f = Dir(dir_ & "/*")
    Do While f <> ""
        ReDim Preserve files(nF): files(nF) = f: nF = nF + 1
        f = Dir()
    Loop
    ' 파일 이름 순서로 (파이썬판과 같은 순서)
    Dim a As Long, b As Long, tmp As String
    For a = 0 To nF - 2
        For b = a + 1 To nF - 1
            If StrComp(files(b), files(a), 0) < 0 Then tmp = files(a): files(a) = files(b): files(b) = tmp
        Next
    Next
    Dim nm As Variant, part As Variant, nPart As Long, j As Long, nL As Long
    ReDim lds(0): nL = 0
    For i = 0 To nF - 1
        sheets = LoadSheets(dir_ & "/" & files(i), nS, nm)
        part = BuildLedgers(dir_ & "/" & files(i), files(i), sheets, nS, "", nm, nPart)
        For j = 0 To nPart - 1
            ReDim Preserve lds(nL)
            lds(nL) = part(j)
            nL = nL + 1
        Next
    Next
    MarkSettlement lds, nL
    ComputeResults lds, nL, baseYmd, False, False, Nothing
    txt = "B" & Chr(9) & baseYmd & Chr(10)
    For i = 0 To nL - 1
        ld = lds(i)
        txt = txt & "L" & Chr(9) & ld(0) & Chr(9) & ld(1) & Chr(9) & ld(2) & Chr(9) & ld(3) & Chr(9) & ld(4) & Chr(9) & ld(5) & Chr(9) & ValText(ld(8)) & Chr(9) & ld(9) & Chr(9) & ld(11) & Chr(9) & ld(18) & Chr(9) & IIf(ld(19), "1", "") & Chr(10)
    Next
    For i = 0 To tCount - 1
        txt = txt & "R" & Chr(9) & tRow(i) & Chr(9) & rStatus(i) & Chr(9) & IIf(rInc(i), "Y", "") & Chr(9) & Fmt6(rScore(i)) & Chr(9)
        If rLed(i) >= 0 Then txt = txt & lds(rLed(i))(0)
        txt = txt & Chr(9) & rWhy(i) & Chr(9) & rNotes(i) & Chr(10)
    Next
    Dim nMis As Long, common As String
    common = DateMismatch(lds, baseYmd, nMis)
    txt = txt & "D" & Chr(9) & common & Chr(9) & nMis & Chr(10)
    Full = txt
    Exit Function
Fail:
    Full = "ERROR " & Err & " " & Error$ & " @" & Erl
End Function

' 매칭 엔진만: 입력 파일의 T(대상)/L(장부, 운용사) 줄 → 점수·근거·배정·상태
Private Function MakeTestLedger(ByVal nm As String, ByVal mgr As String) As Variant
    Dim ld(19) As Variant, cc(0) As String
    ld(1) = nm: ld(2) = "": ld(3) = mgr
    cc(0) = nm
    ld(16) = cc: ld(17) = 1
    MakeTestLedger = ld
End Function

Function Parity(ByVal inText As String) As String
    On Error GoTo Fail
    Dim lines As Variant, ln As Variant, parts As Variant, names() As String, nT As Long, lds() As Variant, nL As Long, i As Long, l As Long
    Dim forced() As Long, txt As String, ld As Variant, cands(0) As String
    lines = Split(inText, Chr(10))
    ReDim names(0): ReDim lds(0)
    For Each ln In lines
        If ln <> "" Then
            parts = Split(ln, Chr(9))
            If parts(0) = "T" Then
                ReDim Preserve names(nT): names(nT) = parts(1): nT = nT + 1
            ElseIf parts(0) = "L" Then
                Dim mgrName As String
                mgrName = ""
                If UBound(parts) >= 2 Then mgrName = parts(2)
                ReDim Preserve lds(nL): lds(nL) = MakeTestLedger(parts(1), mgrName): nL = nL + 1
            End If
        End If
    Next
    ReDim forced(nT - 1)
    For i = 0 To nT - 1
        forced(i) = -2
    Next
    MatchNames names, nT, lds, nL, forced
    For i = 0 To nT - 1
        For l = 0 To nL - 1
            txt = txt & "S" & Chr(9) & i & Chr(9) & l & Chr(9) & Format(mScores(i, l), "0.000000000000") & Chr(9) & mWhys(i, l) & Chr(10)
        Next
        If mAssign(i) >= 0 Then txt = txt & "A" & Chr(9) & i & Chr(9) & mAssign(i) & Chr(9) & MatchStatus(i, mAssign(i), nL) & Chr(10)
    Next
    Parity = txt
    Exit Function
Fail:
    Parity = "ERROR " & Err & " " & Error$ & " @" & Erl
End Function

Function Dbg() As String
    Dim stepName As String, c(1) As String, why As String, cands(0) As String
    On Error GoTo Fail
    stepName = "LcsRatio": Dbg = Dbg & stepName & "=" & LcsRatio("가온101", "가온1011") & ";"
    stepName = "Cosine(noidf)": Dbg = Dbg & stepName & "=" & CosineSim("가온101", "가온1011") & ";"
    c(0) = "가온101": c(1) = "누리2"
    stepName = "BuildIdf": BuildIdf c, 2: Dbg = Dbg & stepName & " ok;"
    stepName = "Cosine(idf)": Dbg = Dbg & stepName & "=" & CosineSim("가온101", "가온1011") & ";"
    cands(0) = "가온101호"
    stepName = "ScoreName": Dbg = Dbg & stepName & "=" & ScoreName("가온 101호", cands, 1, "", why) & " " & why & ";"
    Exit Function
Fail:
    Dbg = Dbg & " ERROR at " & stepName & ": " & Err & " " & Error$
End Function

Function Dbg4() As String
    Dim st As String, sl As Long
    On Error GoTo Fail
    st = "slot1": sl = GramSlot("가온101"): Dbg4 = Dbg4 & st & "=" & sl & ";"
    st = "n": Dbg4 = Dbg4 & st & "=" & gGramN(sl) & ";"
    st = "keys": Dbg4 = Dbg4 & st & "=" & gGramKeys(sl)(0) & ";"
    st = "slot2": sl = GramSlot("가온1011"): Dbg4 = Dbg4 & st & "=" & sl & ";"
    st = "again": sl = GramSlot("가온101"): Dbg4 = Dbg4 & st & "=" & sl & ";"
    Exit Function
Fail:
    Dbg4 = Dbg4 & " ERROR at " & st & ": " & Err & " " & Error$ & " @" & Erl
End Function


' 파일명 목록 → 정리된 이름과 호수 키 (줄마다 "이름<TAB>숫자키")
Function CleanNames(ByVal inText As String) As String
    Dim ln As Variant, t As String, txt As String
    On Error GoTo Fail
    For Each ln In Split(inText, Chr(10))
        If ln <> "" Then
            t = CleanFileName(ln)
            txt = txt & t & Chr(9) & NumsKey(t) & Chr(10)
        End If
    Next
    CleanNames = txt
    Exit Function
Fail:
    CleanNames = "ERROR " & Err & " " & Error$ & " @" & Erl
End Function
