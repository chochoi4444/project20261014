Option Explicit
' ======================================================================
'  총자산 자동입력 - 엑셀 매크로판
'
'  기준가격대장(엑셀)에서 총자산을 찾아 대상 엑셀의 '총자산' 칸에 넣는다.
'  웹·파이썬판과 같은 규칙(펀드명 코사인 유사도 매칭, 순자산+부채 검증).
'  외부 라이브러리·인터넷이 필요 없다.
'
'  설치: 엑셀 > Alt+F11 > 삽입 > 모듈 > 이 글 전체를 붙여 넣기
'  실행: Alt+F8 > RunTotalAssets > 실행
'  (선택) 이 매크로가 든 통합문서에 '매핑' 시트를 만들고
'        A열 = 대상 엑셀 종목명, B열 = 장부 펀드명 또는 파일명 (매칭하지 않으려면 '없음')
' ======================================================================

Private Const AUTO_SCORE As Double = 0.85
Private Const AUTO_MARGIN As Double = 0.08
Private Const REVIEW_SCORE As Double = 0.55
Private Const AUTO_CLEAR_SCORE As Double = 0.75
Private Const AUTO_CLEAR_MARGIN As Double = 0.4
Private Const COSINE_WEIGHT As Double = 0.6

' VarType 값 (이름 상수 대신 숫자)
Private Const VT_INTEGER As Long = 2
Private Const VT_LONG As Long = 3
Private Const VT_SINGLE As Long = 4
Private Const VT_DOUBLE As Long = 5
Private Const VT_CURRENCY As Long = 6
Private Const VT_DATE As Long = 7
Private Const VT_STRING As Long = 8
Private Const VT_DECIMAL As Long = 14
Private Const VT_BYTE As Long = 17
Private Const CMP_BINARY As Long = 0

' 장부 레코드 칸 번호
Private Const LD_FILE As Long = 0
Private Const LD_NAME As Long = 1
Private Const LD_TITLE As Long = 2
Private Const LD_MANAGER As Long = 3
Private Const LD_CODE As Long = 4
Private Const LD_DATE As Long = 5
Private Const LD_KEY As Long = 6
Private Const LD_ERROR As Long = 7
Private Const LD_VALUE As Long = 8
Private Const LD_METHOD As Long = 9
Private Const LD_CHECKS As Long = 10
Private Const LD_CHECKT As Long = 11
Private Const LD_NET As Long = 12
Private Const LD_LIAB As Long = 13
Private Const LD_NAV As Long = 14
Private Const LD_PATH As Long = 15
Private Const LD_CANDS As Long = 16
Private Const LD_NCANDS As Long = 17
Private Const LD_SHEET As Long = 18
Private Const LD_SUPER As Long = 19   ' 같은 이름의 결산후 파일이 있어 뺀 결산전 장부
Private Const LD_SIZE As Long = 20

' ---------- 매칭 결과 (MatchNames가 채움) ----------
Public mScores() As Double
Public mWhys() As String
Public mAssign() As Long

' ---------- 대상 엑셀 (ParseTargetArray가 채움) ----------
Public tCount As Long
Public tRow() As Long
Public tName() As String
Public tCode() As String
Public tCls() As String
Public tExist() As Variant
Public tAssetColRel As Long
Public tHeaderRow As Long
Public tBaseDate As String
Public tAssetHeader As String

' ---------- 행별 결과 (ComputeResults가 채움) ----------
Public rStatus() As String
Public rLed() As Long
Public rScore() As Double
Public rWhy() As String
Public rInc() As Boolean
Public rNotes() As String
Public uNames() As String
Public uCount As Long
Public rU() As Long

' ---------- 내부 ----------
Private gGramSlot As Collection
Private gGramKeys() As Variant
Private gGramCnts() As Variant
Private gGramN() As Long
Private gGramIdx() As Variant
Private gGramCount As Long
Private gIdfIdx As Collection
Private gIdf() As Double
Private gIdfUnseen As Double
Private gUseIdf As Boolean

' ======================================================================
'  문자열 도우미
' ======================================================================
Private Function CK(ByVal s As String) As String
    ' Collection 키는 대소문자를 구분하지 않으므로 글자 코드로 바꿔 쓴다
    Dim i As Long, out As String
    out = "k"
    For i = 1 To Len(s)
        out = out & Hex$(AscW(Mid$(s, i, 1)) And &HFFFF&) & "."
    Next
    CK = out
End Function

Private Function CollHas(c As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c.Item(key)
    If Err.Number = 0 Then
        CollHas = True
    Else
        Err.Clear
        Dim o As Object
        Set o = c.Item(key)
        CollHas = (Err.Number = 0)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function CollGetLong(c As Collection, ByVal key As String, ByVal dflt As Long) As Long
    Dim v As Variant
    On Error Resume Next
    v = c.Item(key)
    If Err.Number <> 0 Then
        CollGetLong = dflt
    Else
        CollGetLong = CLng(v)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function CodeOf(ByVal ch As String) As Long
    CodeOf = AscW(ch) And &HFFFF&
End Function

Private Function IsDigitCh(ByVal ch As String) As Boolean
    IsDigitCh = (ch >= "0" And ch <= "9" And Len(ch) = 1)
End Function

Private Function IsUpperAZ(ByVal ch As String) As Boolean
    Dim w As Long
    If Len(ch) <> 1 Then Exit Function
    w = CodeOf(ch)
    IsUpperAZ = (w >= 65 And w <= 90)
End Function

Private Function IsHangul(ByVal ch As String) As Boolean
    Dim w As Long
    If Len(ch) <> 1 Then Exit Function
    w = CodeOf(ch)
    IsHangul = (w >= &HAC00& And w <= &HD7A3&)
End Function

Private Function IsWS(ByVal ch As String) As Boolean
    Dim w As Long
    If Len(ch) <> 1 Then Exit Function
    w = CodeOf(ch)
    IsWS = (w = 32 Or (w >= 9 And w <= 13) Or w = &HA0& Or w = &H3000& Or w = &HFEFF&)
End Function

Private Function HasLetters(ByVal s As String) As Boolean
    Dim i As Long, ch As String, w As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        w = CodeOf(ch)
        If IsHangul(ch) Or (w >= 65 And w <= 90) Or (w >= 97 And w <= 122) Then HasLetters = True: Exit Function
    Next
End Function

Private Function TrimWS(ByVal s As String) As String
    Dim a As Long, b As Long
    a = 1: b = Len(s)
    Do While a <= b
        If Not IsWS(Mid$(s, a, 1)) Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If Not IsWS(Mid$(s, b, 1)) Then Exit Do
        b = b - 1
    Loop
    If b >= a Then
        TrimWS = Mid$(s, a, b - a + 1)
    Else
        TrimWS = ""
    End If
End Function

Private Function RemoveWS(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not IsWS(ch) Then out = out & ch
    Next
    RemoveWS = out
End Function

Private Function RomanNum(ByVal k As Long) As String
    Select Case k
        Case 0: RomanNum = "I"
        Case 1: RomanNum = "II"
        Case 2: RomanNum = "III"
        Case 3: RomanNum = "IV"
        Case 4: RomanNum = "V"
        Case 5: RomanNum = "VI"
        Case 6: RomanNum = "VII"
        Case 7: RomanNum = "VIII"
        Case 8: RomanNum = "IX"
        Case 9: RomanNum = "X"
        Case 10: RomanNum = "XI"
        Case Else: RomanNum = "XII"
    End Select
End Function

' 유니코드 호환 정규화(NFKC) 중 펀드명에 나오는 부분: 전각 문자, 로마 숫자, 원 숫자, ㈜
Public Function NormText(ByVal s As String) As String
    Dim i As Long, w As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        w = CodeOf(ch)
        If w >= &HFF01& And w <= &HFF5E& Then
            out = out & ChrW(w - &HFEE0&)
        ElseIf w = &H3000& Or w = &HA0& Then
            out = out & " "
        ElseIf w >= &H2160& And w <= &H216B& Then
            out = out & RomanNum(w - &H2160&)
        ElseIf w >= &H2170& And w <= &H217B& Then
            out = out & LCase$(RomanNum(w - &H2170&))
        ElseIf w >= &H2460& And w <= &H2473& Then
            out = out & CStr(w - &H2460& + 1)
        ElseIf w = &H321C& Then
            out = out & "(" & ChrW(&HC8FC&) & ")"
        Else
            out = out & ch
        End If
    Next
    NormText = out
End Function

Public Function NormLabel(ByVal s As String) As String
    Dim t As String, i As Long, ch As String, out As String
    t = NormText(s)
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If Not (IsWS(ch) Or ch = "[" Or ch = "]" Or ch = ":" Or ch = ChrW(&H25B6&) Or ch = ChrW(&HB7&)) Then out = out & ch
    Next
    NormLabel = out
End Function

Private Function ReplaceToken(ByVal t As String, ByVal tok As String, ByVal rep As String, ByVal checkPrev As Boolean, ByVal checkNext As Boolean) As String
    Dim i As Long, n As Long, out As String, ok As Boolean
    n = Len(tok)
    i = 1
    Do While i <= Len(t)
        ok = False
        If Mid$(t, i, n) = tok Then
            ok = True
            If checkPrev And i > 1 Then If IsUpperAZ(Mid$(t, i - 1, 1)) Then ok = False
            If checkNext And i + n <= Len(t) Then If IsUpperAZ(Mid$(t, i + n, 1)) Then ok = False
        End If
        If ok Then
            out = out & rep
            i = i + n
        Else
            out = out & Mid$(t, i, 1)
            i = i + 1
        End If
    Loop
    ReplaceToken = out
End Function

Public Function NormName(ByVal s As String) As String
    Dim t As String, i As Long, ch As String, out As String
    t = UCase$(NormText(s))
    t = ReplaceToken(t, "UP", "유나이티드파트너스", True, True)
    t = Replace(t, "BLACKROCK", "블랙락")
    t = Replace(t, "블랙록", "블랙락")
    t = Replace(t, "SAMSUNG", "삼성")
    t = Replace(t, "SHINHAN", "신한")
    t = Replace(t, "KIWOOM", "키움")
    t = ReplaceToken(t, "HANA", "하나", False, True)
    t = Replace(t, "IGIS", "이지스")
    t = Replace(t, "HYUNDAI", "현대")
    t = Replace(t, "KYOBO", "교보")
    t = Replace(t, "MIRAE", "미래에셋")
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If IsHangul(ch) Or IsUpperAZ(ch) Or IsDigitCh(ch) Then out = out & ch
    Next
    NormName = out
End Function

Public Function CoreName(ByVal s As String) As String
    Dim t As String, w As Variant, i As Long, ch As String, out As String
    t = NormName(s)
    For Each w In Array("전문투자형사모", "전문투자형", "일반사모", "전문사모", "사모", "공모", "투자신탁", "투자회사", "투자조합", _
                        "신탁", "투자", "펀드", "재간접형", "재간접", "증권", "자투자", "모투자", "유한회사", "주식회사", "전문", "일반", _
                        "CLASS", "클래스")
        t = Replace(t, CStr(w), "")
    Next
    ' '제'+숫자 → 숫자
    out = ""
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch = "제" And i < Len(t) Then
            If IsDigitCh(Mid$(t, i + 1, 1)) Then GoTo NextA
        End If
        out = out & ch
NextA:
    Next
    t = out
    ' 숫자+'호'/'종' → 숫자
    out = ""
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If (ch = "호" Or ch = "종") And i > 1 Then
            If IsDigitCh(Mid$(t, i - 1, 1)) Then GoTo NextB
        End If
        out = out & ch
NextB:
    Next
    CoreName = out
End Function

' 이름 속 숫자(호수·종류)를 정렬해 "1|89" 같은 키로 만든다 (5자리 이상은 펀드코드·날짜로 보고 뺀다)
Public Function NumsKey(ByVal s As String) As String
    Dim t As String, i As Long, ch As String, run As String, parts() As String, n As Long, k As Long, j As Long, tmp As String, dup As Boolean
    t = NormText(s) & " "
    ReDim parts(0)
    n = 0
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If IsDigitCh(ch) Then
            run = run & ch
        ElseIf Len(run) > 0 Then
            If Len(run) < 5 Then
                Do While Len(run) > 1 And Left$(run, 1) = "0"
                    run = Mid$(run, 2)
                Loop
                dup = False
                For k = 0 To n - 1
                    If parts(k) = run Then dup = True: Exit For
                Next
                If Not dup Then
                    ReDim Preserve parts(n)
                    parts(n) = run
                    n = n + 1
                End If
            End If
            run = ""
        End If
    Next
    ' 문자열 순서로 정렬 (웹판의 sort()와 같다)
    For i = 1 To n - 1
        tmp = parts(i)
        j = i - 1
        Do While j >= 0
            If StrComp(parts(j), tmp, CMP_BINARY) <= 0 Then Exit Do
            parts(j + 1) = parts(j)
            j = j - 1
        Loop
        parts(j + 1) = tmp
    Next
    If n = 0 Then
        NumsKey = ""
    Else
        NumsKey = JoinN(parts, n, "|")
    End If
End Function

Private Function JoinN(arr() As String, ByVal n As Long, ByVal sep As String) As String
    Dim i As Long, out As String
    For i = 0 To n - 1
        If i > 0 Then out = out & sep
        out = out & arr(i)
    Next
    JoinN = out
End Function

Private Function KeyCount(ByVal key As String) As Long
    If key = "" Then
        KeyCount = 0
    Else
        KeyCount = UBound(Split(key, "|")) + 1
    End If
End Function

Private Function KeysOverlap(ByVal a As String, ByVal b As String) As Boolean
    Dim x As Variant
    For Each x In Split(a, "|")
        If InStr(1, "|" & b & "|", "|" & x & "|", CMP_BINARY) > 0 Then KeysOverlap = True: Exit Function
    Next
End Function

Private Function HasLongNum(ByVal key As String) As Boolean
    Dim x As Variant
    If key = "" Then Exit Function
    For Each x In Split(key, "|")
        If Len(x) >= 2 Then HasLongNum = True: Exit Function
    Next
End Function

' ======================================================================
'  유사도: 코사인(글자 1~3gram, TF-IDF) 60% + 글자 순서 일치(LCS) 40%
' ======================================================================
' 글자 1~3gram과 개수. 같은 문자열은 한 번만 만든다 (칸 번호를 돌려준다)
Private Function GramSlot(ByVal s As String) As Long
    Dim key As String, keys() As String, cnts() As Double, idx As Collection, n As Long, i As Long, k As Long, g As String, p As Long, slot As Long
    If gGramSlot Is Nothing Then Set gGramSlot = New Collection
    key = CK(s)
    slot = CollGetLong(gGramSlot, key, -1)
    If slot >= 0 Then
        GramSlot = slot
        Exit Function
    End If
    Set idx = New Collection
    ReDim keys(0): ReDim cnts(0)
    n = 0
    For k = 1 To 3
        For i = 1 To Len(s) - k + 1
            g = Mid$(s, i, k)
            p = CollGetLong(idx, g, -1)
            If p >= 0 Then
                cnts(p) = cnts(p) + 1
            Else
                ReDim Preserve keys(n): ReDim Preserve cnts(n)
                keys(n) = g: cnts(n) = 1
                idx.Add n, g
                n = n + 1
            End If
        Next
    Next
    slot = gGramCount
    ReDim Preserve gGramKeys(slot): ReDim Preserve gGramCnts(slot): ReDim Preserve gGramN(slot): ReDim Preserve gGramIdx(slot)
    gGramKeys(slot) = keys: gGramCnts(slot) = cnts: gGramN(slot) = n
    Set gGramIdx(slot) = idx
    gGramSlot.Add slot, key
    gGramCount = gGramCount + 1
    GramSlot = slot
End Function

Public Sub BuildIdf(coreNames() As String, ByVal n As Long)
    Dim seen As New Collection, i As Long, j As Long, slot As Long, keys As Variant, m As Long, nU As Long, p As Long
    Dim dfCnt() As Double, nd As Long
    Set gIdfIdx = New Collection
    ReDim dfCnt(0)
    nd = 0
    For i = 0 To n - 1
        If Not CollHas(seen, CK(coreNames(i))) Then
            seen.Add 1, CK(coreNames(i))
            nU = nU + 1
            slot = GramSlot(coreNames(i))
            keys = gGramKeys(slot)
            m = gGramN(slot)
            For j = 0 To m - 1
                p = CollGetLong(gIdfIdx, keys(j), -1)
                If p >= 0 Then
                    dfCnt(p) = dfCnt(p) + 1
                Else
                    ReDim Preserve dfCnt(nd)
                    dfCnt(nd) = 1
                    gIdfIdx.Add nd, keys(j)
                    nd = nd + 1
                End If
            Next
        End If
    Next
    ReDim gIdf(IIf(nd > 0, nd - 1, 0))
    For i = 0 To nd - 1
        gIdf(i) = Log((1 + nU) / (1 + dfCnt(i))) + 1
    Next
    gIdfUnseen = Log(1 + nU) + 1
    gUseIdf = True
End Sub

Private Function GramWeight(ByVal g As String) As Double
    Dim p As Long
    If Not gUseIdf Then GramWeight = 1: Exit Function
    p = CollGetLong(gIdfIdx, g, -1)
    If p >= 0 Then
        GramWeight = gIdf(p)
    Else
        GramWeight = gIdfUnseen
    End If
End Function

Public Function CosineSim(ByVal a As String, ByVal b As String) As Double
    Dim sa As Long, sb As Long, ka As Variant, ca As Variant, kb As Variant, cb As Variant, ib As Collection
    Dim i As Long, p As Long, x As Double, y As Double, dot As Double, na As Double, nb As Double, w As Double
    If Len(a) = 0 Or Len(b) = 0 Then CosineSim = 0: Exit Function
    sa = GramSlot(a): sb = GramSlot(b)
    ka = gGramKeys(sa): ca = gGramCnts(sa): kb = gGramKeys(sb): cb = gGramCnts(sb)
    Set ib = gGramIdx(sb)
    For i = 0 To gGramN(sa) - 1
        w = GramWeight(ka(i))
        x = ca(i) * w
        na = na + x * x
        p = CollGetLong(ib, ka(i), -1)
        If p >= 0 Then dot = dot + x * cb(p) * w
    Next
    For i = 0 To gGramN(sb) - 1
        y = cb(i) * GramWeight(kb(i))
        nb = nb + y * y
    Next
    If na > 0 And nb > 0 Then
        CosineSim = dot / Sqr(na * nb)
    Else
        CosineSim = 0
    End If
End Function

Public Function LcsRatio(ByVal a As String, ByVal b As String) As Double
    Dim la As Long, lb As Long, i As Long, j As Long, prev() As Long, cur() As Long
    la = Len(a): lb = Len(b)
    If la = 0 Or lb = 0 Then LcsRatio = 0: Exit Function
    ReDim prev(lb): ReDim cur(lb)
    For i = 1 To la
        cur(0) = 0
        For j = 1 To lb
            If Mid$(a, i, 1) = Mid$(b, j, 1) Then
                cur(j) = prev(j - 1) + 1
            ElseIf prev(j) >= cur(j - 1) Then
                cur(j) = prev(j)
            Else
                cur(j) = cur(j - 1)
            End If
        Next
        For j = 0 To lb
            prev(j) = cur(j)
        Next
    Next
    LcsRatio = 2 * prev(lb) / (la + lb)
End Function

' 대상 종목명과 장부 하나의 유사도 (0~1). cands = 장부 이름 후보들(본문 펀드명, 파일명)
' 펀드명 맨 앞의 사명(운용사). 펀드명은 사명으로 시작하므로 사명이 다르면 다른 펀드다.
' NormName을 거친 모양으로 적는다 (영문 사명은 별칭 처리로 한글이 된다)
Private Function ManagerList() As Variant
    ManagerList = Array("KB", "NH아문디", "NH", "신한", "삼성", "미래에셋", "한화", "키움", "이지스", "마스턴", "코람코", "블랙락", _
        "유나이티드파트너스", "교보", "현대", "한국투자", "IBK", "DB", "우리", "하나", "하나대체", "하나대체투자", "대신", "메리츠", _
        "흥국", "타임폴리오", "트러스톤", "신영", "베어링", "슈로더", "피델리티", "JP모간", "멀티에셋", "브이아이", "에셋플러스", _
        "유진", "BNK", "DGB", "하이", "KTB", "다올", "HDC", "칸서스", "캡스톤", "라살", "인베스코", "알리안츠", "맥쿼리", _
        "파인스트리트", "제이알", "케이리츠", "한국투자리얼에셋", "삼성SRA", "이지스자산", "코람코자산")
End Function

Private Function StartsWithB(ByVal s As String, ByVal p As String) As Boolean
    StartsWithB = (Len(p) > 0 And Len(s) >= Len(p) And Left$(s, Len(p)) = p)
End Function

Public Function MgrPrefix(ByVal s As String) As String
    Dim t As String, n As String, best As String, m As Variant, ch As String
    ' '(블랙락73호)'처럼 괄호로 시작하면 별칭·메모라 펀드명 앞머리로 보지 않는다
    t = TrimWS(NormText(s))
    If t <> "" Then
        ch = Left$(t, 1)
        If ch = "(" Or ch = "[" Or ch = "{" Or ch = "<" Or ch = ChrW(&HFF08) Or ch = ChrW(&H3010) Then
            MgrPrefix = ""
            Exit Function
        End If
    End If
    n = NormName(s)
    best = ""
    For Each m In ManagerList()
        If Len(CStr(m)) > Len(best) Then
            If StartsWithB(n, CStr(m)) Then best = CStr(m)
        End If
    Next
    MgrPrefix = best
End Function

' 장부 머리글 운용사명에서 '자산운용' 같은 꼬리를 뗀 사명
Public Function HeaderMgr(ByVal manager As String) As String
    Dim t As String, cut As Boolean, w As Variant
    t = NormName(manager)
    Do
        cut = False
        For Each w In Array("자산운용", "투자운용", "자산관리", "운용", "ASSETMANAGEMENT", "MANAGEMENT", "INVESTMENTTRUST", "INVESTMENT", "TRUST", "MGT", "ASSET", "COLTD", "CO", "LTD", "주식회사")
            If Len(t) >= Len(CStr(w)) Then
                If Right$(t, Len(CStr(w))) = CStr(w) Then
                    t = Left$(t, Len(t) - Len(CStr(w)))
                    cut = True
                    Exit For
                End If
            End If
        Next
    Loop While cut And t <> ""
    If Len(t) >= 2 Then
        HeaderMgr = t
    Else
        HeaderMgr = ""
    End If
End Function

Public Function ScoreName(ByVal target As String, cands As Variant, ByVal nCands As Long, ByVal manager As String, ByRef why As String) As Double
    Dim tn As String, tc As String, tnum As String, best As Double, s As Double, cs As Double, w As String
    Dim k As Long, cand As String, cc As String, cnum As String, mgr As String
    tn = NormName(target): tc = CoreName(target): tnum = NumsKey(target)
    best = 0: why = ""
    For k = 0 To nCands - 1
        cand = cands(k)
        If cand <> "" Then
            If NormName(cand) = tn Then
                why = "펀드명 완전 일치"
                ScoreName = 1
                Exit Function
            End If
            cc = CoreName(cand): cnum = NumsKey(cand)
            cs = CosineSim(tc, cc)
            s = COSINE_WEIGHT * cs + (1 - COSINE_WEIGHT) * LcsRatio(tc, cc)
            w = "코사인 " & FixedStr(cs, 2)
            If tnum <> "" And cnum <> "" Then
                ' 호수·종류가 하나라도 다르면 다른 펀드로 본다
                If tnum = cnum Then
                    s = s * 0.85 + 0.15
                    w = w & " · 호수 일치"
                Else
                    s = s * 0.35
                    If KeysOverlap(tnum, cnum) Then
                        w = w & " · 호수·종류 다름"
                    Else
                        w = w & " · 호수 다름"
                    End If
                End If
            ElseIf KeyCount(tnum) <> KeyCount(cnum) Then
                s = s * 0.8
            End If
            If s > best Then best = s: why = w
        End If
    Next
    ' 사명: 장부 쪽(펀드명·파일명 앞머리, 머리글 운용사)과 대상 종목명 앞머리를 맞춘다. 호수보다 우선
    Dim tm As String, hdr As String, lmFirst As String, lmAny As Boolean, mgrHit As Boolean, x As String
    tm = MgrPrefix(target): hdr = HeaderMgr(manager)
    mgrHit = False: lmAny = False: lmFirst = ""
    If hdr <> "" Then
        If StartsWithB(tn, hdr) Then mgrHit = True
    End If
    For k = 0 To nCands
        x = ""
        If k < nCands Then
            If cands(k) <> "" Then x = MgrPrefix(cands(k))
        ElseIf hdr <> "" Then
            x = MgrPrefix(hdr)
            If x = "" Then x = hdr
        End If
        If x <> "" Then
            If Not lmAny Then lmFirst = x
            lmAny = True
            If tm <> "" Then
                If StartsWithB(tm, x) Or StartsWithB(x, tm) Then mgrHit = True
            End If
        End If
    Next
    If mgrHit Then
        best = best + 0.12
        why = why & " · 사명 일치"
        If best > 1 Then best = 1
        ScoreName = best
        Exit Function
    End If
    If tm <> "" And lmAny Then
        why = why & " · 사명 다름(" & tm & "≠" & lmFirst & ")"
        ScoreName = best * 0.3
        Exit Function
    End If
    mgr = NormName(manager)
    If mgr <> "" And Len(tn) >= 4 Then
        If InStr(1, mgr, Left$(tn, 4), CMP_BINARY) > 0 Then
            best = best + 0.1
            why = why & " · 운용사 일치"
        ElseIf IsHangul(Mid$(tn, 1, 1)) And IsHangul(Mid$(tn, 2, 1)) And InStr(1, mgr, Left$(tn, 2), CMP_BINARY) > 0 Then
            best = best + 0.04
        End If
    End If
    If best > 1 Then best = 1
    ScoreName = best
End Function

Private Function FixedStr(ByVal x As Double, ByVal digits As Long) As String
    ' 소수점 자리 고정 (시스템 지역 설정과 무관하게 '.' 사용)
    Dim m As Double, v As Double, s As String, neg As Boolean
    m = 10 ^ digits
    neg = (x < 0)
    v = Int(Abs(x) * m + 0.5)
    s = CStr(v)
    Do While Len(s) <= digits
        s = "0" & s
    Loop
    s = Left$(s, Len(s) - digits) & "." & Right$(s, digits)
    If neg And v <> 0 Then s = "-" & s
    FixedStr = s
End Function

' 대상 종목명들(names, 중복 없음)과 장부들(lds)을 1:1로 짝짓는다.
' forced(t): -2 = 지정 없음, -1 = 매칭하지 않음, 0 이상 = 그 장부
Public Sub MatchNames(names() As String, ByVal nT As Long, lds() As Variant, ByVal nL As Long, forced() As Long)
    Dim corpus() As String, nc As Long, t As Long, l As Long, k As Long, why As String, cands As Variant
    Dim assigned() As Boolean, used() As Boolean
    ' IDF는 대상 종목명 + 장부 이름 전체로 만든다
    ReDim corpus(0): nc = 0
    For t = 0 To nT - 1
        ReDim Preserve corpus(nc): corpus(nc) = CoreName(names(t)): nc = nc + 1
    Next
    For l = 0 To nL - 1
        If IsUsable(lds(l)) Then
            cands = lds(l)(LD_CANDS)
            For k = 0 To lds(l)(LD_NCANDS) - 1
                If cands(k) <> "" Then ReDim Preserve corpus(nc): corpus(nc) = CoreName(cands(k)): nc = nc + 1
            Next
        End If
    Next
    BuildIdf corpus, nc
    ReDim mScores(IIf(nT > 0, nT - 1, 0), IIf(nL > 0, nL - 1, 0))
    ReDim mWhys(IIf(nT > 0, nT - 1, 0), IIf(nL > 0, nL - 1, 0))
    ReDim mAssign(IIf(nT > 0, nT - 1, 0))
    For t = 0 To nT - 1
        For l = 0 To nL - 1
            ' 읽기 실패·결산전(결산후 있음) 장부는 후보가 아니다
            If IsUsable(lds(l)) Then
                mScores(t, l) = ScoreName(names(t), lds(l)(LD_CANDS), lds(l)(LD_NCANDS), lds(l)(LD_MANAGER), why)
            Else
                mScores(t, l) = 0: why = ""
            End If
            mWhys(t, l) = why
        Next
    Next
    ' 호수 유일 일치: 장부 이름이 깨져도 두 자리 이상 호수가 대상 하나와만 맞으면 검토 후보로 올림
    Dim cn As String, hits As Long, hitT As Long, src As String
    For l = 0 To nL - 1
        src = lds(l)(LD_NAME)
        If src = "" Then src = lds(l)(LD_TITLE)
        If src = "" And lds(l)(LD_NCANDS) > 0 Then src = lds(l)(LD_CANDS)(0)
        cn = NumsKey(src)
        If Not IsUsable(lds(l)) Then cn = ""
        If HasLongNum(cn) Then
            hits = 0
            For t = 0 To nT - 1
                If NumsKey(names(t)) = cn Then hits = hits + 1: hitT = t
            Next
            If hits = 1 Then
                If mScores(hitT, l) < 0.62 And InStr(1, mWhys(hitT, l), "사명 다름", CMP_BINARY) = 0 Then mScores(hitT, l) = 0.62: mWhys(hitT, l) = "호수만 유일하게 일치"
            End If
        End If
    Next
    ' 배정: 지정된 매핑 먼저, 나머지는 점수 높은 쌍부터
    ReDim assigned(IIf(nT > 0, nT - 1, 0)): ReDim used(IIf(nL > 0, nL - 1, 0))
    For t = 0 To nT - 1
        mAssign(t) = -1
        If forced(t) <> -2 Then
            assigned(t) = True
            mAssign(t) = forced(t)
            If forced(t) >= 0 Then used(forced(t)) = True
        End If
    Next
    Dim pT() As Long, pL() As Long, pS() As Double, np As Long, i As Long, j As Long, xt As Long, xl As Long, xs As Double
    ReDim pT(0): ReDim pL(0): ReDim pS(0)
    np = 0
    For t = 0 To nT - 1
        If Not assigned(t) Then
            For l = 0 To nL - 1
                If mScores(t, l) >= REVIEW_SCORE Then
                    ReDim Preserve pT(np): ReDim Preserve pL(np): ReDim Preserve pS(np)
                    pT(np) = t: pL(np) = l: pS(np) = mScores(t, l)
                    np = np + 1
                End If
            Next
        End If
    Next
    ' 안정 정렬(점수 내림차순)
    For i = 1 To np - 1
        xt = pT(i): xl = pL(i): xs = pS(i)
        j = i - 1
        Do While j >= 0
            If Not (pS(j) < xs) Then Exit Do
            pT(j + 1) = pT(j): pL(j + 1) = pL(j): pS(j + 1) = pS(j)
            j = j - 1
        Loop
        pT(j + 1) = xt: pL(j + 1) = xl: pS(j + 1) = xs
    Next
    For i = 0 To np - 1
        If Not assigned(pT(i)) And Not used(pL(i)) Then
            assigned(pT(i)) = True
            used(pL(i)) = True
            mAssign(pT(i)) = pL(i)
        End If
    Next
End Sub

' 배정된 장부의 이름 점수만으로 본 상태: auto / review
Public Function MatchStatus(ByVal t As Long, ByVal l As Long, ByVal nL As Long) As String
    Dim secondBest As Double, k As Long, e As Double, margin As Double
    e = mScores(t, l)
    secondBest = 0
    For k = 0 To nL - 1
        If k <> l Then If mScores(t, k) > secondBest Then secondBest = mScores(t, k)
    Next
    margin = e - secondBest
    If e >= AUTO_SCORE And margin >= AUTO_MARGIN Then
        MatchStatus = "auto"
    ElseIf e >= AUTO_CLEAR_SCORE And margin >= AUTO_CLEAR_MARGIN Then
        MatchStatus = "auto"
    Else
        MatchStatus = "review"
    End If
End Function

' ======================================================================
'  숫자·날짜
' ======================================================================
Public Function ToNum(v As Variant, ByRef ok As Boolean) As Double
    Dim t As String, neg As Boolean, i As Long, ch As String, body As String, intPart As String, hasDigit As Boolean
    Dim dotPos As Long, fracPart As String
    ok = False
    Select Case VarType(v)
        Case VT_DOUBLE, VT_SINGLE, VT_INTEGER, VT_LONG, VT_CURRENCY, VT_DECIMAL, VT_BYTE
            ToNum = CDbl(v): ok = True: Exit Function
        Case VT_STRING
        Case Else
            Exit Function
    End Select
    t = RemoveWS(TrimWS(CStr(v)))
    For i = 1 To Len(t)
        If IsDigitCh(Mid$(t, i, 1)) Then hasDigit = True: Exit For
    Next
    If Not hasDigit Then Exit Function
    If Len(t) >= 2 And Left$(t, 1) = "(" And Right$(t, 1) = ")" Then neg = True: t = Mid$(t, 2, Len(t) - 2)
    body = t
    If Left$(body, 1) = "-" Then body = Mid$(body, 2)
    dotPos = InStr(body, ".")
    If dotPos > 0 Then
        intPart = Left$(body, dotPos - 1)
        fracPart = Mid$(body, dotPos + 1)
        If Len(fracPart) = 0 Then Exit Function
        For i = 1 To Len(fracPart)
            If Not IsDigitCh(Mid$(fracPart, i, 1)) Then Exit Function
        Next
    Else
        intPart = body
    End If
    If Len(intPart) = 0 Then Exit Function
    For i = 1 To Len(intPart)
        ch = Mid$(intPart, i, 1)
        If Not (IsDigitCh(ch) Or ch = ",") Then Exit Function
    Next
    If InStr(intPart, ",") > 0 Then
        ' 쉼표는 세 자리마다만 허용
        Dim groups As Variant, g As Long
        groups = Split(intPart, ",")
        If Len(groups(0)) < 1 Or Len(groups(0)) > 3 Then Exit Function
        For g = 1 To UBound(groups)
            If Len(groups(g)) <> 3 Then Exit Function
        Next
    End If
    ToNum = Val(Replace(t, ",", ""))
    If neg Then ToNum = -ToNum
    ok = True
End Function

Private Function ValidYMD(ByVal y As Long, ByVal m As Long, ByVal d As Long) As Boolean
    ValidYMD = (y >= 2000 And y <= 2099 And m >= 1 And m <= 12 And d >= 1 And d <= 31)
End Function

Private Function Iso(ByVal y As Long, ByVal m As Long, ByVal d As Long) As String
    Iso = CStr(y) & "-" & Right$("0" & CStr(m), 2) & "-" & Right$("0" & CStr(d), 2)
End Function

Private Function SkipWS(ByVal t As String, ByVal p As Long) As Long
    Do While p <= Len(t)
        If Not IsWS(Mid$(t, p, 1)) Then Exit Do
        p = p + 1
    Loop
    SkipWS = p
End Function

Private Function DigitRunLen(ByVal t As String, ByVal p As Long, ByVal maxN As Long) As Long
    Dim n As Long
    Do While p + n <= Len(t) And n < maxN
        If Not IsDigitCh(Mid$(t, p + n, 1)) Then Exit Do
        n = n + 1
    Loop
    DigitRunLen = n
End Function

' "2026.09.30", "2026-9-30", "2026년 9월 30일", "20260930" → "2026-09-30" (여러 개면 마지막)
Public Function DateFromText(v As Variant) As String
    Dim t As String, p As Long, q As Long, y As Long, mm As Long, dd As Long, ml As Long, dl As Long, found As String, endPos As Long, k As Long
    If VarType(v) <> VT_STRING Then DateFromText = "": Exit Function
    t = NormText(CStr(v))
    p = 1
    Do While p <= Len(t) - 3
        If Mid$(t, p, 2) = "20" And DigitRunLen(t, p, 4) = 4 Then
            y = CLng(Mid$(t, p, 4))
            q = SkipWS(t, p + 4)
            If q <= Len(t) And InStr(".-/년", Mid$(t, q, 1)) > 0 Then
                q = SkipWS(t, q + 1)
                ml = DigitRunLen(t, q, 2)
                endPos = 0
                For k = ml To 1 Step -1
                    Dim q2 As Long
                    q2 = SkipWS(t, q + k)
                    If q2 <= Len(t) Then
                        If InStr(".-/월", Mid$(t, q2, 1)) > 0 Then
                            mm = CLng(Mid$(t, q, k))
                            q2 = SkipWS(t, q2 + 1)
                            dl = DigitRunLen(t, q2, 2)
                            If dl > 0 Then
                                dd = CLng(Mid$(t, q2, dl))
                                endPos = q2 + dl
                                Exit For
                            End If
                        End If
                    End If
                Next
                If endPos > 0 Then
                    If ValidYMD(y, mm, dd) Then found = Iso(y, mm, dd)
                    p = endPos
                    GoTo ContinueP
                End If
            End If
        End If
        p = p + 1
ContinueP:
    Loop
    If found <> "" Then DateFromText = found: Exit Function
    ' 붙어 있는 8자리 20YYMMDD
    Dim i As Long, rl As Long
    i = 1
    Do While i <= Len(t)
        rl = DigitRunLen(t, i, 999)
        If rl > 0 Then
            If rl = 8 And Mid$(t, i, 2) = "20" Then
                y = CLng(Mid$(t, i, 4)): mm = CLng(Mid$(t, i + 4, 2)): dd = CLng(Mid$(t, i + 6, 2))
                If ValidYMD(y, mm, dd) Then found = Iso(y, mm, dd)
            End If
            i = i + rl
        Else
            i = i + 1
        End If
    Loop
    DateFromText = found
End Function

Public Function DateFromFileName(ByVal s As String) As String
    Dim d As String, t As String, i As Long, rl As Long, yy As Long, mm As Long, dd As Long
    d = DateFromText(s)
    If d <> "" Then DateFromFileName = d: Exit Function
    t = NormText(s)
    i = 1
    Do While i <= Len(t)
        rl = DigitRunLen(t, i, 999)
        If rl > 0 Then
            If rl = 6 Then
                yy = CLng(Mid$(t, i, 2)): mm = CLng(Mid$(t, i + 2, 2)): dd = CLng(Mid$(t, i + 4, 2))
                If yy >= 15 And yy <= 45 And ValidYMD(2000 + yy, mm, dd) Then DateFromFileName = Iso(2000 + yy, mm, dd)
                Exit Function
            End If
            i = i + rl
        Else
            i = i + 1
        End If
    Loop
End Function

Public Function SerialToISO(v As Variant) As String
    Dim n As Double, d As Date, ok As Boolean
    Select Case VarType(v)
        Case VT_DOUBLE, VT_SINGLE, VT_INTEGER, VT_LONG, VT_CURRENCY, VT_DECIMAL
        Case Else
            Exit Function
    End Select
    n = CDbl(v)
    If n < 30000 Or n > 80000 Then Exit Function
    d = DateSerial(1899, 12, 30) + Int(n)
    SerialToISO = Iso(Year(d), Month(d), Day(d))
End Function

' 파일명에서 펀드명만 남긴다: 확장자, 날짜, '(1)' 같은 사본 표시, '기준가격대장' 같은 말 제거
Public Function CleanFileName(ByVal fileName As String) As String
    Dim t As String, p As Long, q As Long, suf As Long, i As Long, rl As Long, out As String, w As Variant
    t = NormText(fileName)
    p = InStrRev(t, ".")
    If p > 0 And p < Len(t) Then t = Left$(t, p - 1)
    t = CompactDates(t)
    ' 앞의 '02.' '3)' 같은 정리 번호는 호수가 아니다
    t = DropLeadingOrdinal(t)
    ' 날짜 뒤의 '_1', '(1)' 같은 사본 표시
    q = Len(t)
    Do While q > 0
        If Not IsWS(Mid$(t, q, 1)) Then Exit Do
        q = q - 1
    Loop
    suf = 0
    If q >= 3 Then
        If Mid$(t, q, 1) = ")" Then
            i = q - 1
            Do While i > 0
                If Not IsDigitCh(Mid$(t, i, 1)) Then Exit Do
                i = i - 1
            Loop
            If i > 0 And i < q - 1 Then If Mid$(t, i, 1) = "(" Then suf = i
        ElseIf IsDigitCh(Mid$(t, q, 1)) And Mid$(t, q - 1, 1) = "_" Then
            suf = q - 1
        End If
    End If
    If suf > 0 Then
        i = suf - 1
        Do While i > 0
            If Not (IsWS(Mid$(t, i, 1)) Or Mid$(t, i, 1) = "_" Or Mid$(t, i, 1) = "-") Then Exit Do
            i = i - 1
        Loop
        rl = 0
        Do While i - rl > 0
            If Not IsDigitCh(Mid$(t, i - rl, 1)) Then Exit Do
            rl = rl + 1
        Loop
        If rl >= 6 Then t = Left$(t, i)
    End If
    ' 끝의 '(1)'
    q = Len(t)
    Do While q > 0
        If Not IsWS(Mid$(t, q, 1)) Then Exit Do
        q = q - 1
    Loop
    If q >= 3 Then
        If Mid$(t, q, 1) = ")" Then
            i = q - 1
            Do While i > 0
                If Not IsDigitCh(Mid$(t, i, 1)) Then Exit Do
                i = i - 1
            Loop
            If i > 0 And i < q - 1 Then
                If Mid$(t, i, 1) = "(" Then
                    i = i - 1
                    Do While i > 0
                        If Not IsWS(Mid$(t, i, 1)) Then Exit Do
                        i = i - 1
                    Loop
                    t = Left$(t, i)
                End If
            End If
        End If
    End If
    ' 날짜 숫자(20YYMMDD, 6자리) → 공백
    out = ""
    i = 1
    Do While i <= Len(t)
        rl = DigitRunLen(t, i, 999)
        If rl > 0 Then
            Dim run As String, k As Long, piece As String
            run = Mid$(t, i, rl)
            If rl = 6 Then
                out = out & " "
            Else
                k = 1
                piece = ""
                Do While k <= rl
                    If Mid$(run, k, 2) = "20" And rl - k + 1 >= 8 Then
                        out = out & piece & " "
                        piece = ""
                        k = k + 8
                    Else
                        piece = piece & Mid$(run, k, 1)
                        k = k + 1
                    End If
                Loop
                out = out & piece
            End If
            i = i + rl
        Else
            out = out & Mid$(t, i, 1)
            i = i + 1
        End If
    Loop
    t = DropFundCodes(out)
    ' '2026년 2분기', '2Q' 같은 분기·연도 표기
    t = RemoveQuarter(t)
    For Each w In Array("(결산전)", "(결산 전)", "(결산후)", "(결산 후)", "(결산반영)", "[결산전]", "[결산 전]", "[결산후]", "[결산 후]", "[결산반영]")
        t = Replace(t, CStr(w), " ")
    Next
    For Each w In Array("기준가격대장", "기준가대장", "신탁재산명세부", "통합명세부", "재산명세부", "명세부", "총계정원장", "계정원장", "원장", "마감자료", "조회", "사본", "복사본", "결산전", "결산 전", "결산후", "결산 후", "결산반영")
        t = Replace(t, CStr(w), " ")
    Next
    t = Replace(t, "_", " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CleanFileName = TrimWS(t)
End Function

Private Function DropLeadingOrdinal(ByVal t As String) As String
    Dim p As Long, rl As Long, q As Long
    p = SkipWS(t, 1)
    rl = DigitRunLen(t, p, 999)
    If rl >= 1 And rl <= 3 Then
        q = SkipWS(t, p + rl)
        If q <= Len(t) Then
            If Mid$(t, q, 1) = "." Or Mid$(t, q, 1) = ")" Then
                DropLeadingOrdinal = Mid$(t, SkipWS(t, q + 1))
                Exit Function
            End If
        End If
    End If
    DropLeadingOrdinal = t
End Function

' 영문자에 붙지 않은 5자리 이상 숫자(펀드코드) → 공백
Private Function DropFundCodes(ByVal t As String) As String
    Dim i As Long, rl As Long, out As String, okL As Boolean, okR As Boolean
    i = 1
    Do While i <= Len(t)
        rl = DigitRunLen(t, i, 999)
        If rl > 0 Then
            okL = True: okR = True
            If i > 1 Then okL = Not IsAsciiAlpha(Mid$(t, i - 1, 1))
            If i + rl <= Len(t) Then okR = Not IsAsciiAlpha(Mid$(t, i + rl, 1))
            If rl >= 5 And okL And okR Then
                out = out & " "
            Else
                out = out & Mid$(t, i, rl)
            End If
            i = i + rl
        Else
            out = out & Mid$(t, i, 1)
            i = i + 1
        End If
    Loop
    DropFundCodes = out
End Function

Private Function IsAsciiAlpha(ByVal ch As String) As Boolean
    Dim c As Long
    c = CodeOf(ch)
    IsAsciiAlpha = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

Private Function IsDateSep(ByVal ch As String, ByVal withUnderscore As Boolean) As Boolean
    IsDateSep = (ch = "." Or ch = "-" Or (withUnderscore And (ch = "_" Or ch = "/")))
End Function

' '2026.09.30', '2026-9-30', '2026_09_30' → '20260930', '26.09.30' → '260930' (뒤에서 날짜 숫자로 함께 지운다)
Private Function CompactDates(ByVal t As String) As String
    Dim i As Long, out As String, q As Long, q2 As Long, ml As Long, dl As Long, prevDigit As Boolean, done As Boolean
    i = 1
    Do While i <= Len(t)
        done = False
        prevDigit = False
        If i > 1 Then prevDigit = IsDigitCh(Mid$(t, i - 1, 1))
        If Not prevDigit And Mid$(t, i, 2) = "20" And DigitRunLen(t, i, 4) = 4 Then
            q = SkipWS(t, i + 4)
            If q <= Len(t) Then
                If IsDateSep(Mid$(t, q, 1), True) Then
                    q = SkipWS(t, q + 1)
                    ml = DigitRunLen(t, q, 3)
                    If ml >= 1 And ml <= 2 Then
                        q2 = SkipWS(t, q + ml)
                        If q2 <= Len(t) Then
                            If IsDateSep(Mid$(t, q2, 1), True) Then
                                q2 = SkipWS(t, q2 + 1)
                                dl = DigitRunLen(t, q2, 3)
                                If dl >= 1 And dl <= 2 Then
                                    out = out & Mid$(t, i, 4) & Right$("0" & Mid$(t, q, ml), 2) & Right$("0" & Mid$(t, q2, dl), 2)
                                    i = q2 + dl
                                    done = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If Not done And Not prevDigit And i + 7 <= Len(t) Then
            ' 26.09.30 / 26-09-30 (각 두 자리, 구분은 '.' 또는 '-')
            If DigitRunLen(t, i, 3) = 2 Then
                If IsDateSep(Mid$(t, i + 2, 1), False) And DigitRunLen(t, i + 3, 3) = 2 Then
                    If IsDateSep(Mid$(t, i + 5, 1), False) And DigitRunLen(t, i + 6, 3) = 2 Then
                        out = out & Mid$(t, i, 2) & Mid$(t, i + 3, 2) & Mid$(t, i + 6, 2)
                        i = i + 8
                        done = True
                    End If
                End If
            End If
        End If
        If Not done Then
            out = out & Mid$(t, i, 1)
            i = i + 1
        End If
    Loop
    CompactDates = out
End Function

Private Function RemoveQuarter(ByVal t As String) As String
    Dim i As Long, out As String, ch As String, q As Long, s As Long, j As Long, ok As Boolean
    i = 1
    Do While i <= Len(t)
        ch = Mid$(t, i, 1)
        ok = False
        If ch >= "1" And ch <= "4" Then
            q = SkipWS(t, i + 1)
            If q <= Len(t) Then
                If UCase$(Mid$(t, q, 1)) = "Q" Then
                    ok = True: q = q + 1
                ElseIf Mid$(t, q, 2) = "분기" Then
                    ok = True: q = q + 2
                End If
            End If
        End If
        If ok Then
            ' 앞쪽의 '2026년 ' 또는 '2026 '도 함께 지운다
            s = Len(out)
            j = s
            Do While j > 0
                If Not IsWS(Mid$(out, j, 1)) Then Exit Do
                j = j - 1
            Loop
            If j > 0 Then If Mid$(out, j, 1) = "년" Then j = j - 1
            Do While j > 0
                If Not IsWS(Mid$(out, j, 1)) Then Exit Do
                j = j - 1
            Loop
            If j >= 4 Then
                If Mid$(out, j - 3, 2) = "20" And DigitRunLen(out, j - 3, 4) = 4 Then out = Left$(out, j - 4)
            End If
            out = out & " "
            i = q
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    ' 남은 '2026년'
    t = out: out = ""
    i = 1
    Do While i <= Len(t)
        If Mid$(t, i, 2) = "20" And DigitRunLen(t, i, 4) = 4 Then
            q = SkipWS(t, i + 4)
            If q <= Len(t) Then
                If Mid$(t, q, 1) = "년" Then out = out & " ": i = q + 1: GoTo NextQ
            End If
        End If
        out = out & Mid$(t, i, 1)
        i = i + 1
NextQ:
    Loop
    RemoveQuarter = out
End Function

' ======================================================================
'  기준가격대장에서 계정 찾기 (시트 값 2차원 배열 → 총자산)
' ======================================================================
Private Function FieldNames(ByVal k As Long) As Variant
    Select Case k
        Case 0: FieldNames = Array("신탁재산총자산", "신탁자산총자산", "신탁재산총자산(감사보수기준)", "신탁자산총자산(감사보수기준)", "총자산", "총자산액", "자산총액", "펀드총자산")
        Case 1: FieldNames = Array("신탁재산순자산총액", "신탁자산순자산총액", "순자산총액", "신탁재산순자산", "순자산")
        Case 2: FieldNames = Array("신탁자산총부채", "신탁재산총부채", "부채", "부채합계", "부채총계", "부채총액", "총부채")
        Case 3: FieldNames = Array("자산", "자산합계", "자산총계")
        Case 4: FieldNames = Array("수익증권기준가격", "수익증권기준가액", "기준가격", "기준가")
        Case Else: FieldNames = Array("수익증권잔존좌수", "잔존좌수", "수익증권좌수", "발행좌수", "총좌수")
    End Select
End Function

' 2차원 배열을 행별 셀 목록으로: rowsC(i) = 열 번호들, rowsV(i) = 값들(숫자는 Double, 글자는 String)
Public Sub ToRows(arr As Variant, ByRef rowsC() As Variant, ByRef rowsV() As Variant, ByRef rowsN() As Long, ByRef nRows As Long)
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long, i As Long, j As Long, v As Variant, n As Long, ok As Boolean, d As Double
    Dim cc() As Long, vv() As Variant
    If Not IsArray(arr) Then
        nRows = 0
        Exit Sub
    End If
    r1 = LBound(arr, 1): r2 = UBound(arr, 1): c1 = LBound(arr, 2): c2 = UBound(arr, 2)
    nRows = r2 - r1 + 1
    ReDim rowsC(1 To nRows): ReDim rowsV(1 To nRows): ReDim rowsN(1 To nRows)
    For i = r1 To r2
        ReDim cc(0): ReDim vv(0)
        n = 0
        For j = c1 To c2
            v = arr(i, j)
            If IsError(v) Or IsEmpty(v) Or IsNull(v) Then GoTo NextCell
            Select Case VarType(v)
                Case VT_STRING
                    If Len(v) = 0 Then GoTo NextCell
                    d = ToNum(v, ok)
                    ReDim Preserve cc(n): ReDim Preserve vv(n)
                    cc(n) = j - c1 + 1
                    If ok Then
                        vv(n) = d
                    Else
                        vv(n) = CStr(v)
                    End If
                    n = n + 1
                Case VT_DOUBLE, VT_SINGLE, VT_INTEGER, VT_LONG, VT_CURRENCY, VT_DECIMAL, VT_BYTE, VT_DATE
                    ReDim Preserve cc(n): ReDim Preserve vv(n)
                    cc(n) = j - c1 + 1: vv(n) = CDbl(v)
                    n = n + 1
            End Select
NextCell:
        Next
        ' 배열은 매번 새로 복사해 넣는다 (같은 배열 변수를 다시 쓰므로)
        rowsC(i - r1 + 1) = CopyLongs(cc, n): rowsV(i - r1 + 1) = CopyVariants(vv, n): rowsN(i - r1 + 1) = n
    Next
End Sub

Private Function CopyLongs(src() As Long, ByVal n As Long) As Variant
    Dim out() As Long, k As Long
    ReDim out(IIf(n > 0, n - 1, 0))
    For k = 0 To n - 1
        out(k) = src(k)
    Next
    CopyLongs = out
End Function

Private Function CopyVariants(src() As Variant, ByVal n As Long) As Variant
    Dim out() As Variant, k As Long
    ReDim out(IIf(n > 0, n - 1, 0))
    For k = 0 To n - 1
        out(k) = src(k)
    Next
    CopyVariants = out
End Function

Private Function IsStrCell(v As Variant) As Boolean
    IsStrCell = (VarType(v) = VT_STRING)
End Function

' '당일잔액'이 있는 차변/대변 머리행인지
Private Function HeaderAt(cc As Variant, vv As Variant, ByVal n As Long, ByRef hCols As Variant, ByRef hLeft As Long, ByRef hRight As Long) As Boolean
    Dim k As Long, joined As String, lab As String, found As Boolean, cols() As Long
    For k = 0 To n - 1
        If IsStrCell(vv(k)) Then joined = joined & NormLabel(vv(k))
    Next
    ' '당일잔액' 또는 '금일잔액'
    If (InStr(joined, "당일") = 0 And InStr(joined, "금일") = 0) Or (InStr(joined, "잔액") = 0 And InStr(joined, "잔고") = 0) Then Exit Function
    For k = 0 To n - 1
        If IsStrCell(vv(k)) Then
            lab = NormLabel(vv(k))
            If (InStr(lab, "당일") > 0 Or InStr(lab, "금일") > 0) And InStr(lab, "증감") = 0 Then
                If Not found Then
                    hLeft = cc(k): hRight = cc(k): found = True
                Else
                    If cc(k) < hLeft Then hLeft = cc(k)
                    If cc(k) > hRight Then hRight = cc(k)
                End If
            End If
        End If
    Next
    If Not found Then Exit Function
    ReDim cols(IIf(n > 0, n - 1, 0))
    For k = 0 To n - 1
        cols(k) = cc(k)
    Next
    hCols = cols
    HeaderAt = True
End Function

' 계정 값 찾기. fv/fpri/flabel(0~5) = 총자산·순자산·부채·자산·기준가·좌수
Public Sub ExtractFields(rowsC() As Variant, rowsV() As Variant, rowsN() As Long, ByVal nRows As Long, _
                         ByRef fv() As Double, ByRef fpri() As Long, ByRef flabel() As String, ByRef headerRow As Long, ByRef hasHeader As Boolean)
    Dim i As Long, k As Long, f As Long, n As Long, cc As Variant, vv As Variant
    Dim hCols As Variant, hLeft As Long, hRight As Long, inHeader As Boolean, tmpCols As Variant, tl As Long, tr As Long
    Dim labels() As String, nLab As Long, joined As String, nTexts As Long, names As Variant, pri As Long, x As Long
    Dim labelCol As Long, leftV As Double, rightV As Double, firstV As Double, leftSet As Boolean, rightSet As Boolean, firstSet As Boolean
    Dim near As Long, d As Double, dd As Double, h As Long, v As Double, vSet As Boolean
    ReDim fv(5): ReDim fpri(5): ReDim flabel(5)
    For f = 0 To 5
        fpri(f) = -1
    Next
    headerRow = 0: hasHeader = False: inHeader = False
    For i = 1 To nRows
        n = rowsN(i)
        If n = 0 Then GoTo NextRow
        cc = rowsC(i): vv = rowsV(i)
        If HeaderAt(cc, vv, n, tmpCols, tl, tr) Then
            hCols = tmpCols: hLeft = tl: hRight = tr: inHeader = True
            hasHeader = True
            If headerRow = 0 Then headerRow = i
            GoTo NextRow
        End If
        ReDim labels(0): nLab = 0: joined = "": nTexts = 0: labelCol = -1
        For k = 0 To n - 1
            If IsStrCell(vv(k)) Then
                ReDim Preserve labels(nLab): labels(nLab) = NormLabel(vv(k)): nLab = nLab + 1
                joined = joined & NormLabel(vv(k))
                nTexts = nTexts + 1
                If cc(k) > labelCol Then labelCol = cc(k)
            End If
        Next
        If nTexts = 0 Then GoTo NextRow
        If nTexts > 1 Then ReDim Preserve labels(nLab): labels(nLab) = joined: nLab = nLab + 1
        For f = 0 To 5
            names = FieldNames(f)
            pri = -1
            For x = 0 To UBound(names)
                For k = 0 To nLab - 1
                    If labels(k) = names(x) Then pri = x: Exit For
                Next
                If pri >= 0 Then Exit For
            Next
            If pri < 0 Then GoTo NextField
            If fpri(f) >= 0 And fpri(f) <= pri Then GoTo NextField
            leftSet = False: rightSet = False: firstSet = False
            For k = 0 To n - 1
                If Not IsStrCell(vv(k)) Then
                    If inHeader Then
                        near = hCols(0): d = 1E+30
                        For h = 0 To UBound(hCols)
                            dd = Abs(hCols(h) - cc(k))
                            If dd < d Then d = dd: near = hCols(h)
                        Next
                        If near = hLeft And Not leftSet Then
                            leftV = vv(k): leftSet = True
                        ElseIf near = hRight And Not rightSet Then
                            rightV = vv(k): rightSet = True
                        End If
                    ElseIf cc(k) > labelCol And Not firstSet Then
                        firstV = vv(k): firstSet = True
                    End If
                End If
            Next
            vSet = False
            If inHeader Then
                If f = 2 Then
                    If rightSet And rightV <> 0 Then
                        v = rightV: vSet = True
                    ElseIf leftSet And leftV <> 0 Then
                        v = leftV: vSet = True
                    ElseIf rightSet Then
                        v = rightV: vSet = True
                    ElseIf leftSet Then
                        v = leftV: vSet = True
                    End If
                Else
                    If leftSet And leftV <> 0 Then
                        v = leftV: vSet = True
                    ElseIf rightSet And rightV <> 0 Then
                        v = rightV: vSet = True
                    ElseIf leftSet Then
                        v = leftV: vSet = True
                    ElseIf rightSet Then
                        v = rightV: vSet = True
                    End If
                End If
            ElseIf firstSet Then
                v = firstV: vSet = True
            End If
            If vSet Then
                fpri(f) = pri: fv(f) = v: flabel(f) = names(pri)
            End If
NextField:
        Next
NextRow:
    Next
End Sub

' 머리글에서 펀드명·펀드코드·기준일·운용사·단위 찾기. 반환: Array(name, code, date, manager, unit)
Public Function ExtractMeta(rowsC() As Variant, rowsV() As Variant, rowsN() As Long, ByVal nRows As Long, ByVal lim As Long) As Variant
    Dim mName As String, mCode As String, mDate As String, mMgr As String, mUnit As String
    Dim i As Long, k As Long, k2 As Long, n As Long, cc As Variant, vv As Variant, raw As String, tl As String, rest As String
    Dim d As String, up As String, p As Long, q As Long, w As Variant
    If lim > nRows Then lim = nRows
    For i = 1 To lim
        n = rowsN(i)
        If n = 0 Then GoTo NextRowM
        cc = rowsC(i): vv = rowsV(i)
        For k = 0 To n - 1
            If Not IsStrCell(vv(k)) Then GoTo NextCellM
            raw = TrimWS(NormText(vv(k)))
            tl = RemoveWS(raw)
            If Left$(tl, 3) = "펀드명" Or tl = "펀드:" Then
                If tl = "펀드:" Then
                    rest = ""
                Else
                    rest = TrimWS(StripFundLabel(raw))
                End If
                If HasLetters(rest) Then
                    If mName = "" Then mName = rest
                Else
                    For k2 = k + 1 To n - 1
                        If Not IsStrCell(vv(k2)) Then
                            If mCode = "" Then mCode = NumCode(vv(k2))
                        ElseIf HasLetters(vv(k2)) Then
                            If mName = "" Then mName = TrimWS(NormText(vv(k2)))
                            Exit For
                        End If
                    Next
                End If
                GoTo NextCellM
            End If
            ' 펀드: 12345 / 펀드코드: ABC-1
            If Left$(tl, 2) = "펀드" Then
                q = 3
                If Mid$(tl, 3, 2) = "코드" Then q = 5
                If Mid$(tl, q, 1) = ":" And Len(tl) > q Then
                    rest = Mid$(tl, q + 1)
                    If IsCodeChars(rest) Then
                        If mCode = "" Then mCode = rest
                        GoTo NextCellM
                    End If
                End If
            End If
            If InStr(tl, "기준산출일자") > 0 Or InStr(tl, "기준가산출일자") > 0 Or InStr(tl, "기준일") > 0 Or InStr(tl, "작업일자") > 0 _
               Or InStr(tl, "기준가격일") > 0 Or InStr(tl, "조회일") > 0 Or InStr(tl, "평가일") > 0 Then
                d = DateFromText(raw)
                If d = "" Then
                    For k2 = k + 1 To n - 1
                        If IsStrCell(vv(k2)) Then
                            d = DateFromText(vv(k2))
                        Else
                            d = SerialToISO(vv(k2))
                        End If
                        If d <> "" Then Exit For
                    Next
                End If
                If d <> "" And mDate = "" Then mDate = d
            End If
            If mMgr = "" Then
                up = UCase$(tl)
                If InStr(up, "자산운용") > 0 Or InStr(up, "투자운용") > 0 Or InStr(up, "ASSETMANAGEMENT") > 0 Or InStr(up, "INVESTMENT") > 0 Then mMgr = raw
            End If
            p = InStr(tl, "단위")
            If p > 0 Then
                q = p + 2
                If Mid$(tl, q, 1) = ":" Then q = q + 1
                If Mid$(tl, q, 1) = "(" Then q = q + 1
                For Each w In Array("천원", "백만원", "억원", "원")
                    If Mid$(tl, q, Len(w)) = w Then mUnit = w: Exit For
                Next
            End If
NextCellM:
        Next
NextRowM:
    Next
    If mDate = "" Then
        For i = 1 To lim
            n = rowsN(i)
            If n > 0 Then
                vv = rowsV(i)
                For k = 0 To n - 1
                    If IsStrCell(vv(k)) Then
                        d = DateFromText(vv(k))
                        If d <> "" Then mDate = d: Exit For
                    End If
                Next
                If mDate <> "" Then Exit For
            End If
        Next
    End If
    ' 제목 아래에 엑셀 날짜 숫자(예: 46295)만 있는 행
    If mDate = "" Then
        For i = 1 To lim
            If rowsN(i) = 1 Then
                vv = rowsV(i)
                If Not IsStrCell(vv(0)) Then
                    d = SerialToISO(vv(0))
                    If d <> "" Then mDate = d: Exit For
                End If
            End If
        Next
    End If
    ' '[10101] 펀드명'처럼 앞에 붙은 펀드코드는 떼어 코드로
    If Len(mName) > 2 Then
        If Left$(mName, 1) = "[" Or Left$(mName, 1) = "(" Then
            If Left$(mName, 1) = "[" Then
                q = InStr(2, mName, "]")
            Else
                q = InStr(2, mName, ")")
            End If
            If q > 2 Then
                rest = TrimWS(Mid$(mName, 2, q - 2))
                If IsCodeChars(rest) And TrimWS(Mid$(mName, q + 1)) <> "" Then
                    If mCode = "" Then mCode = rest
                    mName = TrimWS(Mid$(mName, q + 1))
                End If
            End If
        End If
    End If
    ExtractMeta = Array(mName, mCode, mDate, mMgr, mUnit)
End Function

Private Function NumCode(v As Variant) As String
    If CDbl(v) = Int(CDbl(v)) And Abs(CDbl(v)) < 1E+15 Then
        NumCode = Format$(CDbl(v), "0")
    Else
        NumCode = CStr(v)
    End If
End Function

Private Function IsCodeChars(ByVal s As String) As Boolean
    Dim i As Long, ch As String, w As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1): w = CodeOf(ch)
        If Not ((w >= 48 And w <= 57) Or (w >= 65 And w <= 90) Or (w >= 97 And w <= 122) Or ch = "-") Then Exit Function
    Next
    IsCodeChars = True
End Function

Private Function StripFundLabel(ByVal raw As String) As String
    ' 앞의 '펀 드 명 ▶', '펀드명:' 같은 머리말을 뗀다
    Dim p As Long, w As Variant
    p = SkipWS(raw, 1)
    For Each w In Array("펀", "드", "명")
        If Mid$(raw, p, 1) <> w Then StripFundLabel = raw: Exit Function
        p = SkipWS(raw, p + 1)
    Next
    If p <= Len(raw) Then
        If Mid$(raw, p, 1) = ":" Or Mid$(raw, p, 1) = ChrW(&H25B6&) Then p = p + 1
    End If
    p = SkipWS(raw, p)
    StripFundLabel = Mid$(raw, p)
End Function

Private Function FmtWon(ByVal x As Double) As String
    FmtWon = Format$(Int(x + 0.5), "#,##0")
End Function

' 총자산 = 신탁재산총자산 계정 → 순자산총액+부채 → 자산합계 순, 순자산+부채로 교차검증
Private Sub ComputeTotal(ld As Variant, fv() As Double, fpri() As Long, flabel() As String)
    Dim hasT As Boolean, hasN As Boolean, hasL As Boolean, hasA As Boolean, total As Double, evLabel As String
    If ld(LD_ERROR) <> "" Then
        ld(LD_VALUE) = Empty: ld(LD_CHECKS) = "bad": ld(LD_CHECKT) = "읽기 실패"
        Exit Sub
    End If
    hasT = fpri(0) >= 0: hasN = fpri(1) >= 0: hasL = fpri(2) >= 0: hasA = fpri(3) >= 0
    ' 1순위: 명세부의 '자산' 계정 당일잔액(차변). 평가 반영 총자산(총자산 계정·순자산+부채)과는 대조만 한다
    If hasA Then
        ld(LD_VALUE) = fv(3): ld(LD_METHOD) = flabel(3) & " 당일잔액"
        If hasT Then
            total = fv(0): evLabel = flabel(0)
        ElseIf hasN And hasL Then
            total = fv(1) + fv(2): evLabel = "순자산+부채"
        Else
            evLabel = ""
        End If
        If evLabel = "" Then
            ld(LD_CHECKS) = "mute": ld(LD_CHECKT) = "대조 값 없음"
        ElseIf Abs(fv(3) - total) <= 10 Then
            ld(LD_CHECKS) = "ok": ld(LD_CHECKT) = evLabel & " 값과 일치"
        Else
            ld(LD_CHECKS) = "mute": ld(LD_CHECKT) = evLabel & " 값과 " & FmtWon(fv(3) - total) & "원 차이"
        End If
    ElseIf hasT Then
        ld(LD_VALUE) = fv(0): ld(LD_METHOD) = flabel(0) & " 계정 (자산 계정 없음)"
        If hasN And hasL Then
            If Abs(fv(0) - (fv(1) + fv(2))) <= 10 Then
                ld(LD_CHECKS) = "ok": ld(LD_CHECKT) = "순자산+부채 일치"
            Else
                ld(LD_CHECKS) = "warn": ld(LD_CHECKT) = "순자산+부채와 " & FmtWon(fv(0) - fv(1) - fv(2)) & "원 차이"
            End If
        Else
            ld(LD_CHECKS) = "mute": ld(LD_CHECKT) = "교차검증 값 없음"
        End If
    ElseIf hasN And hasL Then
        total = fv(1) + fv(2)
        ld(LD_VALUE) = total: ld(LD_METHOD) = flabel(1) & " + " & flabel(2) & " (자산 계정 없음)"
        ld(LD_CHECKS) = "warn": ld(LD_CHECKT) = "자산 계정 없음 · 확인"
    Else
        ld(LD_VALUE) = Empty: ld(LD_METHOD) = "": ld(LD_CHECKS) = "bad": ld(LD_CHECKT) = "자산 계정을 찾지 못함"
    End If
End Sub

' 시트 값 배열들(sheets(0..nSheets-1))로 장부 레코드를 만든다. errMsg가 있으면 읽기 실패 장부
Public Function BuildLedger(ByVal path As String, ByVal fileName As String, sheets As Variant, ByVal nSheets As Long, ByVal errMsg As String) As Variant
    Dim res As Variant, n As Long, noNames() As String
    ReDim noNames(0)
    res = BuildLedgers(path, fileName, sheets, nSheets, errMsg, noNames, n)
    BuildLedger = res(0)
End Function

' 장부 파일 하나 → 장부 목록(nOut개). 보통 1개지만, 시트마다 다른 펀드의 원장이 든 통합 파일은 펀드마다 1개
Public Function BuildLedgers(ByVal path As String, ByVal fileName As String, sheets As Variant, ByVal nSheets As Long, ByVal errMsg As String, sheetNames As Variant, ByRef nOut As Long) As Variant
    Dim s As Long, rowsC() As Variant, rowsV() As Variant, rowsN() As Long, nRows As Long
    Dim fv() As Double, fpri() As Long, flabel() As String, headerRow As Long, hasHeader As Boolean, meta As Variant, rank As Long, f As Long, lim As Long
    Dim sFv() As Variant, sPri() As Variant, sLab() As Variant, sMeta() As Variant, sRank() As Long, sHdr() As Boolean, sName() As String
    Dim out() As Variant, emptyFv() As Double, emptyPri() As Long, emptyLab() As String
    nOut = 0
    If errMsg = "" And nSheets = 0 Then errMsg = "내용이 비어 있습니다."
    If errMsg <> "" Then
        ReDim emptyFv(5): ReDim emptyPri(5): ReDim emptyLab(5)
        For f = 0 To 5
            emptyPri(f) = -1
        Next
        ReDim out(0)
        out(0) = MakeLedger(path, fileName, "", emptyFv, emptyPri, emptyLab, Array("", "", "", "", ""), errMsg)
        nOut = 1
        BuildLedgers = out
        Exit Function
    End If
    ReDim sFv(nSheets - 1): ReDim sPri(nSheets - 1): ReDim sLab(nSheets - 1): ReDim sMeta(nSheets - 1)
    ReDim sRank(nSheets - 1): ReDim sHdr(nSheets - 1): ReDim sName(nSheets - 1)
    For s = 0 To nSheets - 1
        ToRows sheets(s), rowsC, rowsV, rowsN, nRows
        ExtractFields rowsC, rowsV, rowsN, nRows, fv, fpri, flabel, headerRow, hasHeader
        If hasHeader Then
            lim = headerRow - 1
        Else
            lim = 15
        End If
        meta = ExtractMeta(rowsC, rowsV, rowsN, nRows, lim)
        rank = 0
        For f = 0 To 5
            If fpri(f) >= 0 Then rank = rank + 1
        Next
        If hasHeader Then rank = rank + 2
        sFv(s) = CopyDoubles(fv): sPri(s) = CopyLongs(fpri, 6): sLab(s) = CopyStrings(flabel)
        sMeta(s) = meta: sRank(s) = rank: sHdr(s) = hasHeader
        sName(s) = ""
        If IsArray(sheetNames) Then
            If s <= UBound(sheetNames) Then sName(s) = CStr(sheetNames(s))
        End If
    Next
    ' 원장 탭 = '당일잔액/금일잔액' 머리글 + 자산·총자산·순자산 계정이 있는 탭(탭 이름과 무관).
    ' 원장 탭에서 서로 다른 펀드명이 2개 이상이면 펀드마다 장부 1개(펀드명 없는 원장 탭은 탭 이름을 펀드명 대신),
    ' 아니면 내용이 가장 많은 탭 하나
    Dim isLed() As Boolean, nNamed As Long, namedKeys() As String, fKey() As String, fIdx() As Long, nF As Long
    Dim k As Long, key As String, hit As Long, pv As Variant, nLed As Long
    ReDim isLed(nSheets - 1): ReDim namedKeys(nSheets - 1): ReDim fKey(nSheets - 1): ReDim fIdx(nSheets - 1)
    nNamed = 0: nF = 0: nLed = 0
    For s = 0 To nSheets - 1
        pv = sPri(s)
        isLed(s) = sHdr(s) And (pv(3) >= 0 Or pv(0) >= 0 Or pv(1) >= 0)
        If isLed(s) Then
            nLed = nLed + 1
            If sMeta(s)(0) <> "" Then
                key = NormName(sMeta(s)(0))
                hit = -1
                For k = 0 To nNamed - 1
                    If namedKeys(k) = key Then hit = k: Exit For
                Next
                If hit < 0 Then namedKeys(nNamed) = key: nNamed = nNamed + 1
            End If
        End If
    Next
    If nNamed >= 2 Then
        For s = 0 To nSheets - 1
            If isLed(s) Then
                If sMeta(s)(0) <> "" Then
                    key = NormName(sMeta(s)(0))
                Else
                    key = NormName(CleanFileName(sName(s) & ".x"))
                End If
                If key = "" Then key = sName(s)
                hit = -1
                For k = 0 To nF - 1
                    If fKey(k) = key Then hit = k: Exit For
                Next
                If hit < 0 Then
                    fKey(nF) = key: fIdx(nF) = s: nF = nF + 1
                ElseIf sRank(fIdx(hit)) < sRank(s) Then
                    fIdx(hit) = s
                End If
            End If
        Next
        ReDim out(nF - 1)
        For k = 0 To nF - 1
            s = fIdx(k)
            out(k) = MakeLedger(path, fileName, sName(s), sFv(s), sPri(s), sLab(s), sMeta(s), "")
        Next
        nOut = nF
    Else
        hit = -1
        For s = 0 To nSheets - 1
            If isLed(s) Or nLed = 0 Then
                If hit < 0 Then
                    hit = s
                ElseIf sRank(s) > sRank(hit) Then
                    hit = s
                End If
            End If
        Next
        ReDim out(0)
        out(0) = MakeLedger(path, fileName, "", sFv(hit), sPri(hit), sLab(hit), sMeta(hit), "")
        nOut = 1
    End If
    BuildLedgers = out
End Function

Private Function CopyDoubles(src() As Double) As Variant
    Dim r() As Double, i As Long
    ReDim r(UBound(src))
    For i = 0 To UBound(src)
        r(i) = src(i)
    Next
    CopyDoubles = r
End Function

Private Function CopyStrings(src() As String) As Variant
    Dim r() As String, i As Long
    ReDim r(UBound(src))
    For i = 0 To UBound(src)
        r(i) = src(i)
    Next
    CopyStrings = r
End Function

' 시트 하나의 추출 결과로 장부 레코드를 만든다. sheetName이 있으면 통합 파일의 시트 장부
Private Function MakeLedger(ByVal path As String, ByVal fileName As String, ByVal sheetName As String, fvIn As Variant, priIn As Variant, labIn As Variant, bmeta As Variant, ByVal errMsg As String) As Variant
    Dim ld As Variant, f As Long, mult As Double, bfv() As Double, bfpri() As Long, bflabel() As String
    ReDim ld(LD_SIZE - 1)
    ReDim bfv(5): ReDim bfpri(5): ReDim bflabel(5)
    For f = 0 To 5
        bfv(f) = fvIn(f): bfpri(f) = priIn(f): bflabel(f) = labIn(f)
    Next
    ld(LD_FILE) = fileName: ld(LD_PATH) = path: ld(LD_ERROR) = errMsg: ld(LD_SHEET) = sheetName: ld(LD_SUPER) = False
    ld(LD_NAME) = "": ld(LD_MANAGER) = "": ld(LD_CODE) = "": ld(LD_DATE) = "": ld(LD_METHOD) = ""
    ld(LD_VALUE) = Empty: ld(LD_NET) = Empty: ld(LD_LIAB) = Empty: ld(LD_NAV) = Empty
    Select Case bmeta(4)
        Case "천원": mult = 1000#
        Case "백만원": mult = 1000000#
        Case "억원": mult = 100000000#
        Case Else: mult = 1
    End Select
    For f = 0 To 3
        bfv(f) = bfv(f) * mult
    Next
    ld(LD_NAME) = bmeta(0): ld(LD_CODE) = bmeta(1): ld(LD_MANAGER) = bmeta(3)
    ' 통합 파일의 시트 장부는 시트 이름(예: 총계정원장_펀드명)을 파일명 대신 쓴다
    If sheetName <> "" Then
        ld(LD_TITLE) = CleanFileName(sheetName & ".x")
    Else
        ld(LD_TITLE) = CleanFileName(fileName)
    End If
    Dim cands() As String, nc As Long
    ReDim cands(1): nc = 0
    If ld(LD_NAME) <> "" Then cands(nc) = ld(LD_NAME): nc = nc + 1
    If ld(LD_TITLE) <> "" Then cands(nc) = ld(LD_TITLE): nc = nc + 1
    ld(LD_CANDS) = cands: ld(LD_NCANDS) = nc
    ld(LD_DATE) = bmeta(2)
    If ld(LD_DATE) = "" Then ld(LD_DATE) = DateFromFileName(fileName)
    If ld(LD_NAME) <> "" Then
        ld(LD_KEY) = NormName(ld(LD_NAME))
    Else
        ld(LD_KEY) = NormName(ld(LD_TITLE))
    End If
    If bfpri(1) >= 0 Then ld(LD_NET) = bfv(1)
    If bfpri(2) >= 0 Then ld(LD_LIAB) = bfv(2)
    If bfpri(4) >= 0 Then ld(LD_NAV) = bfv(4)
    ComputeTotal ld, bfv, bfpri, bflabel
    MakeLedger = ld
End Function

' ======================================================================
'  대상 엑셀 (시트 값 배열에서 '종목명'·'총자산' 머리행 찾기)
' ======================================================================
Public Function ParseTargetArray(arr As Variant) As Boolean
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long, hr As Long, j As Long, i As Long, s As String, v As Variant
    Dim nameCol As Long, assetCol As Long, codeCol As Long, clsCol As Long, dateCol As Long
    If Not IsArray(arr) Then Exit Function
    r1 = LBound(arr, 1): r2 = UBound(arr, 1): c1 = LBound(arr, 2): c2 = UBound(arr, 2)
    For hr = r1 To IIf(r2 < r1 + 14, r2, r1 + 14)
        nameCol = 0: assetCol = 0: codeCol = 0: clsCol = 0: dateCol = 0
        For j = c1 To c2
            v = arr(hr, j)
            If VarType(v) = VT_STRING Then
                s = RemoveWS(CStr(v))
                If nameCol = 0 And (s = "종목명" Or s = "펀드명" Or s = "상품명") Then nameCol = j
            End If
        Next
        For j = c1 To c2
            v = arr(hr, j)
            If VarType(v) = VT_STRING Then
                s = RemoveWS(CStr(v))
                If nameCol = 0 And (InStr(s, "종목명") > 0 Or InStr(s, "펀드명") > 0) Then nameCol = j
                If assetCol = 0 And InStr(s, "총자산") > 0 Then assetCol = j
                If codeCol = 0 And (InStr(s, "표준코드") > 0 Or InStr(UCase$(s), "ISIN") > 0) Then codeCol = j
                If clsCol = 0 And InStr(s, "분류") > 0 Then clsCol = j
                If dateCol = 0 And Left$(s, 3) = "기준일" Then dateCol = j
            End If
        Next
        If nameCol > 0 And assetCol > 0 Then
            tHeaderRow = hr - r1 + 1
            tAssetColRel = assetCol - c1 + 1
            tAssetHeader = CStr(arr(hr, assetCol))
            tBaseDate = DateFromText(arr(hr, assetCol))
            tCount = 0
            ReDim tRow(0): ReDim tName(0): ReDim tCode(0): ReDim tCls(0): ReDim tExist(0)
            For i = hr + 1 To r2
                v = arr(i, nameCol)
                If VarType(v) = VT_STRING Then
                    If TrimWS(CStr(v)) <> "" Then
                        If tBaseDate = "" And dateCol > 0 Then
                            If VarType(arr(i, dateCol)) = VT_STRING Then
                                tBaseDate = DateFromText(arr(i, dateCol))
                            Else
                                tBaseDate = SerialToISO(arr(i, dateCol))
                            End If
                        End If
                        ReDim Preserve tRow(tCount): ReDim Preserve tName(tCount): ReDim Preserve tCode(tCount): ReDim Preserve tCls(tCount): ReDim Preserve tExist(tCount)
                        tRow(tCount) = i - r1 + 1
                        tName(tCount) = TrimWS(CStr(v))
                        If codeCol > 0 Then tCode(tCount) = CellText(arr(i, codeCol))
                        If clsCol > 0 Then tCls(tCount) = CellText(arr(i, clsCol))
                        tExist(tCount) = arr(i, assetCol)
                        tCount = tCount + 1
                    End If
                End If
            Next
            ParseTargetArray = True
            Exit Function
        End If
    Next
End Function

Private Function CellText(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function

Private Function IsTextMark(v As Variant) As Boolean
    Dim ok As Boolean
    If VarType(v) <> VT_STRING Then Exit Function
    If TrimWS(CStr(v)) = "" Then Exit Function
    ToNum v, ok
    IsTextMark = Not ok
End Function

Private Function HasNumber(v As Variant) As Boolean
    Dim ok As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    ToNum v, ok
    HasNumber = ok
End Function

' ======================================================================
'  행별 판정 (웹·파이썬판과 같은 규칙)
'  forcedU(u): 고유 종목명별 지정 매핑 (-2 없음, -1 매칭 안 함, 0 이상 장부 번호)
' ======================================================================
Public Sub ComputeResults(lds() As Variant, ByVal nL As Long, ByVal baseYmd As String, ByVal overwrite As Boolean, ByVal includeReview As Boolean, forcedByName As Collection)
    Dim i As Long, u As Long, idxC As New Collection, lid As Long, isForced As Boolean, nameStatus As String, dateOk As Boolean
    Dim forced() As Long, bestS As Double, bestL As Long, l As Long, ld As Variant, key As String, note As String
    ' 고유 종목명
    uCount = 0
    ReDim uNames(IIf(tCount > 0, tCount - 1, 0)): ReDim rU(IIf(tCount > 0, tCount - 1, 0))
    For i = 0 To tCount - 1
        key = CK(tName(i))
        If CollHas(idxC, key) Then
            rU(i) = idxC.Item(key)
        Else
            uNames(uCount) = tName(i)
            idxC.Add uCount, key
            rU(i) = uCount
            uCount = uCount + 1
        End If
    Next
    ReDim forced(IIf(uCount > 0, uCount - 1, 0))
    For u = 0 To uCount - 1
        forced(u) = -2
        If Not forcedByName Is Nothing Then
            key = CK(NormName(uNames(u)))
            If CollHas(forcedByName, key) Then forced(u) = ResolveForced(forcedByName.Item(key), lds, nL)
        End If
    Next
    If uCount > 0 Then MatchNames uNames, uCount, lds, nL, forced
    ReDim rStatus(IIf(tCount > 0, tCount - 1, 0)): ReDim rLed(IIf(tCount > 0, tCount - 1, 0))
    ReDim rScore(IIf(tCount > 0, tCount - 1, 0)): ReDim rWhy(IIf(tCount > 0, tCount - 1, 0))
    ReDim rInc(IIf(tCount > 0, tCount - 1, 0)): ReDim rNotes(IIf(tCount > 0, tCount - 1, 0))
    For i = 0 To tCount - 1
        u = rU(i)
        rLed(i) = -1: rScore(i) = 0: rWhy(i) = "": rInc(i) = False: rNotes(i) = ""
        If IsTextMark(tExist(i)) Then
            rStatus(i) = "skip": rNotes(i) = "'" & TrimWS(CStr(tExist(i))) & "' 표시 행"
            GoTo NextRow
        End If
        isForced = (forced(u) <> -2)
        lid = mAssign(u)
        If lid < 0 Then
            If isForced Then
                rStatus(i) = "manual": rNotes(i) = "매칭 안 함으로 지정됨"
            Else
                rStatus(i) = "nomatch"
                bestS = -1: bestL = -1
                For l = 0 To nL - 1
                    If mScores(u, l) > bestS Then bestS = mScores(u, l): bestL = l
                Next
                If bestL >= 0 And bestS >= 0.3 Then
                    If lds(bestL)(LD_NAME) <> "" Then
                        note = lds(bestL)(LD_NAME)
                    Else
                        note = lds(bestL)(LD_TITLE)
                    End If
                    rNotes(i) = "가장 가까운 후보 " & FixedStr(bestS, 2) & " " & note
                End If
            End If
            GoTo NextRow
        End If
        ld = lds(lid)
        rLed(i) = lid
        If isForced Then
            rScore(i) = 1: rWhy(i) = "지정된 매핑"
        Else
            rScore(i) = mScores(u, lid): rWhy(i) = mWhys(u, lid)
        End If
        nameStatus = MatchStatus(u, lid, nL)
        dateOk = (baseYmd = "" Or ld(LD_DATE) = "" Or ld(LD_DATE) = baseYmd)
        If Not dateOk Then AddNote rNotes(i), "기준일 " & ld(LD_DATE) & " ≠ " & baseYmd
        If ld(LD_CHECKS) = "warn" Then AddNote rNotes(i), ld(LD_CHECKT)
        If IsEmpty(ld(LD_VALUE)) Then AddNote rNotes(i), "총자산 값 없음"
        If isForced Then
            rStatus(i) = "manual"
        ElseIf nameStatus = "auto" And dateOk And ld(LD_CHECKS) <> "warn" And Not IsEmpty(ld(LD_VALUE)) Then
            rStatus(i) = "auto"
        Else
            rStatus(i) = "review"
            If nameStatus <> "auto" Then
                If mScores(u, lid) < AUTO_SCORE Then
                    rNotes(i) = PrependNote(rNotes(i), "유사도 낮음")
                Else
                    rNotes(i) = PrependNote(rNotes(i), "비슷한 후보 있음")
                End If
            End If
        End If
        rInc(i) = Not IsEmpty(ld(LD_VALUE)) And (rStatus(i) = "auto" Or (rStatus(i) = "manual" And dateOk) Or (includeReview And rStatus(i) = "review"))
        If rInc(i) And HasNumber(tExist(i)) And Not overwrite Then
            rInc(i) = False
            AddNote rNotes(i), "기존 값 있음 (덮어쓰기 꺼짐)"
        End If
NextRow:
    Next
End Sub

Private Sub AddNote(ByRef notes As String, ByVal s As String)
    If notes = "" Then
        notes = s
    Else
        notes = notes & " / " & s
    End If
End Sub

Private Function PrependNote(ByVal notes As String, ByVal s As String) As String
    If notes = "" Then
        PrependNote = s
    Else
        PrependNote = s & " / " & notes
    End If
End Function

Private Function ResolveForced(v As Variant, lds() As Variant, ByVal nL As Long) As Long
    Dim key As String, l As Long
    key = CStr(v)
    If key = "__none__" Then ResolveForced = -1: Exit Function
    ResolveForced = -2
    For l = 0 To nL - 1
        If IsUsable(lds(l)) Then
            If key = lds(l)(LD_KEY) Or key = NormName(lds(l)(LD_TITLE)) Or key = NormName(FileStem(lds(l)(LD_FILE))) Then ResolveForced = l: Exit Function
        End If
    Next
End Function

Private Function IsUsable(ld As Variant) As Boolean
    IsUsable = False
    If ld(LD_ERROR) <> "" Then Exit Function
    If UBound(ld) >= LD_SUPER Then
        If Not IsEmpty(ld(LD_SUPER)) Then
            If CBool(ld(LD_SUPER)) Then Exit Function
        End If
    End If
    IsUsable = True
End Function

' 파일명의 결산 표시: "post"(결산후·결산반영) / "pre"(결산전) / ""
Public Function SettleTag(ByVal fileName As String) As String
    Dim t As String
    t = NormText(fileName)
    If InStr(t, "결산후") > 0 Or InStr(t, "결산 후") > 0 Or InStr(t, "결산반영") > 0 Then
        SettleTag = "post"
    ElseIf InStr(t, "결산전") > 0 Or InStr(t, "결산 전") > 0 Then
        SettleTag = "pre"
    Else
        SettleTag = ""
    End If
End Function

' 결산 표시를 뺀 파일명(확장자 제외) — 결산전/결산후 짝을 찾는 열쇠
Public Function SettleKey(ByVal fileName As String) As String
    Dim t As String, p As Long, w As Variant
    t = NormText(fileName)
    p = InStrRev(t, ".")
    If p > 0 And p < Len(t) Then t = Left$(t, p - 1)
    For Each w In Array("결산전", "결산 전", "결산후", "결산 후", "결산반영")
        t = Replace(t, CStr(w), "")
    Next
    SettleKey = NormName(t)
End Function

' 이름이 같은 결산전·결산후 파일이 모두 있으면 결산전 장부를 빼도록 표시 (시트 장부는 시트까지 같아야 짝)
Public Sub MarkSettlement(lds() As Variant, ByVal nL As Long)
    Dim l As Long, m As Long, key As String, ld As Variant, hit As Boolean
    For l = 0 To nL - 1
        ld = lds(l)
        ld(LD_SUPER) = False
        If SettleTag(ld(LD_FILE)) = "pre" Then
            key = SettleKey(ld(LD_FILE)) & "|" & ld(LD_SHEET)
            hit = False
            For m = 0 To nL - 1
                If lds(m)(LD_ERROR) = "" Then
                    If SettleTag(lds(m)(LD_FILE)) = "post" Then
                        If SettleKey(lds(m)(LD_FILE)) & "|" & lds(m)(LD_SHEET) = key Then hit = True: Exit For
                    End If
                End If
            Next
            ld(LD_SUPER) = hit
        End If
        lds(l) = ld
    Next
End Sub

Private Function LedgerFileLabel(ld As Variant) As String
    If ld(LD_SHEET) <> "" Then
        LedgerFileLabel = ld(LD_FILE) & " [" & ld(LD_SHEET) & "]"
    Else
        LedgerFileLabel = ld(LD_FILE)
    End If
End Function

Private Function FileStem(ByVal f As String) As String
    Dim p As Long
    p = InStrRev(f, ".")
    If p > 0 Then
        FileStem = Left$(f, p - 1)
    Else
        FileStem = f
    End If
End Function

' 매칭된 장부의 가장 흔한 기준일이 기준일과 다르면 그 날짜 (해당 행 수는 nMis)
Public Function DateMismatch(lds() As Variant, ByVal baseYmd As String, ByRef nMis As Long) As String
    Dim dates() As String, cnts() As Long, nd As Long, i As Long, k As Long, d As String, best As Long, common As String, found As Boolean
    nMis = 0
    If baseYmd = "" Then Exit Function
    ReDim dates(0): ReDim cnts(0)
    For i = 0 To tCount - 1
        If rLed(i) >= 0 Then
            d = lds(rLed(i))(LD_DATE)
            If d <> "" Then
                found = False
                For k = 0 To nd - 1
                    If dates(k) = d Then cnts(k) = cnts(k) + 1: found = True: Exit For
                Next
                If Not found Then
                    ReDim Preserve dates(nd): ReDim Preserve cnts(nd)
                    dates(nd) = d: cnts(nd) = 1: nd = nd + 1
                End If
            End If
        End If
    Next
    If nd = 0 Then Exit Function
    best = -1
    For k = 0 To nd - 1
        If cnts(k) > best Then best = cnts(k): common = dates(k)
    Next
    If common = baseYmd Then Exit Function
    For i = 0 To tCount - 1
        If rLed(i) >= 0 Then
            d = lds(rLed(i))(LD_DATE)
            If d <> "" And d <> baseYmd Then nMis = nMis + 1
        End If
    Next
    DateMismatch = common
End Function

Public Function StatusKo(ByVal s As String) As String
    Select Case s
        Case "auto": StatusKo = "자동"
        Case "review": StatusKo = "검토"
        Case "manual": StatusKo = "수동"
        Case "nomatch": StatusKo = "미매칭"
        Case Else: StatusKo = "제외"
    End Select
End Function

' ======================================================================
'  엑셀 실행부 (파일 고르기 · 열기 · 저장)
' ======================================================================
Public Sub RunTotalAssets()
    Dim targetPath As String, ledgerDir As String, baseYmd As String, wbT As Object, ws As Object, found As Boolean
    Dim files() As String, nF As Long, i As Long, lds() As Variant, nL As Long, sheets As Variant, nS As Long, errMsg As String
    Dim prevCalc As Long, prevSec As Long, includeReview As Boolean, mapping As Collection, common As String, nMis As Long
    Dim firstRow As Long, firstCol As Long, nInc As Long, nReview As Long, outPath As String, reportPath As String, msg As String
    Dim folder As String, stem As String, ext As String, p As Long, wsT As Object
    Dim sheetNames As Variant, part As Variant, nPart As Long, k As Long

    targetPath = PickPath(False, "총자산을 채울 대상 엑셀을 고르세요")
    If targetPath = "" Then Exit Sub
    If Not OpenWorkbookByPath(targetPath) Is Nothing Then
        MsgBox "대상 엑셀이 지금 열려 있습니다. 닫은 뒤 다시 실행하세요." & vbCrLf & "(원본은 바꾸지 않고 새 파일로 저장합니다)", vbExclamation, "총자산 자동입력"
        Exit Sub
    End If
    ledgerDir = PickPath(True, "기준가격대장이 들어 있는 폴더를 고르세요")
    If ledgerDir = "" Then Exit Sub

    prevSec = Application.AutomationSecurity
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    On Error GoTo Fail

    ' 1) 대상 엑셀
    Set wbT = Workbooks.Open(Filename:=targetPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    For Each ws In wbT.Worksheets
        If ParseTargetArray(SheetValues(ws)) Then
            firstRow = ws.UsedRange.Row: firstCol = ws.UsedRange.Column
            Set wsT = ws
            found = True
            Exit For
        End If
    Next
    If Not found Then
        wbT.Close SaveChanges:=False
        Err.Raise vbObjectError + 1, , "'종목명'과 '총자산' 머리글이 있는 시트를 찾지 못했습니다."
    End If
    Application.ScreenUpdating = True
    baseYmd = AskBaseDate(tBaseDate)
    If baseYmd = "?" Then wbT.Close SaveChanges:=False: GoTo Done
    Application.ScreenUpdating = False

    ' 2) 기준가격대장 (매크로·외부 링크는 실행하지 않고 읽기 전용으로 연다)
    nF = ListLedgerFiles(ledgerDir, files, targetPath)
    If nF = 0 Then
        wbT.Close SaveChanges:=False
        Err.Raise vbObjectError + 2, , "폴더에 엑셀 파일(.xlsx .xls .xlsm .xlsb .csv)이 없습니다." & vbCrLf & ledgerDir
    End If
    ReDim lds(0)
    nL = 0
    Application.AutomationSecurity = 3
    For i = 0 To nF - 1
        Application.StatusBar = "기준가격대장 읽는 중… " & (i + 1) & "/" & nF
        DoEvents
        errMsg = ReadWorkbookSheets(files(i), sheets, nS, sheetNames)
        part = BuildLedgers(files(i), FileNameOf(files(i)), sheets, nS, errMsg, sheetNames, nPart)
        For k = 0 To nPart - 1
            ReDim Preserve lds(nL)
            lds(nL) = part(k)
            nL = nL + 1
        Next
    Next
    MarkSettlement lds, nL
    Application.AutomationSecurity = prevSec
    Application.StatusBar = False

    ' 3) 매칭·판정
    Set mapping = LoadMapping()
    ComputeResults lds, nL, baseYmd, False, False, mapping
    common = DateMismatch(lds, baseYmd, nMis)
    If common <> "" Then
        Application.ScreenUpdating = True
        If MsgBox("기준일은 " & baseYmd & "인데 매칭된 장부는 대부분 " & common & "자입니다." & vbCrLf & _
                  "그래서 " & nMis & "행이 '검토'로 분류됐습니다." & vbCrLf & vbCrLf & _
                  "기준일을 " & common & "로 바꿔 다시 판정할까요?", vbYesNo + vbQuestion, "총자산 자동입력") = vbYes Then
            baseYmd = common
            ComputeResults lds, nL, baseYmd, False, False, mapping
        End If
    End If
    nReview = 0
    For i = 0 To tCount - 1
        If rStatus(i) = "review" And rLed(i) >= 0 Then
            If Not IsEmpty(lds(rLed(i))(LD_VALUE)) Then nReview = nReview + 1
        End If
    Next
    If nReview > 0 Then
        Application.ScreenUpdating = True
        includeReview = (MsgBox("'검토' 행이 " & nReview & "개 있습니다 (유사도가 낮거나 검증·기준일이 맞지 않음)." & vbCrLf & _
                                "이 행들도 총자산을 입력할까요?" & vbCrLf & vbCrLf & _
                                "아니오를 누르면 '자동'·'수동' 행만 입력합니다. 검토 행은 매칭결과 파일에서 확인하세요.", _
                                vbYesNo + vbDefaultButton2 + vbQuestion, "총자산 자동입력") = vbYes)
        If includeReview Then ComputeResults lds, nL, baseYmd, False, True, mapping
        Application.ScreenUpdating = False
    End If

    ' 4) 저장 (원본은 그대로 두고 새 파일로)
    p = InStrRev(targetPath, Application.PathSeparator)
    folder = Left$(targetPath, p)
    stem = Mid$(targetPath, p + 1)
    p = InStrRev(stem, ".")
    If p > 0 Then
        ext = Mid$(stem, p)
        stem = Left$(stem, p - 1)
    Else
        ext = ".xlsx"
    End If
    nInc = 0
    For i = 0 To tCount - 1
        If rInc(i) Then
            wsT.Cells(firstRow + tRow(i) - 1, firstCol + tAssetColRel - 1).Value = Int(CDbl(lds(rLed(i))(LD_VALUE)) + 0.5)
            nInc = nInc + 1
        End If
    Next
    If nInc > 0 Then
        outPath = folder & stem & "_총자산입력" & ext
        wbT.SaveAs Filename:=outPath, FileFormat:=wbT.FileFormat
    End If
    wbT.Close SaveChanges:=False
    Set wbT = Nothing
    reportPath = folder & stem & "_매칭결과.xlsx"
    WriteReport reportPath, lds, nL, baseYmd

    msg = "장부 " & nL & "개를 읽었습니다." & vbCrLf & SummaryLine() & vbCrLf & vbCrLf
    If nInc > 0 Then
        msg = msg & "총자산 " & nInc & "칸을 채운 파일:" & vbCrLf & outPath & vbCrLf & vbCrLf
    Else
        msg = msg & "입력할 행이 없어 엑셀은 저장하지 않았습니다." & vbCrLf & vbCrLf
    End If
    msg = msg & "매칭 근거:" & vbCrLf & reportPath
    Application.ScreenUpdating = True
    MsgBox msg, vbInformation, "총자산 자동입력"
Done:
    RestoreApp prevSec
    Exit Sub
Fail:
    Dim emsg As String
    emsg = Err.Description
    On Error Resume Next
    If Not wbT Is Nothing Then wbT.Close SaveChanges:=False
    RestoreApp prevSec
    MsgBox "오류가 났습니다." & vbCrLf & vbCrLf & emsg & vbCrLf & vbCrLf & _
           "같은 이름의 결과 파일(_총자산입력, _매칭결과)이 열려 있으면 닫고 다시 실행하세요.", vbExclamation, "총자산 자동입력"
End Sub

Private Sub RestoreApp(ByVal prevSec As Long)
    On Error Resume Next
    Application.AutomationSecurity = prevSec
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Function SheetValues(ws As Object) As Variant
    Dim v As Variant, a(1 To 1, 1 To 1) As Variant
    v = ws.UsedRange.Value2
    If IsArray(v) Then
        SheetValues = v
    Else
        a(1, 1) = v
        SheetValues = a
    End If
End Function

Private Function PickPath(ByVal folderMode As Boolean, ByVal title As String) As String
    With Application.FileDialog(IIf(folderMode, 4, 3))
        .Title = title
        .AllowMultiSelect = False
        If Not folderMode Then
            .Filters.Clear
            .Filters.Add "Excel", "*.xlsx;*.xlsm;*.xls"
        End If
        If .Show = -1 Then PickPath = .SelectedItems(1)
    End With
End Function

Private Function AskBaseDate(ByVal dflt As String) As String
    Dim s As String
    Do
        s = InputBox("기준일을 확인하세요 (예: 2026-09-30)." & vbCrLf & vbCrLf & _
                     "대상 엑셀 머리글에서 읽은 값이 들어 있습니다. 장부 기준일과 다르면 '검토'로 분류됩니다." & vbCrLf & _
                     "비워 두면 기준일을 확인하지 않습니다.", "총자산 자동입력 - 기준일", dflt)
        If StrPtr(s) = 0 Then AskBaseDate = "?": Exit Function
        s = TrimWS(s)
        If s = "" Then AskBaseDate = "": Exit Function
        If DateFromText(s) <> "" Then AskBaseDate = DateFromText(s): Exit Function
        MsgBox "2026-09-30 형식으로 적어 주세요.", vbExclamation, "총자산 자동입력"
    Loop
End Function

Private Function FileNameOf(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, Application.PathSeparator)
    FileNameOf = Mid$(path, p + 1)
End Function

' 폴더(하위 폴더 포함)의 엑셀 파일 목록. 결과·임시·대상 파일은 뺀다
Private Function ListLedgerFiles(ByVal root As String, ByRef files() As String, ByVal targetPath As String) As Long
    Dim queue() As String, nq As Long, qi As Long, d As String, f As String, sep As String, n As Long, lower As String, subs() As String, ns As Long, k As Long
    sep = Application.PathSeparator
    ReDim queue(0): queue(0) = root: nq = 1: qi = 0
    ReDim files(0): n = 0
    Do While qi < nq
        d = queue(qi): qi = qi + 1
        If Right$(d, 1) <> sep Then d = d & sep
        ReDim subs(0): ns = 0
        f = Dir(d & "*", vbDirectory)
        Do While f <> ""
            If f <> "." And f <> ".." Then
                If (GetAttr(d & f) And vbDirectory) = vbDirectory Then
                    ReDim Preserve subs(ns): subs(ns) = d & f: ns = ns + 1
                Else
                    lower = LCase$(f)
                    If (lower Like "*.xlsx" Or lower Like "*.xlsm" Or lower Like "*.xlsb" Or lower Like "*.xls" Or lower Like "*.csv") _
                       And Left$(f, 2) <> "~$" And InStr(f, "_총자산입력") = 0 And InStr(f, "_매칭결과") = 0 _
                       And LCase$(d & f) <> LCase$(targetPath) And LCase$(d & f) <> LCase$(ThisWorkbook.FullName) Then
                        ReDim Preserve files(n): files(n) = d & f: n = n + 1
                    End If
                End If
            End If
            f = Dir()
        Loop
        For k = 0 To ns - 1
            ReDim Preserve queue(nq): queue(nq) = subs(k): nq = nq + 1
        Next
    Loop
    ListLedgerFiles = n
End Function

' 장부 파일의 모든 시트 값을 읽는다. 실패하면 오류 문구를 돌려준다
Private Function ReadWorkbookSheets(ByVal path As String, ByRef sheets As Variant, ByRef nS As Long, ByRef names As Variant) As String
    Dim wb As Object, ws As Object, arr() As Variant, nm() As String, k As Long, wasOpen As Boolean
    nS = 0
    On Error GoTo Bad
    ' 이미 열어 둔 장부는 그대로 읽고 닫지 않는다
    Set wb = OpenWorkbookByPath(path)
    wasOpen = Not wb Is Nothing
    If Not wasOpen Then
        Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, _
                                Notify:=False, AddToMru:=False, Password:="__no_password__")
    End If
    ReDim arr(wb.Worksheets.Count - 1): ReDim nm(wb.Worksheets.Count - 1)
    For Each ws In wb.Worksheets
        arr(k) = SheetValues(ws)
        nm(k) = ws.Name
        k = k + 1
    Next
    If Not wasOpen Then wb.Close SaveChanges:=False
    sheets = arr
    names = nm
    nS = k
    ReadWorkbookSheets = ""
    Exit Function
Bad:
    ReadWorkbookSheets = "파일을 열 수 없습니다. 암호·DRM이 걸려 있는지 확인하세요."
    On Error Resume Next
    If Not wb Is Nothing And Not wasOpen Then wb.Close SaveChanges:=False
End Function

' 이미 열려 있는 통합문서를 경로로 찾는다 (없으면 Nothing)
Private Function OpenWorkbookByPath(ByVal path As String) As Object
    Dim w As Object
    For Each w In Workbooks
        If LCase$(w.FullName) = LCase$(path) Then
            Set OpenWorkbookByPath = w
            Exit Function
        End If
    Next
End Function

' (선택) 이 통합문서의 '매핑' 시트: A열 대상 종목명, B열 장부 펀드명/파일명 ('없음'이면 매칭하지 않음)
Private Function LoadMapping() As Collection
    Dim ws As Object, r As Long, a As String, b As String, c As New Collection, key As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("매핑")
    On Error GoTo 0
    If ws Is Nothing Then Set LoadMapping = c: Exit Function
    For r = 2 To ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        a = TrimWS(CStr(ws.Cells(r, 1).Value))
        b = TrimWS(CStr(ws.Cells(r, 2).Value))
        If a <> "" Then
            key = CK(NormName(a))
            If Not CollHas(c, key) Then
                If b = "" Or b = "없음" Then
                    c.Add "__none__", key
                Else
                    c.Add NormName(b), key
                End If
            End If
        End If
    Next
    Set LoadMapping = c
End Function

Private Function SummaryLine() As String
    Dim k As Variant, i As Long, n As Long, nInc As Long, s As String
    For Each k In Array("auto", "review", "manual", "nomatch", "skip")
        n = 0
        For i = 0 To tCount - 1
            If rStatus(i) = k Then n = n + 1
        Next
        If s <> "" Then s = s & " · "
        s = s & StatusKo(CStr(k)) & " " & n
    Next
    For i = 0 To tCount - 1
        If rInc(i) Then nInc = nInc + 1
    Next
    SummaryLine = s & " · 입력 " & nInc & "행"
End Function

' 매칭 결과를 새 통합문서로 저장 (행별 결과 + 장부별 추출값)
Private Sub WriteReport(ByVal path As String, lds() As Variant, ByVal nL As Long, ByVal baseYmd As String)
    Dim wb As Object, ws As Object, i As Long, r As Long, ld As Variant, color As Long, data() As Variant, ncol As Long, l As Long
    Set wb = Workbooks.Add(-4167)
    Set ws = wb.Worksheets(1)
    ws.Name = "매칭결과"
    ncol = 13
    ReDim data(0 To tCount, 1 To ncol)
    data(0, 1) = "행": data(0, 2) = "종목명": data(0, 3) = "표준코드": data(0, 4) = "상태": data(0, 5) = "입력"
    data(0, 6) = "유사도": data(0, 7) = "매칭근거": data(0, 8) = "기준가격대장 파일": data(0, 9) = "장부 펀드명"
    data(0, 10) = "장부 기준일": data(0, 11) = "총자산": data(0, 12) = "검증": data(0, 13) = "메모"
    For i = 0 To tCount - 1
        r = i + 1
        data(r, 1) = tRow(i): data(r, 2) = tName(i): data(r, 3) = tCode(i)
        data(r, 4) = StatusKo(rStatus(i)): data(r, 5) = IIf(rInc(i), "Y", "")
        If rScore(i) > 0 Then data(r, 6) = Round(rScore(i), 3)
        data(r, 7) = rWhy(i): data(r, 13) = rNotes(i)
        If rLed(i) >= 0 Then
            ld = lds(rLed(i))
            data(r, 8) = LedgerFileLabel(ld)
            data(r, 9) = IIf(ld(LD_NAME) <> "", ld(LD_NAME), ld(LD_TITLE))
            data(r, 10) = ld(LD_DATE)
            If Not IsEmpty(ld(LD_VALUE)) Then data(r, 11) = Int(CDbl(ld(LD_VALUE)) + 0.5)
            data(r, 12) = ld(LD_CHECKT)
        End If
    Next
    ws.Range(ws.Cells(1, 1), ws.Cells(tCount + 1, ncol)).Value = data
    ws.Range(ws.Cells(1, 1), ws.Cells(1, ncol)).Font.Bold = True
    ws.Columns(11).NumberFormat = "#,##0"
    For i = 0 To tCount - 1
        Select Case rStatus(i)
            Case "auto": color = RGB(227, 243, 234)
            Case "review": color = RGB(251, 240, 214)
            Case "manual": color = RGB(226, 239, 233)
            Case Else: color = -1
        End Select
        If color >= 0 Then ws.Range(ws.Cells(i + 2, 1), ws.Cells(i + 2, ncol)).Interior.Color = color
    Next
    ws.Columns("A:M").AutoFit
    ' 장부별
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(1))
    ws.Name = "장부별"
    ReDim data(0 To nL, 1 To 9)
    data(0, 1) = "파일": data(0, 2) = "인식한 펀드명": data(0, 3) = "운용사": data(0, 4) = "기준일": data(0, 5) = "총자산"
    data(0, 6) = "순자산": data(0, 7) = "부채": data(0, 8) = "산출 근거": data(0, 9) = "검증 / 오류"
    For l = 0 To nL - 1
        ld = lds(l)
        data(l + 1, 1) = LedgerFileLabel(ld)
        data(l + 1, 2) = IIf(ld(LD_NAME) <> "", ld(LD_NAME), ld(LD_TITLE))
        data(l + 1, 3) = ld(LD_MANAGER): data(l + 1, 4) = ld(LD_DATE)
        If Not IsEmpty(ld(LD_VALUE)) Then data(l + 1, 5) = Int(CDbl(ld(LD_VALUE)) + 0.5)
        If Not IsEmpty(ld(LD_NET)) Then data(l + 1, 6) = ld(LD_NET)
        If Not IsEmpty(ld(LD_LIAB)) Then data(l + 1, 7) = ld(LD_LIAB)
        data(l + 1, 8) = ld(LD_METHOD)
        If ld(LD_ERROR) <> "" Then
            data(l + 1, 9) = ld(LD_ERROR)
        ElseIf Not IsUsable(ld) Then
            data(l + 1, 9) = "같은 이름의 결산후 파일이 있어 제외"
        Else
            data(l + 1, 9) = ld(LD_CHECKT)
        End If
    Next
    ws.Range(ws.Cells(1, 1), ws.Cells(nL + 1, 9)).Value = data
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 9)).Font.Bold = True
    ws.Range("E:G").NumberFormat = "#,##0"
    ws.Columns("A:I").AutoFit
    wb.Worksheets(1).Activate
    wb.SaveAs Filename:=path, FileFormat:=51
    wb.Close SaveChanges:=False
End Sub
