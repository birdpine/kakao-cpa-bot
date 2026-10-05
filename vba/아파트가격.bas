Attribute VB_Name = "아파트가격"
'==============================================================================
' 아파트 가격 조회 (기준시가 / 유사매매사례가액(추정) / 실거래가액)
'------------------------------------------------------------------------------
' [입력 시트 구조]
'   A5:A~ : 소재지 (예: "대구 수성구 상록로 69, 래미안범어, 105동 1906호")
'   2행   : 평가기준일 (B2, H2, N2 ... 6열 단위, 병합셀의 첫 셀에 날짜)
'   3행   : 기준시가 / 매매사례액 / 실거래가액 (각 2열)
'   4행   : 공시일·금액 / 고시일자·금액 / 매매계약일·금액
'           → 미조회 시 "N/A", 실거래가액·매매사례가액이 없으면 "-"
'
' [사용 API] - 모두 무료, 개별 인증키 필요 ("아파트_설정" 시트에 입력)
'   1) 행정안전부 도로명주소 검색 API (business.juso.go.kr) : 주소 → PNU·법정동코드
'   2) 국토교통부 공동주택가격 속성조회 (VWorld)           : 동·호별 공동주택가격
'   3) 국토교통부 아파트 매매 실거래가 상세 (data.go.kr)    : 실거래·유사매매사례
'
' [판단 기준]
'   - 기준시가 : 상증세법 제61조 제1항 제4호 → 평가기준일 현재 공시된 공동주택가격
'                (공시일 이전 평가기준일은 직전 연도 가격 적용, 공시일은 설정 시트)
'   - 실거래가액 : 동·층·전용면적이 평가대상과 일치하는 거래(호수는 미공개 → 추정)
'   - 매매사례가액 : 상증세법 시행령 제49조 제4항, 시행규칙 제15조 제3항 준용 추정
'                홈택스 "평가기간 내 유사 물건" : 증여(전6·후3) 우선, 없으면 상속(전6·후6)
'                동일 단지 + 전용면적 차이 5% 이내 + 공동주택가격 차이 5% 이내,
'                공동주택가격 차이 최소 → 동률 시 평가기준일에 가장 가까운 거래
'                ※ 홈택스 조회값과 다를 수 있음(호 단위 공시가격 매칭 불가)
'==============================================================================
Option Explicit

Private Const SHT_CFG As String = "아파트_설정"
Private Const SHT_LOG As String = "아파트_조회로그"
Private Const SHT_APT As String = "아파트"         ' 조회 양식 시트명
Private Const FIRST_ROW As Long = 5
Private Const HDR_DATE_ROW As Long = 2
Private Const FIRST_COL As Long = 2
Private Const GROUP_W As Long = 6           ' 평가기준일 1개당 열 수
Private Const AREA_TOL As Double = 0.1       ' 동일 호 판정 시 전용면적 허용오차(㎡)
' 평가기간 (상증세법 제60조 제1항) : 평가기준일 전 6개월 ~
'   매매사례가액 : ① 증여 기준(후 3개월)으로 검색 → 없으면 ② 상속 기준(후 6개월) → 없으면 "-"
'   실거래가액   : 증여 기준(후 3개월)
Private Const EVAL_MONTHS_BEFORE As Long = 6
Private Const EVAL_MONTHS_AFTER As Long = 3          ' 증여
Private Const EVAL_MONTHS_AFTER_INH As Long = 6      ' 상속

Private Const URL_JUSO As String = "https://business.juso.go.kr/addrlink/addrLinkApi.do"
Private Const URL_VWORLD As String = "https://api.vworld.kr/ned/data/getApartHousingPriceAttr"
Private Const URL_RTMS_DEFAULT As String = _
    "https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev"

' 설정값 (아파트_설정 시트에서 읽음)
Private cJusoKey As String
Private cRtmsKey As String
Private cVwKey As String
Private cVwDomain As String
Private cRtmsUrl As String
Private cMonthsBefore As Long
Private cMonthsAfter As Long
Private cPubMonth As Long
Private cPubDay As Long
Private cPubDates As Object                  ' 연도 → 실제 공시일 (설정 시트 E:F)
Private cRateTol As Double

' 단지 세대 레코드 : 0 동(정규화) / 1 호(정규화) / 2 층 / 3 전용면적 / 4 공동주택가격
' 거래 레코드     : 0 계약일 / 1 단지명 / 2 동(정규화) / 3 층 / 4 전용면적 / 5 거래금액(원)
'                   6 법정동코드(읍면동 5자리) / 7 지번
Private gUnitCache As Object
Private gTradeCache As Object
Private wsLog As Worksheet
Private gLogRow As Long

'==============================================================================
' 메인 실행 : [아파트] 시트에 결과 기록 (시트가 없으면 현재 활성 시트)
'==============================================================================
Public Sub 아파트가격조회()
    Dim ws As Worksheet
    Dim dateCols As Collection
    Dim col As Long, lastRow As Long, r As Long
    Dim sInput As String

    If Not LoadConfig() Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_APT)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = ActiveSheet
    If ws.Name = SHT_CFG Or ws.Name = SHT_LOG Then
        MsgBox "조회 시트를 선택한 후 실행하십시오.", vbExclamation
        Exit Sub
    End If

    ' 평가기준일 열 수집 (2행, 6열 간격)
    Set dateCols = New Collection
    col = FIRST_COL
    Do While Len(Trim$(CStr(ws.Cells(HDR_DATE_ROW, col).Value))) > 0
        If IsDate(ws.Cells(HDR_DATE_ROW, col).Value) Then dateCols.Add col
        col = col + GROUP_W
    Loop
    If dateCols.Count = 0 Then
        MsgBox "2행에 평가기준일이 없습니다.", vbExclamation
        Exit Sub
    End If

    Set gUnitCache = CreateObject("Scripting.Dictionary")
    Set gTradeCache = CreateObject("Scripting.Dictionary")
    InitLog

    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    On Error GoTo EH

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = FIRST_ROW To lastRow
        sInput = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(sInput) > 0 Then
            Application.StatusBar = "조회 중: " & r & "행 / " & sInput
            ProcessRow ws, r, sInput, dateCols         ' A열이 빈 행은 조회하지 않음
            DoEvents
        End If
    Next r

Finish:
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    wsLog.Columns("A:L").AutoFit
    MsgBox "조회 완료. 세부 근거는 [" & SHT_LOG & "] 시트를 확인하십시오.", vbInformation
    Exit Sub
EH:
    MsgBox "오류(" & r & "행): " & Err.Description, vbCritical
    Resume Finish
End Sub

'------------------------------------------------------------------------------
' 한 행 처리
'------------------------------------------------------------------------------
Private Sub ProcessRow(ws As Worksheet, ByVal r As Long, ByVal sInput As String, dateCols As Collection)
    Dim kwAddr As String, kwAll As String, dongIn As String, hoIn As String
    Dim addr As Object, msg As String
    Dim pnu As String, lawd As String, umd As String, jibunKey As String
    Dim col As Variant, evalDate As Date

    If Not ParseInput(sInput, kwAddr, kwAll, dongIn, hoIn) Then
        WriteAllNA ws, r, dateCols
        AddLog r, sInput, "", "입력오류", "동·호수를 인식할 수 없음 (예: 105동 1906호)"
        Exit Sub
    End If

    Set addr = JusoLookup(kwAddr, msg)
    If addr Is Nothing And kwAll <> kwAddr Then Set addr = JusoLookup(kwAll, msg)
    If addr Is Nothing Then
        WriteAllNA ws, r, dateCols
        AddLog r, sInput, "", "주소오류", msg
        Exit Sub
    End If

    ' PNU = 법정동코드(10) + 산여부(1:일반, 2:산) + 본번(4) + 부번(4)
    pnu = addr("admCd") & IIf(addr("mtYn") = "1", "2", "1") & _
          Format$(Val(addr("lnbrMnnm")), "0000") & Format$(Val(addr("lnbrSlno")), "0000")
    lawd = Left$(addr("admCd"), 5)
    umd = Mid$(addr("admCd"), 6, 5)
    jibunKey = CStr(Val(addr("lnbrMnnm"))) & IIf(Val(addr("lnbrSlno")) > 0, "-" & CStr(Val(addr("lnbrSlno"))), "")

    AddLog r, sInput, "", "주소확인", addr("roadAddr") & " / " & addr("jibunAddr") & _
           " / 건물명:" & addr("bdNm") & " / PNU:" & pnu & " / " & dongIn & "동 " & hoIn & "호"

    For Each col In dateCols
        evalDate = CDate(ws.Cells(HDR_DATE_ROW, CLng(col)).Value)
        ProcessDate ws, r, CLng(col), evalDate, sInput, pnu, lawd, umd, jibunKey, CStr(addr("bdNm")), dongIn, hoIn
    Next col
End Sub

'------------------------------------------------------------------------------
' 평가기준일 1개 처리
'   col   : 공시일      col+1 : 기준시가
'   col+2 : 고시일자    col+3 : 매매사례가액
'   col+4 : 매매계약일  col+5 : 실거래가액
'------------------------------------------------------------------------------
Private Sub ProcessDate(ws As Worksheet, ByVal r As Long, ByVal col As Long, ByVal evalDate As Date, _
                        ByVal sInput As String, ByVal pnu As String, ByVal lawd As String, _
                        ByVal umd As String, ByVal jibunKey As String, ByVal bdNm As String, _
                        ByVal dongIn As String, ByVal hoIn As String)
    Dim yr As Long, msg As String
    Dim units As Collection, trades As Collection
    Dim subj As Variant, t As Variant
    Dim dtFrom As Date, dtTo As Date, dtToInh As Date
    Dim bestOwn As Variant, bestSim As Variant
    Dim bestDiff As Double, basis As String

    yr = PubYear(evalDate)

    '--- 1) 기준시가(공동주택가격)
    Set units = GetComplexUnits(pnu, yr, msg)
    subj = FindUnit(units, dongIn, hoIn)
    If IsEmpty(subj) Then
        WriteGroupNA ws, r, col
        AddLog r, sInput, evalDate, "기준시가", yr & "년 공동주택가격에서 해당 동·호 미검색 " & _
               "(단지 PNU 상이 가능) " & msg
        Exit Sub
    End If
    PutDate ws.Cells(r, col), PubDate(yr)
    PutValue ws.Cells(r, col + 1), CDbl(subj(4))
    AddLog r, sInput, evalDate, "기준시가", yr & "년 공동주택가격", , subj(0), subj(2), subj(3), , subj(4)

    '--- 2) 거래자료 : 평가기준일 전 6개월 ~ 후 6개월(상속 기준)까지 한 번에 수집
    dtFrom = DateAdd("m", -cMonthsBefore, evalDate)
    dtTo = DateAdd("m", cMonthsAfter, evalDate)                 ' 증여 기준 종료일
    dtToInh = DateAdd("m", EVAL_MONTHS_AFTER_INH, evalDate)     ' 상속 기준 종료일
    Set trades = GetComplexTrades(lawd, umd, jibunKey, bdNm, dtFrom, dtToInh, msg)
    If Len(msg) > 0 Then AddLog r, sInput, evalDate, "실거래API", msg

    '--- 3) 실거래가액 : 증여 평가기간 내 동·층·면적 일치, 평가기준일에 가장 가까운 거래
    For Each t In trades
        If t(0) <= dtTo And IsSameUnit(t, subj) Then
            AddLog r, sInput, evalDate, "실거래(후보)", "동·층·면적 일치", t(0), t(2), t(3), t(4), t(5)
            If IsEmpty(bestOwn) Then
                bestOwn = t
            ElseIf Abs(t(0) - evalDate) < Abs(bestOwn(0) - evalDate) Then
                bestOwn = t
            End If
        End If
    Next t
    If IsEmpty(bestOwn) Then
        PutDash ws.Cells(r, col + 4)
        PutDash ws.Cells(r, col + 5)
    Else
        PutDate ws.Cells(r, col + 4), CDate(bestOwn(0))
        PutValue ws.Cells(r, col + 5), CDbl(bestOwn(5))
        AddLog r, sInput, evalDate, "실거래(채택)", "평가기준일 최근접", bestOwn(0), bestOwn(2), bestOwn(3), bestOwn(4), bestOwn(5)
    End If

    '--- 4) 유사매매사례가액 : 홈택스 "(유사)매매 시가(평가기간 내 유사 물건)" 기준
    '       ① 증여 평가기간(전 6개월 ~ 후 3개월) → ② 없으면 상속 평가기간(전 6개월 ~ 후 6개월) → ③ 없으면 "-"
    basis = "증여"
    bestSim = FindSimilar(trades, units, subj, evalDate, dtTo, r, sInput, basis, bestDiff)
    If IsEmpty(bestSim) Then
        basis = "상속"
        AddLog r, sInput, evalDate, "유사사례", "증여 평가기간 내 사례 없음 → 상속 평가기간(후 6개월)으로 재검색"
        bestSim = FindSimilar(trades, units, subj, evalDate, dtToInh, r, sInput, basis, bestDiff)
    End If

    ClearNote ws.Cells(r, col + 3)
    If IsEmpty(bestSim) Then
        PutDash ws.Cells(r, col + 2)
        PutDash ws.Cells(r, col + 3)
        AddLog r, sInput, evalDate, "유사사례", "증여·상속 평가기간 모두 사례 없음 → ""-"" 표시"
    Else
        ' 고시일자 : 비교에 사용한 공동주택가격의 고시일 (홈택스 화면의 '고시일자')
        PutDate ws.Cells(r, col + 2), PubDate(yr)
        PutValue ws.Cells(r, col + 3), CDbl(bestSim(5))
        If basis = "상속" Then AddNote ws.Cells(r, col + 3), "상속 평가기간(후 6개월) 기준 사례"
        AddLog r, sInput, evalDate, "유사사례(채택)", basis & " 기준, 공시가격 차이 최소", bestSim(0), bestSim(2), _
               bestSim(3), bestSim(4), bestSim(5), , bestDiff
    End If
End Sub

' 평가기간(dtFrom은 수집 시 반영, 종료일 dtEnd) 내 유사매매사례 선택
'   면적·공동주택가격 차이 5% 이내 → 공시가격 차이 최소 → 동률 시 평가기준일 최근접
Private Function FindSimilar(trades As Collection, units As Collection, subj As Variant, _
                             ByVal evalDate As Date, ByVal dtEnd As Date, ByVal r As Long, _
                             ByVal sInput As String, ByVal basis As String, ByRef bestDiff As Double) As Variant
    Dim t As Variant, best As Variant
    Dim cp As Double, diff As Double, areaDiff As Double

    bestDiff = 99
    For Each t In trades
        If t(0) <= dtEnd Then
            areaDiff = Abs(t(4) - subj(3)) / subj(3)
            If areaDiff <= cRateTol Then
                cp = AvgPubPrice(units, t)
                If cp > 0 Then
                    diff = Abs(cp - subj(4)) / subj(4)
                    If diff <= cRateTol Then
                        AddLog r, sInput, evalDate, "유사사례(" & basis & " 후보)", "면적차 " & Format$(areaDiff, "0.00%"), _
                               t(0), t(2), t(3), t(4), t(5), cp, diff
                        If diff < bestDiff - 0.0000001 Then
                            bestDiff = diff: best = t
                        ElseIf Abs(diff - bestDiff) <= 0.0000001 Then
                            If Abs(t(0) - evalDate) < Abs(best(0) - evalDate) Then best = t
                        End If
                    End If
                Else
                    AddLog r, sInput, evalDate, "유사사례(제외)", "비교세대 공동주택가격 미매칭", _
                           t(0), t(2), t(3), t(4), t(5)
                End If
            End If
        End If
    Next t
    FindSimilar = best
End Function

'==============================================================================
' 입력 해석 : "주소, 단지명, 105동 1906호" → 주소키워드 / 전체키워드 / 동 / 호
'==============================================================================
Private Function ParseInput(ByVal s As String, ByRef kwAddr As String, ByRef kwAll As String, _
                            ByRef dongIn As String, ByRef hoIn As String) As Boolean
    Dim re As Object, m As Object, rest As String, parts() As String

    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    ' "105동 1906호" (오타 "도"도 허용)
    re.Pattern = "([0-9A-Za-z]+)\s*[동도]\s*,?\s*([0-9A-Za-z]+)\s*호"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
    Else
        ' "105-1906" 형식
        re.Pattern = "([0-9]+)\s*-\s*([0-9]+)\s*호?\s*$"
        If Not re.Test(s) Then Exit Function
        Set m = re.Execute(s)(0)
    End If
    dongIn = m.SubMatches(0)
    hoIn = m.SubMatches(1)

    rest = Trim$(Replace(s, m.Value, ""))
    Do While Right$(rest, 1) = ","
        rest = Trim$(Left$(rest, Len(rest) - 1))
    Loop
    parts = Split(rest, ",")
    kwAddr = Trim$(parts(0))
    kwAll = Application.WorksheetFunction.Trim(Replace(rest, ",", " "))
    ParseInput = (Len(kwAddr) > 0)
End Function

'==============================================================================
' 도로명주소 API
'==============================================================================
Private Function JusoLookup(ByVal kw As String, ByRef msg As String) As Object
    Dim dom As Object, n As Object, d As Object, f As Variant, url As String

    url = URL_JUSO & "?confmKey=" & cJusoKey & "&currentPage=1&countPerPage=5&keyword=" & _
          Application.WorksheetFunction.EncodeURL(kw)
    Set dom = HttpGetXml(url, msg)
    If dom Is Nothing Then Exit Function

    If NodeText(dom, "//common/errorCode") <> "0" Then
        msg = "주소API: " & NodeText(dom, "//common/errorMessage")
        Exit Function
    End If
    Set n = dom.selectSingleNode("//juso")
    If n Is Nothing Then
        msg = "주소 검색결과 없음: " & kw
        Exit Function
    End If

    Set d = CreateObject("Scripting.Dictionary")
    For Each f In Array("roadAddr", "jibunAddr", "admCd", "mtYn", "lnbrMnnm", "lnbrSlno", "bdNm")
        d(f) = NodeText(n, CStr(f))
    Next f
    Set JusoLookup = d
End Function

'==============================================================================
' VWorld 공동주택가격 : PNU·기준연도 단위로 단지 전체 세대 조회 (캐시)
'==============================================================================
Private Function GetComplexUnits(ByVal pnu As String, ByVal yr As Long, ByRef msg As String) As Collection
    Dim key As String, col As Collection, dom As Object, nodes As Object, n As Object
    Dim pageNo As Long, url As String, rec(0 To 4) As Variant, hoTxt As String, flTxt As String

    key = pnu & "|" & yr
    If gUnitCache.Exists(key) Then
        Set GetComplexUnits = gUnitCache(key)
        Exit Function
    End If

    Set col = New Collection
    pageNo = 1
    Do
        url = URL_VWORLD & "?key=" & cVwKey & "&domain=" & cVwDomain & "&pnu=" & pnu & _
              "&stdrYear=" & yr & "&format=xml&numOfRows=1000&pageNo=" & pageNo
        Set dom = HttpGetXml(url, msg)
        If dom Is Nothing Then Exit Do
        Set nodes = dom.selectNodes("//*[pblntfPc]")
        If nodes.Length = 0 Then
            If pageNo = 1 Then msg = "VWorld 응답: " & Left$(dom.XML, 200)
            Exit Do
        End If
        For Each n In nodes
            hoTxt = NodeText(n, "hoNm")
            flTxt = NumOnly(NodeText(n, "floorNm"))
            rec(0) = NormDH(NodeText(n, "dongNm"))
            rec(1) = NormDH(hoTxt)
            ' 층 정보가 없으면 호수에서 추정 (1906호 → 19층)
            If Len(flTxt) > 0 Then rec(2) = CLng(flTxt) Else rec(2) = CLng(Val(NumOnly(hoTxt)) \ 100)
            rec(3) = Val(NodeText(n, "prvuseAr"))
            rec(4) = Val(Replace(NodeText(n, "pblntfPc"), ",", ""))
            col.Add rec
        Next n
        If nodes.Length < 1000 Then Exit Do
        pageNo = pageNo + 1
    Loop

    If col.Count > 0 Then gUnitCache.Add key, col
    Set GetComplexUnits = col
End Function

Private Function FindUnit(units As Collection, ByVal dongIn As String, ByVal hoIn As String) As Variant
    Dim u As Variant, dN As String, hN As String
    dN = NormDH(dongIn): hN = NormDH(hoIn)
    For Each u In units
        If u(0) = dN And u(1) = hN Then
            FindUnit = u
            Exit Function
        End If
    Next u
End Function

' 비교대상 거래의 공동주택가격 : 동·층·면적이 같은 세대 공시가격 평균 (호 미공개)
Private Function AvgPubPrice(units As Collection, t As Variant) As Double
    Dim u As Variant, s As Double, k As Long
    For Each u In units
        If u(2) = t(3) And Abs(u(3) - t(4)) <= AREA_TOL Then
            If Len(t(2)) = 0 Or u(0) = t(2) Then
                s = s + u(4): k = k + 1
            End If
        End If
    Next u
    If k > 0 Then AvgPubPrice = s / k
End Function

Private Function IsSameUnit(t As Variant, subj As Variant) As Boolean
    ' 실거래자료에 동 정보가 없는 거래는 동일 호로 보지 않음
    IsSameUnit = (Len(t(2)) > 0 And t(2) = subj(0) And t(3) = subj(2) And Abs(t(4) - subj(3)) <= AREA_TOL)
End Function

'==============================================================================
' 국토부 아파트 매매 실거래가 : 기간 내 해당 단지 거래 (해제거래 제외)
'==============================================================================
Private Function GetComplexTrades(ByVal lawd As String, ByVal umd As String, ByVal jibunKey As String, _
                                  ByVal bdNm As String, ByVal dtFrom As Date, ByVal dtTo As Date, _
                                  ByRef msg As String) As Collection
    Dim res As Collection, mon As Collection, t As Variant, d As Date, bdN As String

    Set res = New Collection
    bdN = NormName(bdNm)
    d = DateSerial(Year(dtFrom), Month(dtFrom), 1)
    Do While d <= dtTo And d <= Date
        Set mon = GetMonthTrades(lawd, Format$(d, "yyyymm"), msg)
        For Each t In mon
            If t(0) >= dtFrom And t(0) <= dtTo Then
                If t(6) = umd And (t(7) = jibunKey Or (Len(bdN) > 0 And NormName(t(1)) = bdN)) Then
                    res.Add t
                End If
            End If
        Next t
        d = DateAdd("m", 1, d)
    Loop
    Set GetComplexTrades = res
End Function

Private Function GetMonthTrades(ByVal lawd As String, ByVal ym As String, ByRef msg As String) As Collection
    Dim key As String, col As Collection, dom As Object, nodes As Object, n As Object
    Dim pageNo As Long, url As String, code As String, rec(0 To 7) As Variant, ok As Boolean

    key = lawd & "|" & ym
    If gTradeCache.Exists(key) Then
        Set GetMonthTrades = gTradeCache(key)
        Exit Function
    End If

    Set col = New Collection
    ok = True
    pageNo = 1
    Do
        url = cRtmsUrl & "?serviceKey=" & cRtmsKey & "&LAWD_CD=" & lawd & "&DEAL_YMD=" & ym & _
              "&pageNo=" & pageNo & "&numOfRows=1000"
        Set dom = HttpGetXml(url, msg)
        If dom Is Nothing Then ok = False: Exit Do
        code = NodeText(dom, "//resultCode")
        If code <> "" And code <> "00" And code <> "000" Then
            msg = "실거래API(" & ym & "): " & code & " " & NodeText(dom, "//resultMsg")
            ok = False: Exit Do
        End If
        Set nodes = dom.selectNodes("//item")
        For Each n In nodes
            ' 해제(취소) 거래 제외
            If UCase$(Trim$(NodeText(n, "cdealType"))) <> "O" Then
                rec(0) = DateSerial(Val(NodeText(n, "dealYear")), Val(NodeText(n, "dealMonth")), Val(NodeText(n, "dealDay")))
                rec(1) = NodeText(n, "aptNm")
                rec(2) = NormDH(NodeText(n, "aptDong"))
                rec(3) = CLng(Val(NodeText(n, "floor")))
                rec(4) = Val(NodeText(n, "excluUseAr"))
                rec(5) = Val(Replace(NodeText(n, "dealAmount"), ",", "")) * 10000#   ' 만원 → 원
                rec(6) = Trim$(NodeText(n, "umdCd"))
                rec(7) = Trim$(NodeText(n, "jibun"))
                col.Add rec
            End If
        Next n
        If nodes.Length < 1000 Then Exit Do
        pageNo = pageNo + 1
    Loop

    If ok Then gTradeCache.Add key, col     ' 오류 응답은 캐시하지 않음
    Set GetMonthTrades = col
End Function

'==============================================================================
' 공통 유틸
'==============================================================================
Private Function HttpGetXml(ByVal url As String, ByRef msg As String) As Object
    Dim http As Object, dom As Object
    On Error GoTo EH
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 10000, 30000, 60000
    http.Open "GET", url, False
    http.send
    If http.Status <> 200 Then
        msg = "HTTP " & http.Status & " " & Left$(url, InStr(url & "?", "?") - 1)
        Exit Function
    End If
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.SetProperty "SelectionLanguage", "XPath"
    If Not dom.Load(http.responseBody) Then
        msg = "XML 해석 실패: " & Left$(http.responseText, 200)
        Exit Function
    End If
    Set HttpGetXml = dom
    Exit Function
EH:
    msg = "통신오류: " & Err.Description
End Function

Private Function NodeText(ByVal parent As Object, ByVal xp As String) As String
    Dim n As Object
    Set n = parent.selectSingleNode(xp)
    If Not n Is Nothing Then NodeText = Trim$(n.Text)
End Function

' 동·호 정규화 : "제105동" → "105", "1906호" → "1906", "A동" → "A"
Private Function NormDH(ByVal s As String) As String
    s = UCase$(Replace(Trim$(s), " ", ""))
    If Left$(s, 1) = "제" Then s = Mid$(s, 2)
    If Right$(s, 1) = "동" Or Right$(s, 1) = "호" Then s = Left$(s, Len(s) - 1)
    If Len(s) > 0 And s Like String$(Len(s), "#") Then s = CStr(CLng(s))
    NormDH = s
End Function

Private Function NormName(ByVal s As String) As String
    NormName = UCase$(Replace(Replace(Replace(Trim$(s), " ", ""), "(", ""), ")", ""))
End Function

Private Function NumOnly(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then NumOnly = NumOnly & ch
    Next i
End Function

' 평가기준일 현재 공시된 공동주택가격의 기준연도
Private Function PubYear(ByVal d As Date) As Long
    If d >= PubDate(Year(d)) Then PubYear = Year(d) Else PubYear = Year(d) - 1
End Function

' 연도별 공동주택가격 공시일 : 설정 시트 E:F 표 우선, 없으면 기본 월·일
Private Function PubDate(ByVal yr As Long) As Date
    If cPubDates.Exists(yr) Then
        PubDate = cPubDates(yr)
    Else
        PubDate = DateSerial(yr, cPubMonth, cPubDay)
    End If
End Function

Private Sub PutValue(ByVal c As Range, ByVal v As Double)
    If v > 0 Then
        c.Value = v
        c.NumberFormat = "#,##0"
        c.HorizontalAlignment = xlRight
    Else
        c.Value = "N/A"
        c.HorizontalAlignment = xlCenter
    End If
End Sub

Private Sub WriteAllNA(ws As Worksheet, ByVal r As Long, dateCols As Collection)
    Dim col As Variant
    For Each col In dateCols
        WriteGroupNA ws, r, CLng(col)
    Next col
End Sub

' 평가기준일 1개 그룹(6열) 전체 N/A
Private Sub WriteGroupNA(ws As Worksheet, ByVal r As Long, ByVal col As Long)
    Dim k As Long
    For k = 0 To GROUP_W - 1
        PutValue ws.Cells(r, col + k), 0
    Next k
End Sub

Private Sub PutDate(ByVal c As Range, ByVal d As Date)
    c.Value = d
    c.NumberFormat = "yyyy-mm-dd"
    c.HorizontalAlignment = xlCenter
End Sub

' 실거래가액·매매사례가액 없음 표시
Private Sub PutDash(ByVal c As Range)
    c.Value = "-"
    c.HorizontalAlignment = xlCenter
End Sub

' 셀 메모 (상속 기준 사례 표시용)
Private Sub AddNote(ByVal c As Range, ByVal s As String)
    ClearNote c
    c.AddComment s
End Sub

Private Sub ClearNote(ByVal c As Range)
    If Not c.Comment Is Nothing Then c.Comment.Delete
End Sub

'==============================================================================
' 설정 시트
'==============================================================================
Private Function LoadConfig() As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_CFG)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHT_CFG
        ws.Range("A1:C1").Value = Array("항목", "값", "비고")
        ws.Range("A2:C2").Value = Array("도로명주소 API 승인키", "", "business.juso.go.kr 검색API 신청")
        ws.Range("A3:C3").Value = Array("공공데이터포털 서비스키", "", "data.go.kr '아파트 매매 실거래가 상세 자료' 활용신청, 인코딩(Encoding) 키")
        ws.Range("A4:C4").Value = Array("VWorld 인증키", "", "vworld.kr 오픈API 인증키 (공동주택가격속성조회)")
        ws.Range("A5:C5").Value = Array("VWorld 도메인", "", "인증키 발급 시 등록한 서비스 URL")
        ws.Range("A6:C6").Value = Array("평가기간(전)", "6개월", "증여 기준 고정 (코드 상수)")
        ws.Range("A7:C7").Value = Array("평가기간(후)", "3개월", "증여 기준 고정 (상속은 6개월)")
        ws.Range("A8:C8").Value = Array("공동주택가격 공시월", 4, "공시일 이전 평가기준일은 직전연도 가격 적용")
        ws.Range("A9:C9").Value = Array("공동주택가격 공시일", 30, "연도별 실제 공시일 확인 필요")
        ws.Range("A10:C10").Value = Array("유사재산 허용비율", 0.05, "면적·공시가격 차이 5% (상증칙 §15③)")
        ws.Range("A11:C11").Value = Array("실거래 API URL", URL_RTMS_DEFAULT, "엔드포인트 변경 시 수정")
        ws.Range("E1:G1").Value = Array("연도", "공동주택가격 공시일", "비고")
        ws.Range("E2:G2").Value = Array(2024, DateSerial(2024, 4, 30), "확인 필요")
        ws.Range("E3:G3").Value = Array(2025, DateSerial(2025, 4, 30), "확인 필요")
        ws.Range("E4:G4").Value = Array(2026, DateSerial(2026, 4, 30), "홈택스 고시일자로 확인")
        ws.Range("F2:F4").NumberFormat = "yyyy-mm-dd"
        ws.Range("A1:C1,E1:G1").Font.Bold = True
        ws.Columns("A:G").AutoFit
        MsgBox "[" & SHT_CFG & "] 시트를 생성했습니다. 인증키를 입력한 후 다시 실행하십시오.", vbInformation
        Exit Function
    End If

    cJusoKey = Trim$(CStr(ws.Range("B2").Value))
    cRtmsKey = Trim$(CStr(ws.Range("B3").Value))
    cVwKey = Trim$(CStr(ws.Range("B4").Value))
    cVwDomain = Trim$(CStr(ws.Range("B5").Value))
    cMonthsBefore = EVAL_MONTHS_BEFORE
    cMonthsAfter = EVAL_MONTHS_AFTER
    cPubMonth = CLng(Val(ws.Range("B8").Value))
    cPubDay = CLng(Val(ws.Range("B9").Value))
    cRateTol = CDbl(Val(ws.Range("B10").Value))
    cRtmsUrl = Trim$(CStr(ws.Range("B11").Value))
    If cRtmsUrl = "" Then cRtmsUrl = URL_RTMS_DEFAULT
    If cPubMonth < 1 Or cPubMonth > 12 Then cPubMonth = 4
    If cPubDay < 1 Or cPubDay > 31 Then cPubDay = 30
    ' 연도별 공시일 표 (기존 설정 시트에 없으면 추가)
    If Trim$(CStr(ws.Range("E1").Value)) = "" Then
        ws.Range("E1:G1").Value = Array("연도", "공동주택가격 공시일", "비고")
        ws.Range("E2:G2").Value = Array(2024, DateSerial(2024, 4, 30), "확인 필요")
        ws.Range("E3:G3").Value = Array(2025, DateSerial(2025, 4, 30), "확인 필요")
        ws.Range("E4:G4").Value = Array(2026, DateSerial(2026, 4, 30), "홈택스 고시일자로 확인")
        ws.Range("F2:F4").NumberFormat = "yyyy-mm-dd"
        ws.Range("E1:G1").Font.Bold = True
    End If
    Set cPubDates = LoadPubDates(ws)
    If cRateTol <= 0 Then cRateTol = 0.05

    If cJusoKey = "" Or cRtmsKey = "" Or cVwKey = "" Then
        MsgBox "[" & SHT_CFG & "] 시트 B2:B4에 인증키를 입력하십시오.", vbExclamation
        Exit Function
    End If
    LoadConfig = True
End Function

' 설정 시트 E:F (연도, 공시일) 표 읽기
Private Function LoadPubDates(ws As Worksheet) As Object
    Dim d As Object, r As Long
    Set d = CreateObject("Scripting.Dictionary")
    r = 2
    Do While Len(Trim$(CStr(ws.Cells(r, 5).Value))) > 0
        If IsNumeric(ws.Cells(r, 5).Value) And IsDate(ws.Cells(r, 6).Value) Then
            d(CLng(ws.Cells(r, 5).Value)) = CDate(ws.Cells(r, 6).Value)
        End If
        r = r + 1
    Loop
    Set LoadPubDates = d
End Function

'==============================================================================
' 조회로그 시트 (감사증거용 근거 기록)
'==============================================================================
Private Sub InitLog()
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(SHT_LOG)
    On Error GoTo 0
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLog.Name = SHT_LOG
    End If
    wsLog.Cells.Clear
    wsLog.Range("A1:L1").Value = Array("행", "입력", "평가기준일", "구분", "내용", "계약일", _
                                       "동", "층", "전용면적(㎡)", "거래금액(원)", "공동주택가격(원)", "공시가격 차이율")
    wsLog.Range("A1:L1").Font.Bold = True
    wsLog.Range("A1:L1").HorizontalAlignment = xlCenter
    gLogRow = 2
End Sub

Private Sub AddLog(ByVal r As Long, ByVal sInput As String, ByVal evalDate As Variant, ByVal kind As String, _
                   ByVal note As String, Optional ByVal dealDate As Variant, Optional ByVal dong As Variant, _
                   Optional ByVal fl As Variant, Optional ByVal area As Variant, Optional ByVal amt As Variant, _
                   Optional ByVal pub As Variant, Optional ByVal rate As Variant)
    With wsLog
        .Cells(gLogRow, 1).Value = r
        .Cells(gLogRow, 2).Value = sInput
        If IsDate(evalDate) Then .Cells(gLogRow, 3).Value = CDate(evalDate): .Cells(gLogRow, 3).NumberFormat = "yyyy-mm-dd"
        .Cells(gLogRow, 4).Value = kind
        .Cells(gLogRow, 5).Value = note
        If Not IsMissing(dealDate) Then .Cells(gLogRow, 6).Value = dealDate: .Cells(gLogRow, 6).NumberFormat = "yyyy-mm-dd"
        If Not IsMissing(dong) Then .Cells(gLogRow, 7).Value = "'" & dong
        If Not IsMissing(fl) Then .Cells(gLogRow, 8).Value = fl
        If Not IsMissing(area) Then .Cells(gLogRow, 9).Value = area: .Cells(gLogRow, 9).NumberFormat = "0.00"
        If Not IsMissing(amt) Then .Cells(gLogRow, 10).Value = amt: .Cells(gLogRow, 10).NumberFormat = "#,##0"
        If Not IsMissing(pub) Then .Cells(gLogRow, 11).Value = pub: .Cells(gLogRow, 11).NumberFormat = "#,##0"
        If Not IsMissing(rate) Then .Cells(gLogRow, 12).Value = rate: .Cells(gLogRow, 12).NumberFormat = "0.00%"
    End With
    gLogRow = gLogRow + 1
End Sub
' ===== 아파트가격 코드 끝 (총 729줄) =====
