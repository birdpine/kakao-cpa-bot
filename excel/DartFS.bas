Attribute VB_Name = "DartFS"
'==============================================================================
' DART 재무제표 가져오기 (연도별 BS·PL 시트)
'
'  사용법
'   1. 이 모듈을 매크로 사용 통합문서(.xlsm)에 가져오기(Import)
'   2. 최초 1회 SetupDartSheet 실행 → '조회' 시트와 [재무제표 가져오기] 버튼 생성
'   3. '조회' 시트에 회사명·시작연도·종료연도·구분 입력 후 버튼 클릭
'   4. BS(2025), PL(2025), BS(2024), PL(2024) ... 시트와 '조회정보' 시트가 생성됨
'      (다시 조회하면 기존 BS(*), PL(*), 조회정보 시트는 삭제 후 새로 생성)
'
'  동작 방식
'   - DART 재무제표 서버(/dart/api/annual)를 호출해 엑셀 파일을 받은 뒤
'     그 시트들을 이 통합문서로 복사합니다. (감사보고서 원문 파싱은 서버에서 수행)
'   - 동명 회사가 여러 개면 후보 목록(고유번호)을 시트에 표시 →
'     고유번호를 입력하고 다시 실행
'
'  요구사항: Windows용 Excel 2013 이상 (WinHTTP, WorksheetFunction.EncodeURL 사용)
'==============================================================================
Option Explicit

' --- '조회' 시트 입력 위치 ---------------------------------------------------
Private Const INPUT_SHEET As String = "조회"
Private Const CELL_NAME As String = "B3"      ' 회사명
Private Const CELL_FROM As String = "B4"      ' 시작연도
Private Const CELL_TO As String = "B5"        ' 종료연도
Private Const CELL_FS As String = "B6"        ' 구분 (별도/연결)
Private Const CELL_CODE As String = "B7"      ' 고유번호 (동명 회사일 때만)
Private Const CELL_SERVER As String = "B9"    ' 서버 주소
Private Const CELL_TOKEN As String = "B10"    ' 접근 토큰 (서버에 설정한 경우만)
Private Const CELL_STATUS As String = "B12"   ' 진행상태 표시
Private Const CAND_HEADER_ROW As Long = 14    ' 동명 회사 후보 목록 시작 행

' --- 기타 설정 ---------------------------------------------------------------
Private Const INFO_SHEET As String = "조회정보"            ' 서버 엑셀의 '정보' 시트를 이 이름으로 복사
Private Const DEFAULT_SERVER As String = "http://localhost:8000"
Private Const MAX_YEARS As Long = 15
Private Const RECEIVE_TIMEOUT_MS As Long = 300000           ' 연도가 많으면 수 분 걸릴 수 있음

'------------------------------------------------------------------------------
' 최초 1회 실행: '조회' 시트 양식과 버튼 생성 (이미 입력한 값은 유지)
'------------------------------------------------------------------------------
Public Sub SetupDartSheet()
    Dim ws As Worksheet
    Dim btn As Object
    Dim labels As Variant
    Dim i As Long

    Set ws = GetSheet(INPUT_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = INPUT_SHEET
    End If

    With ws
        .Range("A1").Value = "DART 재무제표 조회 (연도별 BS·PL)"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        labels = Array("회사명", "시작연도", "종료연도", "구분", "고유번호(선택)", "", "서버주소", "접근토큰(선택)", "", "상태")
        For i = LBound(labels) To UBound(labels)
            .Cells(3 + i, 1).Value = labels(i)
            .Cells(3 + i, 1).Font.Bold = True
        Next i

        ' 기본값 (비어 있을 때만)
        If IsEmpty(.Range(CELL_FROM).Value) Then .Range(CELL_FROM).Value = Year(Date) - 3
        If IsEmpty(.Range(CELL_TO).Value) Then .Range(CELL_TO).Value = Year(Date) - 1
        If IsEmpty(.Range(CELL_FS).Value) Then .Range(CELL_FS).Value = "별도"
        If IsEmpty(.Range(CELL_SERVER).Value) Then .Range(CELL_SERVER).Value = DEFAULT_SERVER

        ' 구분: 목록 선택
        With .Range(CELL_FS).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="별도,연결"
        End With

        ' 고유번호는 앞자리 0 보존을 위해 텍스트 서식
        .Range(CELL_CODE).NumberFormat = "@"

        ' 입력칸 표시
        With .Range(CELL_NAME & ":" & CELL_CODE & "," & CELL_SERVER & ":" & CELL_TOKEN)
            .Interior.Color = RGB(255, 249, 219)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(191, 191, 191)
        End With

        .Range("D7").Value = "※ 동명 회사가 여러 개일 때만 아래 목록의 고유번호(8자리)를 입력하세요. 평소에는 비워 둡니다."
        .Range("D8").Value = "※ 연도 = 사업연도 종료연도 (예: 2024년 12월 결산 → 2024)"
        .Range("D9").Value = "※ 서버를 내 PC에서 실행하면 http://localhost:8000, 배포했으면 그 주소(https://...)"
        .Range("D7:D9").Font.Color = RGB(102, 102, 102)

        .Columns("A").ColumnWidth = 16
        .Columns("B").ColumnWidth = 34
        .Columns("C").ColumnWidth = 14

        ' 버튼 (기존 버튼은 지우고 다시 생성)
        If .Buttons.Count > 0 Then .Buttons.Delete
        Set btn = .Buttons.Add(.Range("D3").Left, .Range("D3").Top, 160, .Range("D3:D5").Height)
        btn.Caption = "재무제표 가져오기"
        btn.OnAction = "GetDartStatements"
    End With

    ws.Activate
    ws.Range(CELL_NAME).Select
End Sub

'------------------------------------------------------------------------------
' [재무제표 가져오기] 버튼
'------------------------------------------------------------------------------
Public Sub GetDartStatements()
    Dim ws As Worksheet
    Dim corpName As String
    Dim corpCode As String
    Dim serverUrl As String
    Dim token As String
    Dim fsDiv As String
    Dim years As String
    Dim url As String
    Dim yearFrom As Long
    Dim yearTo As Long
    Dim yearCount As Long
    Dim y As Long
    Dim http As Object
    Dim body As String
    Dim tempPath As String
    Dim sheetCount As Long

    Set ws = GetSheet(INPUT_SHEET)
    If ws Is Nothing Then
        MsgBox "'" & INPUT_SHEET & "' 시트가 없습니다. SetupDartSheet 매크로를 먼저 실행하세요.", vbExclamation
        Exit Sub
    End If
    ClearCandidates ws

    ' --- 입력값 검증 ---------------------------------------------------------
    corpName = Trim$(CStr(ws.Range(CELL_NAME).Value))
    corpCode = NormalizeCorpCode(ws.Range(CELL_CODE).Value)
    serverUrl = Trim$(CStr(ws.Range(CELL_SERVER).Value))
    token = Trim$(CStr(ws.Range(CELL_TOKEN).Value))

    If corpName = "" And corpCode = "" Then
        MsgBox "회사명을 입력하세요.", vbExclamation
        Exit Sub
    End If
    If corpCode <> "" And Not corpCode Like "########" Then
        MsgBox "고유번호는 숫자 8자리입니다. (예: 00126380)", vbExclamation
        Exit Sub
    End If
    If Not IsYear(ws.Range(CELL_FROM).Value) Or Not IsYear(ws.Range(CELL_TO).Value) Then
        MsgBox "시작연도·종료연도를 4자리 숫자로 입력하세요. (예: 2022, 2024)", vbExclamation
        Exit Sub
    End If
    If serverUrl = "" Then serverUrl = DEFAULT_SERVER
    Do While Right$(serverUrl, 1) = "/"
        serverUrl = Left$(serverUrl, Len(serverUrl) - 1)
    Loop

    yearFrom = CLng(ws.Range(CELL_FROM).Value)
    yearTo = CLng(ws.Range(CELL_TO).Value)
    If yearFrom > yearTo Then
        y = yearFrom: yearFrom = yearTo: yearTo = y
    End If
    yearCount = yearTo - yearFrom + 1
    If yearCount > MAX_YEARS Then
        MsgBox "한 번에 " & MAX_YEARS & "개 연도까지 조회할 수 있습니다.", vbExclamation
        Exit Sub
    End If

    ' 최근 연도부터: "2024,2023,2022"
    For y = yearTo To yearFrom Step -1
        If years <> "" Then years = years & ","
        years = years & CStr(y)
    Next y

    fsDiv = IIf(Trim$(CStr(ws.Range(CELL_FS).Value)) = "연결", "CFS", "OFS")

    url = serverUrl & "/dart/api/annual?years=" & years & "&fs_div=" & fsDiv
    If corpCode <> "" Then
        url = url & "&corp_code=" & corpCode
    Else
        url = url & "&corp_name=" & Application.WorksheetFunction.EncodeURL(corpName)
    End If

    ' --- 서버 호출 -----------------------------------------------------------
    SetStatus ws, "DART 조회 중… (" & yearCount & "개 연도, 수십 초~수 분 걸릴 수 있습니다)"
    Application.Cursor = xlWait
    On Error GoTo Fail

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, RECEIVE_TIMEOUT_MS   ' resolve, connect, send, receive
    http.Open "GET", url, False
    If token <> "" Then http.SetRequestHeader "X-Access-Token", token
    http.Send

    Select Case http.Status
        Case 200
            tempPath = SaveTempFile(http.ResponseBody)
            sheetCount = ImportSheets(tempPath)
            Kill tempPath
            SetStatus ws, "완료 (" & Format$(Now, "yyyy-mm-dd hh:nn") & "): 시트 " & sheetCount & "개 생성. " & _
                          "'" & INFO_SHEET & "' 시트에서 연도별 출처와 확인사항을 보세요."
            ThisWorkbook.Worksheets(INFO_SHEET).Activate
        Case 409
            body = Utf8ToString(http.ResponseBody)
            ShowCandidates ws, body
            SetStatus ws, "동명 회사가 여러 개입니다. 아래 목록에서 고유번호를 " & CELL_CODE & "에 입력하고 다시 실행하세요."
        Case Else
            body = Utf8ToString(http.ResponseBody)
            SetStatus ws, "오류 (" & http.Status & "): " & ErrorDetail(body)
            MsgBox ErrorDetail(body), vbExclamation, "DART 재무제표 조회 오류 (" & http.Status & ")"
    End Select

Cleanup:
    Application.Cursor = xlDefault
    Exit Sub

Fail:
    SetStatus ws, "실패: " & Err.Description
    MsgBox "서버 호출에 실패했습니다." & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "서버주소(" & serverUrl & ")와 서버 실행 여부를 확인하세요.", vbExclamation
    Resume Cleanup
End Sub

'------------------------------------------------------------------------------
' 받은 엑셀의 시트를 이 통합문서 끝에 복사. 반환값: 복사한 시트 수
'------------------------------------------------------------------------------
Private Function ImportSheets(ByVal path As String) As Long
    Dim src As Workbook
    Dim sh As Worksheet
    Dim newName As String
    Dim copied As Long
    Dim errNumber As Long
    Dim errMessage As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Fail

    RemovePreviousResults

    Set src = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    For Each sh In src.Worksheets
        newName = IIf(sh.Name = "정보", INFO_SHEET, sh.Name)
        sh.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count).Name = newName
        copied = copied + 1
    Next sh
    src.Close SaveChanges:=False
    Set src = Nothing

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ImportSheets = copied
    Exit Function

Fail:
    errNumber = Err.Number              ' 정리 작업 전에 오류 정보 보관
    errMessage = Err.Description
    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise errNumber, , "시트 복사 중 오류: " & errMessage
End Function

'------------------------------------------------------------------------------
' 이전 조회 결과(BS(연도), PL(연도), 조회정보) 시트 삭제
'------------------------------------------------------------------------------
Private Sub RemovePreviousResults()
    Dim i As Long
    Dim nm As String

    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        nm = ThisWorkbook.Worksheets(i).Name
        If nm Like "BS(####)" Or nm Like "PL(####)" Or nm = INFO_SHEET Then
            ThisWorkbook.Worksheets(i).Delete
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' 동명 회사 후보 표시. 서버 응답: 행마다 "고유번호<TAB>회사명<TAB>구분"
'------------------------------------------------------------------------------
Private Sub ShowCandidates(ByVal ws As Worksheet, ByVal body As String)
    Dim lines As Variant
    Dim parts As Variant
    Dim r As Long
    Dim i As Long

    With ws
        .Cells(CAND_HEADER_ROW, 1).Resize(1, 3).Value = Array("고유번호", "회사명", "구분")
        With .Cells(CAND_HEADER_ROW, 1).Resize(1, 3)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(221, 235, 247)
        End With

        lines = Split(Replace(body, vbCr, ""), vbLf)
        r = CAND_HEADER_ROW + 1
        For i = LBound(lines) To UBound(lines)
            If Len(Trim$(lines(i))) > 0 Then
                parts = Split(lines(i), vbTab)
                .Cells(r, 1).NumberFormat = "@"         ' 고유번호 앞자리 0 보존
                .Cells(r, 1).Value = parts(0)
                If UBound(parts) >= 1 Then .Cells(r, 2).Value = parts(1)
                If UBound(parts) >= 2 Then .Cells(r, 3).Value = parts(2)
                r = r + 1
            End If
        Next i
    End With
End Sub

Private Sub ClearCandidates(ByVal ws As Worksheet)
    With ws.Range(ws.Cells(CAND_HEADER_ROW, 1), ws.Cells(ws.Rows.Count, 3))
        .ClearContents
        .Interior.Pattern = xlNone
    End With
End Sub

'------------------------------------------------------------------------------
' 보조 함수
'------------------------------------------------------------------------------
Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub SetStatus(ByVal ws As Worksheet, ByVal message As String)
    ws.Range(CELL_STATUS).Value = message
    DoEvents
End Sub

Private Function IsYear(ByVal v As Variant) As Boolean
    If IsNumeric(v) And Not IsEmpty(v) Then
        IsYear = (CDbl(v) = Int(CDbl(v)) And CDbl(v) >= 1999 And CDbl(v) <= Year(Date))
    End If
End Function

' 일반 서식 셀에 00126380을 입력하면 126380이 되므로 8자리로 복원
Private Function NormalizeCorpCode(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    If s <> "" And IsNumeric(s) And Len(s) < 8 Then s = Format$(CLng(s), "00000000")
    NormalizeCorpCode = s
End Function

' 응답 바이트를 임시 .xlsx 파일로 저장
Private Function SaveTempFile(ByVal bytes As Variant) As String
    Dim path As String
    Dim stream As Object

    path = Environ$("TEMP") & "\dart_fs_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1                 ' adTypeBinary
    stream.Open
    stream.Write bytes
    stream.SaveToFile path, 2       ' adSaveCreateOverWrite
    stream.Close
    SaveTempFile = path
End Function

' 서버 응답(UTF-8)을 문자열로
Private Function Utf8ToString(ByVal bytes As Variant) As String
    Dim stream As Object

    If IsEmpty(bytes) Then Exit Function
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1                 ' adTypeBinary
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2                 ' adTypeText
    stream.Charset = "utf-8"
    Utf8ToString = stream.ReadText
    stream.Close
End Function

' {"detail":"메시지"} 형식이면 메시지만 추출
Private Function ErrorDetail(ByVal body As String) As String
    Const KEY As String = """detail"":"""
    Dim p As Long
    Dim q As Long

    p = InStr(body, KEY)
    If p > 0 Then
        p = p + Len(KEY)
        q = InStr(p, body, """}")
        If q > p Then
            ErrorDetail = Replace(Mid$(body, p, q - p), "\""", """")
            Exit Function
        End If
    End If
    ErrorDetail = Left$(body, 300)
End Function
