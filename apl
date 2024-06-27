Sub GetSmartStoreOrders()
    Dim http As Object
    Dim json As Object
    Dim apiUrl As String
    Dim apiKey As String
    Dim apiSecret As String
    Dim apiAccessKey As String
    Dim timestamp As String
    Dim signature As String
    Dim ws As Worksheet
    Dim i As Integer
    Dim signData As String

    ' API URL 및 키 설정
    apiKey = "ncp_1nym9t_01" ' 여기에 네이버 스마트스토어 API 키를 입력하세요
    apiSecret = "your_api_secret" ' 여기에 네이버 스마트스토어 API 비밀 키를 입력하세요
    apiAccessKey = "your_access_key" ' 여기에 네이버 스마트스토어 API 접근 키를 입력하세요
    apiUrl = "https://api.commerce.naver.com/orders/v2/list" ' 네이버 스마트스토어 주문 목록 API URL을 입력하세요

    ' 현재 타임스탬프 생성
    timestamp = CStr(1000 * (Now() - #1970-01-01#) * 86400)

    ' 서명 데이터 생성
    signData = "GET " & apiUrl & vbLf & timestamp & vbLf & apiSecret

    ' 서명 생성
    signature = HMAC_SHA256(apiSecret, signData)

    ' "Orders" 시트 설정, 없으면 시트 추가
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Orders")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Orders"
    End If
    On Error GoTo 0
    ws.Cells.Clear

    ' HTTP 요청
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.Open "GET", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "X-NCP-APIGW-TIMESTAMP", timestamp
    http.setRequestHeader "X-NCP-APIGW-API-KEY", apiKey
    http.setRequestHeader "X-NCP-APIGW-SIGNATURE-V2", signature
    http.setRequestHeader "X-NCP-APIGW-ACCESS-KEY", apiAccessKey
    http.Send

    ' 응답 상태 확인
    If http.Status = 200 Then
        ' JSON 파싱
        Set json = JsonConverter.ParseJson(http.responseText)

        ' 데이터 엑셀에 기록
        i = 2
        ws.Cells(1, 1).Value = "Order ID"
        ws.Cells(1, 2).Value = "Customer Name"
        ws.Cells(1, 3).Value = "Total Price"
        For Each order In json("orders")
            ws.Cells(i, 1).Value = order("orderId")
            ws.Cells(i, 2).Value = order("customerName")
            ws.Cells(i, 3).Value = order("totalPayment")
            i = i + 1
        Next order

        MsgBox "주문 데이터가 성공적으로 업데이트되었습니다."
    Else
        MsgBox "서버 응답 오류: " & http.Status & " - " & http.StatusText
    End If
End Sub
