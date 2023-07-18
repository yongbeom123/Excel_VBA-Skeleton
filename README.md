# Excel_VBA-Skeleton

Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub GetTickData()
    Dim kiwoom As Object
    Set kiwoom = CreateObject("KHOPENAPI.KHOpenAPICtrl.1")
    
    ' 키움증권 API 초기화
    kiwoom.CommConnect
    
    ' 로그인 대기 (실제로는 이벤트를 활용하여 로그인 확인하는 것이 좋음)
    Sleep 3000
    
    ' 원하는 종목 코드 설정 (ex: 삼성전자 종목 코드: "005930")
    Dim stockCode As String
    stockCode = "YOUR_STOCK_CODE"
    
    ' 틱 데이터 수신 요청
    kiwoom.SetInputValue "종목코드", stockCode
    kiwoom.SetInputValue "조회일자", Format(Date, "yyyyMMdd")
    kiwoom.SetInputValue "시간단위", "1"
    kiwoom.SetInputValue "수정주가구분", "0"
    kiwoom.CommRqData "주식틱차트조회", "OPT10079", 0, "0101"
    
    ' 데이터 수신 대기 (실제로는 이벤트를 활용하여 데이터 수신 확인하는 것이 좋음)
    Sleep 5000
    
    ' 데이터 수신 및 기록
    Dim rowCount As Integer
    rowCount = kiwoom.GetRepeatCnt("주식틱차트조회")
    
    Dim i As Integer
    For i = 0 To rowCount - 1
        Dim timeData As String
        Dim priceData As Double
        
        timeData = kiwoom.GetCommData("주식틱차트조회", "체결시간", i)
        priceData = kiwoom.GetCommData("주식틱차트조회", "현재가", i)
        
        ' 데이터를 원하는 시트에 기록
        Sheets("틱데이터").Cells(i + 2, 1).Value = timeData
        Sheets("틱데이터").Cells(i + 2, 2).Value = priceData
    Next i
    
    ' 키움증권 API 종료
    kiwoom.CommTerminate
End Sub
