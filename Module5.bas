Attribute VB_Name = "Module4"
Sub SendEventsToGoogleCalendar_001()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("予定") ' シート名を適宜変更

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' 最終行を取得

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    Dim apiUrl As String
    apiUrl = "https://script.google.com/macros/s/AKfycby-LpWIEFhpGAl5Rb-kErgwLB95s1xBBHq1rbzrwArS0cjvEqN2B1_DNkcq9avjV9KO/exec" ' GASのエンドポイントURLを入れる

    Dim json As String
    Dim eventsData As String

    Dim i As Integer
    Dim eventCount As Integer
    eventCount = 0 ' イベントの数をカウント（カンマ処理に使用）

    For i = 2 To lastRow
    On Error Resume Next ' エラー発生時のハンドリング
       
        Dim startDate As String, startTime As String, endDate As String, endTime As String
        Dim title As String, description As String, location As String
        Dim isAllDay As Boolean
        Dim eventJson As String
   
    eventsData = "{""events"":["


        ' 各列のデータ取得（Null や Empty を防ぐ）
        startDate = Trim(CStr(ws.Cells(i, 1).Value)) ' 開始日付 (A列)
        startTime = WorksheetFunction.Text(ws.Cells(i, 2).Value, "HH:mm:ss")
        endDate = Trim(CStr(ws.Cells(i, 3).Value))   ' 終了日付 (C列)
        endTime = WorksheetFunction.Text(ws.Cells(i, 4).Value, "HH:mm:ss")        
        title = Trim(CStr(ws.Cells(i, 6).Value))     ' 予定詳細 (F列)
        description = Trim(CStr(ws.Cells(i, 7).Value)) ' メモ (G列)
        description = Replace(description, vbCrLf, "\n")
        description = Replace(description, vbLf, "\n")
        description = Replace(description, vbCr, "\n")
        location = Trim(CStr(ws.Cells(i, 9).Value))  ' 施設 (I列)

        ' 日付と時刻のフォーマット修正
        If IsDate(startDate) Then startDate = Format(CDate(startDate), "yyyy-mm-dd") Else startDate = ""
        If IsDate(endDate) Then endDate = Format(CDate(endDate), "yyyy-mm-dd") Else endDate = ""
        If IsDate(startTime) Then startTime = Format(CDate(startTime), "hh:mm:ss") Else startTime = "00:00:00"
        If IsDate(endTime) Then endTime = Format(CDate(endTime), "hh:mm:ss") Else endTime = "00:00:00"

        ' 終日予定の判定
        If Trim(ws.Cells(i, 2).Value) = "" And Trim(ws.Cells(i, 4).Value) = "" Then
            isAllDay = True
        Else
            isAllDay = False
        End If

        ' JSONの構築
        eventJson = "{"
        eventJson = eventJson & """title"":""" & title & ""","
        eventJson = eventJson & """description"":""" & description & ""","
        eventJson = eventJson & """location"":""" & location & ""","

        If isAllDay Then
            eventJson = eventJson & """startTime"":""" & startDate & "T00:00:00"","
            eventJson = eventJson & """endTime"":""" & endDate & "T23:59:59"","
            eventJson = eventJson & """isAllDay"":true"
        Else
            eventJson = eventJson & """startTime"":""" & startDate & "T" & startTime & ""","
            eventJson = eventJson & """endTime"":""" & endDate & "T" & endTime & ""","
            eventJson = eventJson & """isAllDay"":false"
        End If

'        close1 = "}"
        eventJson = eventJson & "}" ' JSON の閉じ

        ' 最初のイベントではカンマなし、それ以降はカンマをつける
        eventsData = eventsData & eventJson
 
    eventsData = eventsData & "]}" ' JSONの閉じ

    ws.Cells(i, 21).Value = eventsData

' ?? イミディエイトウィンドウに `Sending JSON:` を出力
Debug.Print "Sending JSON: " & eventsData

    ' API リクエスト送信
    With http
        .Open "POST", apiUrl, False
        .SetRequestHeader "Content-Type", "application/json"
        .Send eventsData
    End With

 On Error GoTo 0
    Next i

' ?? API のレスポンスも確認
Debug.Print "API Response: " & http.responseText

End Sub


