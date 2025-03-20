Attribute VB_Name = "これを大事に"
Sub SendEventsToGoogleCalendar_Simple()
    On Error GoTo ErrorHandler ' エラー発生時のハンドリング

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
    eventsData = "{""events"":["

    Dim i As Integer
    Dim eventCount As Integer
    eventCount = 0 ' イベントの数をカウント（カンマ処理に使用）

    For i = 2 To lastRow
        Dim startDate As String, startTime As String, endDate As String, endTime As String
        Dim title As String, description As String, location As String
        Dim isAllDay As Boolean
        Dim eventJson As String
        Dim close1 As String
        Dim close2 As String

        ' 各列のデータ取得（Null や Empty を防ぐ）
        startDate = Trim(CStr(ws.Cells(i, 1).Value)) ' 開始日付 (A列)
        'startTime = Trim(CStr(ws.Cells(i, 2).Value)) ' 開始時刻 (B列)
        startTime = WorksheetFunction.Text(ws.Cells(i, 2).Value, "HH:mm:ss")
         ' 開始時刻 (B列)
         ' 開始時刻 (B列)
        
        endDate = Trim(CStr(ws.Cells(i, 3).Value))   ' 終了日付 (C列)
        'endTime = Trim(CStr(ws.Cells(i, 4).Value))   ' 終了時刻 (D列)
        endTime = WorksheetFunction.Text(ws.Cells(i, 4).Value, "HH:mm:ss")           ' 終了時刻 (D列)
   ' 終了時刻 (D列)
        title = Trim(CStr(ws.Cells(i, 6).Value))     ' 予定詳細 (F列)
        description = Trim(CStr(ws.Cells(i, 7).Value)) ' メモ (G列)
        description = Replace(description, vbCrLf, "\n")
        description = Replace(description, vbLf, "\n")
        description = Replace(description, vbCr, "\n")
        location = Trim(CStr(ws.Cells(i, 9).Value))  ' 施設 (I列)

        ' 空白行はスキップ
        If title = "" Then GoTo NextRow

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

        close1 = "}"
        eventJson = eventJson & close1 ' JSON の閉じ

        ' 最初のイベントではカンマなし、それ以降はカンマをつける
        If eventCount > 0 Then
            eventsData = eventsData & "," & eventJson
        Else
            eventsData = eventsData & eventJson
        End If

        eventCount = eventCount + 1
        

 
    ws.Cells(1, 20).Value = eventJson
  
    eventsData = eventsData & "]}" ' JSONの閉じ

    ws.Cells(2, 20).Value = eventsData


' ?? イミディエイトウィンドウに `Sending JSON:` を出力
Debug.Print "Sending JSON: " & eventsData

    ' API リクエスト送信
    With http
        .Open "POST", apiUrl, False
        .SetRequestHeader "Content-Type", "application/json"
        .Send eventsData
    End With

' ?? API のレスポンスも確認
Debug.Print "API Response: " & http.responseText

    ' 簡易確認用のメッセージ
    MsgBox "Googleカレンダーにデータを送信しました！", vbInformation
    Exit Sub ' 正常終了
    Next i

' エラーハンドリング
ErrorHandler:
    MsgBox "エラー発生: " & Err.description, vbCritical, "エラー"
End Sub

