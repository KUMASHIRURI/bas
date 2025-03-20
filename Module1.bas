Attribute VB_Name = "Module1"
Sub SendEventsToGoogleCalendar_Simple()
    On Error GoTo ErrorHandler ' �G���[�������̃n���h�����O

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("�\��") ' �V�[�g����K�X�ύX

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' �ŏI�s���擾

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    Dim apiUrl As String
    apiUrl = "https://script.google.com/macros/s/AKfycby-LpWIEFhpGAl5Rb-kErgwLB95s1xBBHq1rbzrwArS0cjvEqN2B1_DNkcq9avjV9KO/exec" ' GAS�̃G���h�|�C���gURL������

    Dim json As String
    Dim eventsData As String
    eventsData = "{""events"":["

    Dim i As Integer

    For i = 2 To lastRow
        Dim startDate As String, startTime As String, endDate As String, endTime As String
        Dim title As String, description As String, location As String
        Dim isAllDay As Boolean
        Dim eventJson As String
        Dim close1 As String
        Dim close2 As String

        ' �e��̃f�[�^�擾�iNull �� Empty ��h���j
        startDate = Trim(CStr(ws.Cells(i, 1).Value)) ' �J�n���t (A��)
        startTime = WorksheetFunction.Text(ws.Cells(i, 2).Value, "HH:mm:ss")
        endDate = Trim(CStr(ws.Cells(i, 3).Value))   ' �I�����t (C��)
        endTime = WorksheetFunction.Text(ws.Cells(i, 4).Value, "HH:mm:ss")           ' �I������ (D��)
        title = Trim(CStr(ws.Cells(i, 6).Value))     ' �\��ڍ� (F��)
        description = Trim(CStr(ws.Cells(i, 7).Value)) ' ���� (G��)
        description = Replace(description, vbCrLf, "\n")
        description = Replace(description, vbLf, "\n")
        description = Replace(description, vbCr, "\n")
        location = Trim(CStr(ws.Cells(i, 9).Value))  ' �{�� (I��)

        ' �󔒍s�̓X�L�b�v
        If title = "" Then GoTo NextRow

        ' ���t�Ǝ����̃t�H�[�}�b�g�C��
        If IsDate(startDate) Then startDate = Format(CDate(startDate), "yyyy-mm-dd") Else startDate = ""
        If IsDate(endDate) Then endDate = Format(CDate(endDate), "yyyy-mm-dd") Else endDate = ""
        If IsDate(startTime) Then startTime = Format(CDate(startTime), "hh:mm:ss") Else startTime = "00:00:00"
        If IsDate(endTime) Then endTime = Format(CDate(endTime), "hh:mm:ss") Else endTime = "00:00:00"

        ' �I���\��̔���
        If Trim(ws.Cells(i, 2).Value) = "" And Trim(ws.Cells(i, 4).Value) = "" Then
            isAllDay = True
        Else
            isAllDay = False
        End If

        ' JSON�̍\�z
        eventJson = "{"
        eventJson = eventJson & """title"":""" & title & ""","
        eventJson = eventJson & """description"":""" & description & ""","
        eventJson = eventJson & """location"":""" & location & ""","

        If isAllDay Then
            eventJson = eventJson & """startTime"":""" & startDate & "T00:00:00"","
            eventJson = eventJson & """endTime"":""" & endDate & "T23:59:59"","
            eventJson = eventJson & """isAllDay"":true}"
        Else
            eventJson = eventJson & """startTime"":""" & startDate & "T" & startTime & ""","
            eventJson = eventJson & """endTime"":""" & endDate & "T" & endTime & ""","
            eventJson = eventJson & """isAllDay"":false}"
        End If

' ?? �C�~�f�B�G�C�g�E�B���h�E�� `Sending JSON:` ���o��
Debug.Print "Sending JSON: " & eventJson

    ' API ���N�G�X�g���M
    With http
        .Open "POST", apiUrl, False
        .SetRequestHeader "Content-Type", "application/json"
        .Send eventJson
    End With

    ws.Cells(i, 12).Value = http.responseText

' ?? API �̃��X�|���X���m�F
Debug.Print "API Response: " & http.responseText

NextRow:
    Next i
 


    ' �ȈՊm�F�p�̃��b�Z�[�W
    MsgBox "Google�J�����_�[�Ƀf�[�^�𑗐M���܂����I", vbInformation
    Exit Sub ' ����I��

' �G���[�n���h�����O
ErrorHandler:
    MsgBox "�G���[����: " & Err.description, vbCritical, "�G���["
End Sub


