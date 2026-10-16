Attribute VB_Name = "ZenginExport"
Option Explicit
' Windows Excel. APIキーは環境変数PLEASANTER_API_KEYから読み取ります。
Private quiet As Boolean
Private lastMessage As String
Private Const FIRST_ROW As Long = 7
Private Const COLS As String = "ResultId,SiteId,UpdatedTime,Title,Status,ClassA,ClassB,ClassC,ClassD,ClassE,ClassF,NumA,DateA,DateB,DescriptionA"
Private Sub Report(ByVal text As String, Optional ByVal style As VbMsgBoxStyle = vbInformation)
    lastMessage = text
    AppendLog "通知", IIf(style = vbExclamation, "エラー", "情報"), text
    If Not quiet Then MsgBox text, style
End Sub

Private Function Ws() As Worksheet
    Set Ws = ThisWorkbook.Worksheets("支払データ")
End Function
Private Function Setting(ByVal row As Long) As String
    Setting = CStr(SettingCell(row).Value2)
End Function
Private Function GetApiKey() As String
    GetApiKey = Trim$(Environ$("PLEASANTER_API_KEY"))
    Need GetApiKey <> "" And GetApiKey <> "DUMMY_API_KEY", _
        "環境変数 PLEASANTER_API_KEY にAPIキーを設定し、Excelをすべて終了してから起動し直してください。"
End Function
Private Sub Need(ByVal ok As Boolean, ByVal msg As String)
    If Not ok Then Err.Raise vbObjectError + 800, "全銀CSV", msg
End Sub
Private Function Matches(ByVal text As String, ByVal pattern As String) As Boolean
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern: Matches = re.Test(text)
End Function
Private Function Digits(ByVal value As Variant, ByVal width As Long, ByVal label As String, Optional ByVal exact As Boolean = False) As String
    Dim s As String: s = CStr(value)
    Need Matches(s, "^[0-9]+$") And Len(s) <= width, label & ": 半角数字1～" & width & "桁で入力してください。"
    If exact Then Need Len(s) = width, label & ": " & width & "桁で入力してください。"
    Digits = String$(width - Len(s), "0") & s
End Function
Private Function Kana(ByVal value As Variant, ByVal width As Long, ByVal label As String) As String
    Dim s As String: s = CStr(value)
    Need Matches(s, "^[0-9A-Zｱ-ﾝﾞﾟ()./ -]+$") And Len(Trim$(s)) > 0 And Len(s) <= width, _
        label & ": 使用可能な半角文字で1～" & width & "文字にしてください。長音は半角ハイフンです。"
    Kana = s
End Function
Private Function IsoDay(ByVal value As Variant) As String
    Dim s As String, d As Date
    If VarType(value) = vbDate Or IsNumeric(value) Then
        d = CDate(value): s = Format$(d, "yyyy-mm-dd")
    Else
        s = Left$(CStr(value), 10)
        Need Matches(s, "^[0-9]{4}-[0-9]{2}-[0-9]{2}$"), "振込指定日が不正です。"
        d = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Right$(s, 2)))
        Need Format$(d, "yyyy-mm-dd") = s, "存在しない振込指定日です。"
    End If
    Need d >= DateSerial(2000, 1, 1), "振込指定日が未入力です。"
    IsoDay = s
End Function
Private Function ViewJson() As String
    ViewJson = """ApiDataType"":""Default"",""Incomplete"":false,""ColumnSorterHash"":{""ResultId"":""asc""}"
End Function
Private Function Normalize(ByVal source As Object) As Object
    Dim r As Object, group As Variant, key As Variant, c As Variant
    Set r = CreateObject("Scripting.Dictionary")
    For Each c In Split(COLS, ",")
        r(CStr(c)) = JVal(source, CStr(c))
    Next
    For Each group In Array("ClassHash", "NumHash", "DateHash", "DescriptionHash")
        If source.Exists(CStr(group)) Then
            For Each key In source(CStr(group)).Keys
                r(CStr(key)) = JVal(source(CStr(group)), CStr(key))
            Next
        End If
    Next
    If source.Exists("ItemTitle") Then r("Title") = JVal(source, "ItemTitle")
    Set Normalize = r
End Function
Private Sub ConnectSettings()
    InitializeProtection
    Need ThisWorkbook.FileFormat = 52, "このブックを.xlsm形式で保存してください。"
    Need Not ThisWorkbook.ReadOnly, "読み取り専用のブックでは実行できません。"
    Need Matches(Setting(5), "^https?://"), "設定B5にプリザンターのURLを入力してください。"
    Need Matches(Setting(6), "^[0-9]+$"), "設定B6のサイトIDを確認してください。"
    Dim apiKey As String: apiKey = GetApiKey()
    ' 保護したまま、マクロによる取得状態の記録を許可します。
    ' Protection initialized for all sheets before processing.
End Sub
Private Function Api(ByVal id As String, ByVal action As String, ByVal body As String) As Object
    Dim http As Object, result As Object, url As String, decoder As Object, responseText As String
    url = Setting(5)
    Do While Right$(url, 1) = "/": url = Left$(url, Len(url) - 1): Loop
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, 30000
    http.Open "POST", url & "/api/items/" & id & "/" & action, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.Send "{""ApiVersion"":1.1,""ApiKey"":" & JsonQuote(GetApiKey()) & IIf(body = "", "", "," & body) & "}"
    Need http.Status <> 401, "ID " & id & ": HTTP 401。環境変数 PLEASANTER_API_KEY のAPIキーを確認し、変更後はExcelを起動し直してください。"
    Need http.Status = 200, "ID " & id & ": HTTP " & http.Status
    ' JSONはUTF-8。サーバーがcharsetを省略しても文字化けさせない。
    Set decoder = CreateObject("ADODB.Stream")
    decoder.Type = 1: decoder.Open: decoder.Write http.ResponseBody
    decoder.Position = 0: decoder.Type = 2: decoder.Charset = "utf-8"
    responseText = decoder.ReadText: decoder.Close
    Set result = ParseJson(responseText)
    Need CStr(JVal(result, "StatusCode")) = "200", "ID " & id & ": APIエラー " & CStr(JVal(result, "StatusCode"))
    Set Api = result
End Function
Private Function GetOne(ByVal id As String) As Object
    Dim result As Object, rows As Collection
    Set result = Api(id, "get", """View"":{" & ViewJson() & "}")
    Set rows = result("Response")("Data")
    Need rows.count = 1, "ID " & id & "を取得できません。"
    Set GetOne = Normalize(rows(1))
    Need CStr(JVal(GetOne, "SiteId")) = Setting(6) And CStr(JVal(GetOne, "ResultId")) = id, "サイトまたはIDが一致しません。"
End Function
Private Function Fingerprint(ByVal r As Object) As String
    Dim c As Variant, s As String
    For Each c In Split(COLS, ",")
        If CStr(c) <> "Title" Then s = s & JsonQuote(CStr(JVal(r, CStr(c)))) & "|"
    Next
    Fingerprint = s
End Function
Private Function LastRow() As Long
    LastRow = Ws.Cells(Ws.rows.count, 2).End(xlUp).row
End Function
Private Sub NoPending()
    Dim i As Long
    For i = FIRST_ROW To LastRow()
        Need CStr(Ws.Cells(i, 15).Value2) <> "更新待ち" And Left$(CStr(Ws.Cells(i, 15).Value2), 5) <> "更新失敗:", _
            "未完了の更新があります。先にRetryUpdatesを実行してください。"
    Next
End Sub

Public Sub ExportByDate()
    On Error GoTo Failed
ConnectSettings:     NoPending
    Dim inputDay As String, chosen As Collection, total As Double, text As String
    AppendLog "出力開始", "開始"
    inputDay = CalendarSelectedDay()
    If inputDay = "" Then AppendLog "日付選択", "キャンセル": Exit Sub
    Need Len(inputDay) = 10, "yyyy-mm-ddで入力してください。"
    SettingCell(7).value = DateValue(IsoDay(inputDay))
    quiet = True: FetchPayments: quiet = False
    Need Setting(14) = "API取得済", "取得できませんでした: " & lastMessage
    text = BuildCsv(chosen, total)
    If MsgBox("振込指定日 " & inputDay & vbCrLf & chosen.count & "件 / " & Format$(total, "#,##0") & "円" & vbCrLf & "CSVを保存しますか？", vbOKCancel + vbQuestion, "全銀CSV出力") <> vbOK Then AppendLog "内容確認", "キャンセル": Exit Sub
    ExportCsv
    Exit Sub
Failed:
    quiet = False: Report Err.Description, vbExclamation
End Sub

Private Sub RecheckTargets(ByVal chosen As Collection)
    Dim day As String, body As String, offset As Long, result As Object, data As Collection
    Dim seen As Object, rec As Variant, r As Object, row As Variant, id As String, totalCount As Long
    Set seen = CreateObject("Scripting.Dictionary")
    day = IsoDay(Ws.Cells(CLng(chosen(1)), 10).value)
    body = """View"":{" & ViewJson() & ",""ColumnFilterHash"":{""Status"":""[100]"",""DateA"":" & _
        JsonQuote("[""" & day & " 00:00:00," & day & " 23:59:59""]") & "}}"
    For Each row In chosen: seen.Add CStr(Ws.Cells(CLng(row), 2).Value2), CLng(row): Next
    Do
        Set result = Api(Setting(6), "get", body & ",""Offset"":" & offset)
        Set data = result("Response")("Data")
        Need CLng(result("Response")("TotalCount")) = chosen.count, "対象件数が変わりました。再取得してください。"
        For Each rec In data
            Set r = Normalize(rec): id = CStr(r("ResultId"))
            Need seen.Exists(id), "出力対象が変わりました。再取得してください。"
            CheckSource CLng(seen(id)), r: seen.Remove id
        Next
        offset = offset + data.count
        If offset >= chosen.count Then Exit Do
        Need data.count > 0, "取得が途中で停止しました。"
    Loop
    Need seen.count = 0, "取得件数が一致しません。"
End Sub

Public Sub FetchPayments()
    On Error GoTo Failed
ConnectSettings:     NoPending
    SettingCell(14).Value2 = "取得中"
    Dim day As String, body As String, offset As Long, result As Object, data As Collection
    Dim all As New Collection, rec As Variant, r As Object, i As Long, seen As Object
    day = IsoDay(SettingCell(7).value)
    body = """View"":{" & ViewJson() & ",""ColumnFilterHash"":{""Status"":""[100]"",""DateA"":" & _
        JsonQuote("[""" & day & " 00:00:00," & day & " 23:59:59""]") & "}}"
    Set seen = CreateObject("Scripting.Dictionary")
    Do
        Set result = Api(Setting(6), "get", body & ",""Offset"":" & offset)
        Set data = result("Response")("Data")
        For Each rec In data
            Set r = Normalize(rec)
            Need Not seen.Exists(CStr(r("ResultId"))), "取得中に重複IDが発生しました。取得し直してください。"
            seen.Add CStr(r("ResultId")), True
            Need CStr(r("SiteId")) = Setting(6), "別サイトのデータが含まれています。"
            Need CStr(r("Status")) = "100" And IsoDay(r("DateA")) = day, "取得条件が一致しません。"
            all.Add r
        Next
        offset = offset + data.count
        If offset >= CLng(result("Response")("TotalCount")) Then Exit Do
        Need data.count > 0, "ページングが途中で停止しました。"
    Loop
    If LastRow() >= FIRST_ROW Then Ws.Range("A7:Q" & LastRow()).ClearContents
    Need all.count <= 999999, "取得件数が999999件を超えています。"
    Ws.Range("B7:H" & Application.Max(7, 6 + all.count)).NumberFormat = "@"
    Ws.Range("L7:Q" & Application.Max(7, 6 + all.count)).NumberFormat = "@"
    i = FIRST_ROW
    For Each rec In all
        Set r = rec
        Ws.Cells(i, 1).Value2 = 1
        Ws.Cells(i, 2).NumberFormat = "@": Ws.Cells(i, 2).Value2 = CStr(r("ResultId"))
        Ws.Cells(i, 3).Value2 = CStr(JVal(r, "Title"))
        Ws.Range("D" & i & ":H" & i).NumberFormat = "@"
        Ws.Cells(i, 4).Value2 = CStr(JVal(r, "ClassB"))
        Ws.Cells(i, 5).Value2 = CStr(JVal(r, "ClassC"))
        Ws.Cells(i, 6).Value2 = CStr(JVal(r, "ClassD"))
        Ws.Cells(i, 7).Value2 = CStr(JVal(r, "ClassE"))
        Ws.Cells(i, 8).Value2 = CStr(JVal(r, "ClassF"))
        Ws.Cells(i, 9).Value2 = JVal(r, "NumA")
        Ws.Cells(i, 10).value = DateValue(IsoDay(JVal(r, "DateA")))
        Ws.Cells(i, 11).Value2 = JVal(r, "Status")
        Ws.Cells(i, 12).NumberFormat = "@": Ws.Cells(i, 12).Value2 = CStr(JVal(r, "UpdatedTime"))
        Ws.Cells(i, 13).NumberFormat = "@": Ws.Cells(i, 13).Value2 = Fingerprint(r)
        i = i + 1
    Next
    Ws.Range("I7:I" & Application.Max(7, i)).NumberFormat = "#,##0"
    Ws.Range("J7:J" & Application.Max(7, i)).NumberFormat = "yyyy-mm-dd"
    Ws.Range("B3").Formula = "=COUNTIFS(A7:A" & Application.Max(7, i - 1) & ",1)"
    Ws.Range("D3").Formula = "=SUMIFS(I7:I" & Application.Max(7, i - 1) & ",A7:A" & Application.Max(7, i - 1) & ",1)"
    SettingCell(14).Value2 = "API取得済"
    SettingCell(15).Value2 = Setting(5) & "|" & Setting(6)
    ThisWorkbook.Save
    AppendLog "データ取得", "成功", "", all.count, Ws.Range("D3").Value2
    Report all.count & "件取得しました。すべて出力対象です。", vbInformation
    Exit Sub
Failed:
    Report Err.Description, vbExclamation
End Sub

Private Function BuildCsv(ByRef chosen As Collection, ByRef total As Double, Optional ByVal preview As Boolean = True) As String
    Dim i As Long, r As Long, n As Long, id As String, day As String, d As String, amount As Double
    Dim fields As Variant, lines As New Collection, ids As Object, sh As Worksheet, item As Variant, text As String
    Set chosen = New Collection: Set ids = CreateObject("Scripting.Dictionary")
    total = 0
    For i = FIRST_ROW To LastRow()
        If CStr(Ws.Cells(i, 2).Value2) <> "" Then
            id = CStr(Ws.Cells(i, 2).Value2)
            Need Matches(id, "^[1-9][0-9]{0,19}$"), "行" & i & ": レコードIDが不正です。"
            Need Not ids.Exists(id), "IDが重複しています: " & id
            ids.Add id, True
            Need (Ws.Cells(i, 11).Value2 = 100), "ID " & id & ": 未出力以外は出力できません。"
            d = IsoDay(Ws.Cells(i, 10).value)
            If day = "" Then day = d
            Need day = d, "異なる振込指定日が混在しています。"
            Need IsNumeric(Ws.Cells(i, 9).Value2), "金額を数値で入力してください。"
            amount = CDbl(Ws.Cells(i, 9).Value2)
            Need amount > 0 And amount <= 9999999999# And amount = Fix(amount), "ID " & id & ": 金額は1～9999999999の整数です。"
            Need Matches(CStr(Ws.Cells(i, 6).Value2), "^[124]$"), "預金種目は1・2・4です。"
            total = total + amount: Need total <= 999999999999#, "合計金額が12桁を超えています。"
            chosen.Add i
            fields = Array("2", Digits(Ws.Cells(i, 4).Value2, 4, "銀行番号", True), "", _
                Digits(Ws.Cells(i, 5).Value2, 3, "支店番号", True), "", "", CStr(Ws.Cells(i, 6).Value2), _
                Digits(Ws.Cells(i, 7).Value2, 7, "口座番号"), Kana(Ws.Cells(i, 8).Value2, 30, "受取人名"), _
                Digits(Format$(amount, "0"), 10, "金額"), "1", id, "", "", "")
            lines.Add Join(fields, ",")
        End If
    Next
    Need chosen.count > 0 And chosen.count <= 999999, "指定日に未出力のレコードはありません。"
    fields = Array("1", "21", "0", Digits(Setting(8), 10, "委託者コード", True), Kana(Setting(9), 40, "委託者名"), _
        Mid$(day, 6, 2) & Right$(day, 2), "0036", "", Digits(Setting(10), 3, "依頼人支店番号", True), "", "1", Digits(Setting(11), 7, "依頼人口座番号"), "")
    text = Join(fields, ",") & vbCrLf
    For Each item In lines: text = text & CStr(item) & vbCrLf: Next
    text = text & "8," & Digits(chosen.count, 6, "件数") & "," & Digits(Format$(total, "0"), 12, "合計金額") & "," & vbCrLf & "9," & vbCrLf
    If preview Then
        Set sh = ThisWorkbook.Worksheets("CSV確認")
        sh.Range("A6:O" & Application.Max(6, sh.UsedRange.rows.count)).ClearContents
        sh.Range("A6:O" & chosen.count + 8).NumberFormat = "@"
        r = 6
        For Each item In Split(Left$(text, Len(text) - 2), vbCrLf)
            fields = Split(CStr(item), ",")
            For n = 0 To UBound(fields): sh.Cells(r, n + 1).Value2 = fields(n): Next
            r = r + 1
        Next
        sh.Range("A3").Value2 = "振込指定日 " & day & " / " & chosen.count & "件 / " & Format$(total, "#,##0") & "円"
    End If
    BuildCsv = text
End Function

Public Sub PrepareCsv()
    On Error GoTo Failed
    Dim chosen As Collection, total As Double, text As String
    text = BuildCsv(chosen, total)
    ThisWorkbook.Worksheets("CSV確認").Activate
    Report "整形しました。CSV確認は確認用です。修正はプリザンターで行い、再取得してください。", vbInformation
    Exit Sub
Failed:
    Report Err.Description, vbExclamation
End Sub

Private Sub CheckSource(ByVal row As Long, ByVal r As Object)
    Need CStr(Ws.Cells(row, 13).Value2) = Fingerprint(r), "ID " & Ws.Cells(row, 2).Value2 & ": 取得後にサーバーの値が変わりました。再取得してください。"
    Dim c As Long, names As Variant
    names = Array("ClassB", "ClassC", "ClassD", "ClassE", "ClassF")
    For c = 0 To 4
        Need CStr(Ws.Cells(row, c + 4).Value2) = CStr(JVal(r, CStr(names(c)))), "口座情報がExcel上で変更されています。プリザンターを修正して再取得してください。"
    Next
    Need CDbl(Ws.Cells(row, 9).Value2) = CDbl(r("NumA")), "金額がExcel上で変更されています。"
    Need IsoDay(Ws.Cells(row, 10).value) = IsoDay(r("DateA")), "振込指定日がExcel上で変更されています。"
End Sub

Private Function CsvBytes(ByVal text As String) As Byte()
    Dim b() As Byte, i As Long, n As Long
    ReDim b(0 To Len(text) - 1)
    For i = 1 To Len(text)
        n = AscW(Mid$(text, i, 1)): If n < 0 Then n = n + 65536
        If n <= 127 Then
            b(i - 1) = n
        ElseIf n >= &HFF61& And n <= &HFF9F& Then
            b(i - 1) = n - &HFF61& + &HA1&
        Else
            Err.Raise vbObjectError + 801, , "CSVに半角以外の文字があります。"
        End If
    Next
    CsvBytes = b
End Function
Private Sub SaveBytes(ByVal path As String, ByVal text As String)
    Dim stream As Object, b() As Byte
    b = CsvBytes(text)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open: stream.Write b
    stream.SaveToFile path, 1 ' adSaveCreateNotExist: do not overwrite an existing export.
    stream.Close
End Sub

Private Sub ShowPaymentDataHome()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("支払データ")
    sh.Activate
    Application.GoTo sh.Range("A1"), True
End Sub

Public Sub ExportCsv()
    On Error GoTo Failed
ConnectSettings:     NoPending
    Need Setting(14) = "API取得済", "サンプルのまま出力せず、FetchPaymentsでAPI取得してください。"
    Need Setting(15) = Setting(5) & "|" & Setting(6), "取得後に接続先が変わっています。再取得してください。"
    Dim chosen As Collection, total As Double, text As String, row As Variant, current As Object
    Dim stamp As String, name As String, folder As String, path As String, savedFile As Boolean, exportTime As Date
    text = BuildCsv(chosen, total)
    RecheckTargets chosen
    For Each row In chosen
        Set current = GetOne(CStr(Ws.Cells(CLng(row), 2).Value2))
        CheckSource CLng(row), current
    Next
    folder = Setting(12): If folder = "" Then folder = ThisWorkbook.path
    Need CreateObject("Scripting.FileSystemObject").FolderExists(folder), "出力先フォルダが存在しません。"
    exportTime = Now
    stamp = Format$(exportTime, "yyyy-mm-dd\Thh:nn:ss")
    name = "楽天銀行_総合振込_振込日" & Replace(IsoDay(Ws.Cells(CLng(chosen(1)), 10).value), "-", "") & "_作成" & Format$(exportTime, "yyyymmdd_hhnnss") & ".csv"
    path = folder & "\" & name
    ThisWorkbook.Save
    SaveBytes path, text: savedFile = True
    AppendLog "CSV保存", "成功", path, chosen.count, total, name
    For Each row In chosen
        Ws.Cells(CLng(row), 14).NumberFormat = "@": Ws.Cells(CLng(row), 14).Value2 = name
        Ws.Cells(CLng(row), 15).Value2 = "更新待ち"
        Ws.Cells(CLng(row), 16).NumberFormat = "@": Ws.Cells(CLng(row), 16).Value2 = stamp
        Ws.Cells(CLng(row), 17).NumberFormat = "@": Ws.Cells(CLng(row), 17).Value2 = path
    Next
    ThisWorkbook.Save
    RetryUpdates True, folder, name
    Exit Sub
Failed:
    If savedFile Then
        Report "CSVは保存されています: " & path & vbCrLf & "再出力せず、保存済みの各行の更新状態を確認してください。" & vbCrLf & Err.Description, vbExclamation
    Else
        Report Err.Description, vbExclamation
    End If
End Sub

Private Function UpdateOne(ByVal row As Long) As String
    On Error GoTo Failed
    Dim r As Object, response As Object, id As String, name As String, stamp As String
    id = CStr(Ws.Cells(row, 2).Value2): name = CStr(Ws.Cells(row, 14).Value2): stamp = CStr(Ws.Cells(row, 16).Value2)
    Need CreateObject("Scripting.FileSystemObject").FileExists(CStr(Ws.Cells(row, 17).Value2)), "保存済みCSVが見つかりません。"
    Set r = GetOne(id)
    If Not (CStr(JVal(r, "Status")) = "900" And CStr(JVal(r, "DescriptionA")) = name And Left$(CStr(JVal(r, "DateB")), 19) = stamp) Then
        CheckSource row, r
        Set response = Api(id, "update", """Status"":900," & _
            """DateHash"":{""DateB"":" & JsonQuote(stamp) & "},""DescriptionHash"":{""DescriptionA"":" & JsonQuote(name) & "}")
        Set r = GetOne(id)
        Need CStr(JVal(r, "Status")) = "900" And CStr(JVal(r, "DescriptionA")) = name And Left$(CStr(JVal(r, "DateB")), 19) = stamp, "更新後の値を確認できません。"
    End If
    Ws.Cells(row, 11).Value2 = 900: Ws.Cells(row, 1).Value2 = 0
    UpdateOne = "更新済": Exit Function
Failed:
    UpdateOne = "更新失敗:" & Err.Description
End Function

Public Sub RetryUpdates(Optional ByVal fromCsvExport As Boolean = False, Optional ByVal savedFolder As String = "", Optional ByVal savedName As String = "")
    On Error GoTo Failed
    ConnectSettings
    Need Setting(15) = Setting(5) & "|" & Setting(6), "取得時と接続先が異なります。"
    Dim i As Long, failedCount As Long, okCount As Long, result As String
    For i = FIRST_ROW To LastRow()
        result = CStr(Ws.Cells(i, 15).Value2)
        If result = "更新待ち" Or Left$(result, 5) = "更新失敗:" Then
            result = UpdateOne(i): Ws.Cells(i, 15).Value2 = result
            If result = "更新済" Then okCount = okCount + 1 Else failedCount = failedCount + 1
            ThisWorkbook.Save
        End If
    Next
    AppendLog "状態更新・再試行", IIf(failedCount = 0, "成功", "一部失敗"), "更新成功 " & okCount & "件 / 更新失敗 " & failedCount & "件", okCount + failedCount
    If failedCount = 0 Then
        If fromCsvExport Then
            Report "CSV出力と支払管理の更新が完了しました。" & vbCrLf & _
                "保存フォルダ：" & savedFolder & vbCrLf & _
                "ファイル名：" & savedName, vbInformation
        Else
            Report "支払管理の更新が完了しました。" & vbCrLf & _
                "今回の更新済み：" & okCount & "件", vbInformation
        End If
        ShowPaymentDataHome
    Else
        Report "支払管理の更新が一部失敗しました。" & vbCrLf & _
            "更新済み " & okCount & "件 / 更新失敗 " & failedCount & "件" & vbCrLf & _
            "「支払データ」シートのO列（更新結果）を確認してください。" & vbCrLf & _
            "CSVは保存済みのため、再出力せず「更新だけ再試行」を実行してください。", vbExclamation
    End If
    Exit Sub
Failed:
    Report Err.Description, vbExclamation
End Sub

Public Function SelfTest() As String
    ' Offline only: no API calls, no message box, no bank upload.
    On Error GoTo Failed
    Dim obj As Object, chosen As Collection, total As Double, text As String, b() As Byte
    Set obj = ParseJson("{""a"":true,""b"":[1,2],""s"":""\uFF76\uFF9E"",""n"":null}")
    Need obj("a") And obj("b").count = 2 And obj("s") = "ｶﾞ", "JSONテスト失敗"
    b = CsvBytes("ｶﾞ" & vbCrLf)
    Need b(0) = &HB6 And b(1) = &HDE And b(2) = 13 And b(3) = 10, "文字コードテスト失敗"
    text = BuildCsv(chosen, total)
    Need chosen.count = 3 And total = 56245, "サンプル集計が一致しません。"
    Need InStr(text, "8,000003,000000056245," & vbCrLf) > 0, "トレーラが一致しません。"
    SelfTest = "PASS: JSON / encoding / 3 records / 56245 yen": Exit Function
Failed:
    SelfTest = "FAIL: " & Err.Description
End Function

Public Function LocalMockWorkflowTest(ByVal baseUrl As String, ByVal folder As String) As String
    ' Test-only entry point. Use a disposable copy and the bundled local mock server.
    On Error GoTo Failed
    Need Left$(baseUrl, 17) = "http://127.0.0.1:", "ローカルモック以外では実行できません。"
    InitializeProtection
    quiet = True
    Dim sh As Worksheet: Set sh = ThisWorkbook.Worksheets("設定")
    sh.Unprotect
    Need GetApiKey() = "LOCAL_TEST_KEY", "模擬テスト用の環境変数が設定されていません。"
    SettingCell(13).Value2 = "PLEASANTER_API_KEY"
    sh.Protect UserInterfaceOnly:=True
    SettingCell(5).Value2 = baseUrl
    SettingCell(6).Value2 = "9825"
    SettingCell(7).value = DateSerial(2026, 9, 30)
    SettingCell(8).Value2 = "0000000000": SettingCell(9).Value2 = "ﾃｽﾄ(ｶ"
    SettingCell(10).Value2 = "000": SettingCell(11).Value2 = "0000000"
    SettingCell(12).Value2 = folder
    FetchPayments
    Need LastRow() = 9 And Setting(14) = "API取得済", "取得テスト失敗: " & lastMessage
    Need CStr(Ws.Range("D7").Value2) = "0010", "先頭ゼロの取得テスト失敗"
    ExportCsv
    Need CStr(Ws.Range("O7").Value2) = "更新済", "更新テスト失敗: " & lastMessage
    Need Left$(CStr(Ws.Range("O8").Value2), 5) = "更新失敗:", "一部失敗の記録テスト失敗"
    Need CStr(Ws.Range("O9").Value2) = "更新済", "後続行の更新テスト失敗"
    Need InStr(lastMessage, "「支払データ」シートのO列（更新結果）") > 0 And InStr(lastMessage, "更新だけ再試行") > 0, "更新失敗メッセージのテスト失敗: " & lastMessage
    RetryUpdates
    Need CStr(Ws.Range("O8").Value2) = "更新済", "更新再試行テスト失敗: " & lastMessage
    Need InStr(lastMessage, "支払管理の更新が完了しました。") > 0 And InStr(lastMessage, "今回の更新済み：1件") > 0, "更新再試行メッセージのテスト失敗: " & lastMessage
    LocalMockWorkflowTest = "PASS: paging / save CSV / partial failure / retry / update verification"
    quiet = False: Exit Function
Failed:
    quiet = False
    LocalMockWorkflowTest = "FAIL: " & Err.Description
End Function

Public Function EnvironmentKeyMissingTest() As String
    On Error GoTo Expected
    Dim key As String: key = GetApiKey()
    EnvironmentKeyMissingTest = "FAIL: 未設定エラーが発生しません。": Exit Function
Expected:
    If InStr(Err.Description, "PLEASANTER_API_KEY") > 0 And InStr(Err.Description, "起動し直") > 0 Then
        EnvironmentKeyMissingTest = "PASS: missing environment key / restart guidance"
    Else
        EnvironmentKeyMissingTest = "FAIL: " & Err.Description
    End If
End Function
