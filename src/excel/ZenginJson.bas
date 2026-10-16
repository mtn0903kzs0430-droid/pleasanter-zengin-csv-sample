Attribute VB_Name = "ZenginJson"
Option Explicit
' JSON reader for API responses. No reference libraries required.
Private src As String, pos As Long

Public Function ParseJson(ByVal text As String) As Object
    src = text: pos = 1
    Set ParseJson = ReadValue()
    SkipSpace
    If pos <= Len(src) Then BadJson
End Function

Private Sub BadJson()
    Err.Raise vbObjectError + 901, "JSON", "Invalid JSON at position " & pos
End Sub

Private Sub SkipSpace()
    Do While pos <= Len(src)
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid$(src, pos, 1), vbBinaryCompare) = 0 Then Exit Do
        pos = pos + 1
    Loop
End Sub

Private Sub Expect(ByVal ch As String)
    SkipSpace
    If Mid$(src, pos, 1) <> ch Then BadJson
    pos = pos + 1
End Sub

Private Function ReadValue() As Variant
    Dim obj As Object, arr As Collection, key As String, start As Long, token As String
    SkipSpace
    Select Case Mid$(src, pos, 1)
        Case "{"
            Set obj = CreateObject("Scripting.Dictionary")
            pos = pos + 1: SkipSpace
            If Mid$(src, pos, 1) <> "}" Then
                Do
                    key = ReadString(): Expect ":"
                    If obj.Exists(key) Then BadJson
                    obj.Add key, ReadValue()
                    SkipSpace
                    If Mid$(src, pos, 1) = "}" Then Exit Do
                    Expect ","
                Loop
            End If
            Expect "}": Set ReadValue = obj
        Case "["
            Set arr = New Collection
            pos = pos + 1: SkipSpace
            If Mid$(src, pos, 1) <> "]" Then
                Do
                    arr.Add ReadValue()
                    SkipSpace
                    If Mid$(src, pos, 1) = "]" Then Exit Do
                    Expect ","
                Loop
            End If
            Expect "]": Set ReadValue = arr
        Case """": ReadValue = ReadString()
        Case "t"
            If Mid$(src, pos, 4) <> "true" Then BadJson
            pos = pos + 4: ReadValue = True
        Case "f"
            If Mid$(src, pos, 5) <> "false" Then BadJson
            pos = pos + 5: ReadValue = False
        Case "n"
            If Mid$(src, pos, 4) <> "null" Then BadJson
            pos = pos + 4: ReadValue = Null
        Case Else
            start = pos
            Do While pos <= Len(src)
                If InStr(1, "-+0123456789.eE", Mid$(src, pos, 1), vbBinaryCompare) = 0 Then Exit Do
                pos = pos + 1
            Loop
            token = Mid$(src, start, pos - start)
            If token = "" Then BadJson
            Dim re As Object
            Set re = CreateObject("VBScript.RegExp")
            re.pattern = "^-?(0|[1-9][0-9]*)(\.[0-9]+)?([eE][+-]?[0-9]+)?$"
            If Not re.Test(token) Then BadJson
            ' IDs larger than 15 digits remain text (Excel cannot store those precisely).
            If Len(Replace(token, "-", "")) > 15 And InStr(token, ".") = 0 And InStr(LCase$(token), "e") = 0 Then
                ReadValue = token
            Else
                ReadValue = Val(token)
            End If
    End Select
End Function

Private Function ReadString() As String
    Dim ch As String, result As String, hex As String, code As Long
    Expect """"
    Do While pos <= Len(src)
        ch = Mid$(src, pos, 1): pos = pos + 1
        If ch = """" Then ReadString = result: Exit Function
        If ch = "\" Then
            ch = Mid$(src, pos, 1): pos = pos + 1
            Select Case ch
                Case """", "\", "/": result = result & ch
                Case "b": result = result & Chr$(8)
                Case "f": result = result & Chr$(12)
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "u"
                    hex = Mid$(src, pos, 4)
                    If Len(hex) <> 4 Or hex Like "*[!0-9A-Fa-f]*" Then BadJson
                    code = CLng("&H" & hex): pos = pos + 4
                    If code > 32767 Then code = code - 65536
                    result = result & ChrW$(code)
                Case Else: BadJson
            End Select
        Else
            If AscW(ch) >= 0 And AscW(ch) < 32 Then BadJson
            result = result & ch
        End If
    Loop
    BadJson
End Function

Public Function JsonQuote(ByVal text As String) As String
    Dim i As Long, ch As String, n As Long, result As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1): n = AscW(ch)
        Select Case ch
            Case """": result = result & "\"""
            Case "\": result = result & "\\"
            Case vbCr: result = result & "\r"
            Case vbLf: result = result & "\n"
            Case vbTab: result = result & "\t"
            Case Else
                If n >= 0 And n < 32 Then
                    result = result & "\u" & Right$("0000" & hex$(n), 4)
                Else
                    result = result & ch
                End If
        End Select
    Next
    JsonQuote = """" & result & """"
End Function

Public Function JVal(ByVal obj As Object, ByVal key As String) As Variant
    If Not obj.Exists(key) Then JVal = "": Exit Function
    If IsNull(obj(key)) Then JVal = "": Exit Function
    JVal = obj(key)
End Function
