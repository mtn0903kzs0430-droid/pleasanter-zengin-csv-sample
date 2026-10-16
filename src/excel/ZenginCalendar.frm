VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ZenginCalendar 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "ZenginCalendar.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ZenginCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedDay As String
Private displayedMonth As Date
Private selectedDate As Date
Private handlers As Collection
Private Sub UserForm_Initialize()
    Me.Caption = "振込指定日を選択"
    Me.width = 370: Me.Height = 360
    displayedMonth = DateSerial(Year(Date), Month(Date), 1)
    selectedDate = Date
    Set handlers = New Collection
    AddButton "prev", "前月", 12, 12, 65, "prev"
    AddButton "next", "翌月", 280, 12, 65, "next"
    AddLabel "month", "", 90, 16, 180
    Dim i As Long, labels As Variant
    labels = Array("日", "月", "火", "水", "木", "金", "土")
    For i = 0 To 6
        AddLabel "week" & i, CStr(labels(i)), 12 + i * 48, 48, 44
    Next
    For i = 0 To 41
        AddButton "day" & i, "", 12 + (i Mod 7) * 48, 72 + (i \ 7) * 31, 44, "day" & i
    Next
    AddLabel "selected", "", 12, 265, 330
    AddButton "today", "本日", 12, 295, 60, "today"
    AddButton "ok", "出力へ進む", 100, 295, 120, "ok"
    AddButton "cancel", "キャンセル", 235, 295, 110, "cancel"
    Me.Controls("ok").Default = True
    Me.Controls("cancel").Cancel = True
    Draw
End Sub
Private Sub AddLabel(ByVal name As String, ByVal text As String, ByVal x As Single, ByVal y As Single, ByVal width As Single)
    Dim c As Object: Set c = Me.Controls.Add("Forms.Label.1", name)
    c.Caption = text: c.Left = x: c.Top = y: c.width = width: c.Height = 22: c.Font.Size = 12
End Sub
Private Sub AddButton(ByVal name As String, ByVal text As String, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal action As String)
    Dim h As New CalendarButton
    Set h.Button = Me.Controls.Add("Forms.CommandButton.1", name)
    h.Button.Caption = text: h.Button.Left = x: h.Button.Top = y: h.Button.width = width: h.Button.Height = 27: h.Button.Font.Size = 11
    Set h.Owner = Me: h.action = action: handlers.Add h
End Sub
Private Sub Draw()
    Dim i As Long, n As Long, days As Long, c As Object
    Me.Controls("month").Caption = Format$(displayedMonth, "yyyy年m月")
    days = day(DateSerial(Year(displayedMonth), Month(displayedMonth) + 1, 0))
    For i = 0 To 41
        n = i - Weekday(displayedMonth, vbSunday) + 2
        Set c = Me.Controls("day" & i)
        c.Visible = n >= 1 And n <= days
        c.Tag = ""
        If c.Visible Then
            c.Caption = CStr(n)
            c.Tag = Format$(DateSerial(Year(displayedMonth), Month(displayedMonth), n), "yyyy-mm-dd")
            c.BackColor = IIf(c.Tag = Format$(selectedDate, "yyyy-mm-dd"), RGB(255, 173, 51), RGB(240, 240, 240))
        End If
    Next
    SelectedDay = Format$(selectedDate, "yyyy-mm-dd")
    Me.Controls("selected").Caption = "選択日：" & Format$(selectedDate, "yyyy/mm/dd")
End Sub
Private Sub UserForm_Activate()
    Dim i As Long
    For i = 0 To 41
        If Me.Controls("day" & i).Tag = SelectedDay Then Me.Controls("day" & i).SetFocus: Exit For
    Next
End Sub
Public Sub Handle(ByVal action As String)
    Select Case action
    Case "prev": displayedMonth = DateAdd("m", -1, displayedMonth)
    Case "next": displayedMonth = DateAdd("m", 1, displayedMonth)
    Case "today": selectedDate = Date: displayedMonth = DateSerial(Year(Date), Month(Date), 1)
    Case "ok": Me.Hide: Exit Sub
    Case "cancel": SelectedDay = "": Me.Hide: Exit Sub
    Case Else
        If Left$(action, 3) = "day" Then selectedDate = DateValue(Me.Controls(action).Tag)
    End Select
    Draw
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: SelectedDay = "": Me.Hide
End Sub

