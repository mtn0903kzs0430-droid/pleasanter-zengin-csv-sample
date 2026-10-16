Attribute VB_Name = "ZenginWorkbook"
Option Explicit
Public Function SettingCell(ByVal logicalRow As Long) As Range
    Dim addr As String
    Select Case logicalRow
    Case 5: addr = "B5"
    Case 6: addr = "B6"
    Case 7: addr = "J2"
    Case 8: addr = "B12"
    Case 9: addr = "B13"
    Case 10: addr = "B14"
    Case 11: addr = "B15"
    Case 12: addr = "B8"
    Case 13: addr = "B7"
    Case 14: addr = "J3"
    Case 15: addr = "J4"
    End Select
    Set SettingCell = ThisWorkbook.Worksheets("ê›íË").Range(addr)
End Function
Public Sub InitializeProtection()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        sh.Unprotect
        sh.Cells.Locked = True
        If sh.name = "ê›íË" Then
            sh.Range("B5:B6,B8,B12:B15").Locked = False
        End If
        sh.Protect UserInterfaceOnly:=True, AllowFiltering:=True
        sh.EnableSelection = xlNoRestrictions
    Next
End Sub
Public Sub AppendLog(ByVal action As String, ByVal result As String, Optional ByVal detail As String = "", Optional ByVal count As Variant, Optional ByVal total As Variant, Optional ByVal fileName As String = "")
    Dim sh As Worksheet, r As Long
    Set sh = ThisWorkbook.Worksheets("é¿çsÉçÉO")
    r = sh.Cells(sh.rows.count, 1).End(xlUp).row + 1
    sh.Cells(r, 1).value = Now
    sh.Cells(r, 2).Value2 = action
    If IsDate(SettingCell(7).value) Then sh.Cells(r, 3).value = SettingCell(7).value
    If Not IsMissing(count) Then sh.Cells(r, 4).Value2 = count
    If Not IsMissing(total) Then sh.Cells(r, 5).Value2 = total
    sh.Cells(r, 6).Value2 = fileName
    sh.Cells(r, 7).Value2 = result
    sh.Cells(r, 8).Value2 = detail
    sh.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    sh.Cells(r, 3).NumberFormat = "yyyy/mm/dd"
    sh.Cells(r, 5).NumberFormat = "#,##0"
    ThisWorkbook.Save
End Sub
Public Function CalendarSelectedDay() As String
    Dim f As New ZenginCalendar
    f.Show vbModal
    CalendarSelectedDay = f.SelectedDay
    Unload f
End Function
Public Function CalendarSelfTest() As String
    Dim f As New ZenginCalendar
    Load f
    If f.SelectedDay <> Format$(Date, "yyyy-mm-dd") Then Err.Raise 5, , "Calendar initial day"
    f.Handle "next"
    f.Handle "today"
    If f.SelectedDay <> Format$(Date, "yyyy-mm-dd") Then Err.Raise 5, , "Calendar today"
    Unload f
    CalendarSelfTest = "PASS: calendar initialized to current month and today / navigation"
End Function
