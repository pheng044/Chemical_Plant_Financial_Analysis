Attribute VB_Name = "add_navigation"
Sub add_nav()
' ----------------------------------------------------------------------------
' Patrick Heng
' 2026-13-04
' Subroutine to add navigation system to Excel workbook
' ----------------------------------------------------------------------------

' Variable declaration
Dim ws As Worksheet
Dim i As Integer
Dim N As Integer
Dim ws_name As String

' Initialize index variable
i = 1

' If 'Home Sheet' already exists, skip delete step
On Error GoTo skip
    ThisWorkbook.Worksheets("Home Sheet").Delete
skip:

' Add home sheet for navigation
Sheets.Add(Before:=Sheets(1)).Name = "Home Sheet"

' Header for home sheet
Sheets("Home Sheet").Cells(2, 2) = "CLICK TO NAVIGATE"

' Loop through all sheets to get names, print to home sheet
For Each ws In Worksheets
    Sheets("Home Sheet").Cells(i + 2, 2) = ws.Name
    ws.Cells.EntireColumn.AutoFit
    i = i + 1
Next ws

' Total number of sheets
N = i - 1

' Loop through all sheets to add hyperlinks
For i = 1 To N
    ' Add hyperlink to home sheet
    ws_name = Sheets("Home Sheet").Cells(2 + i, 2).Value()
    Sheets("Home Sheet").Hyperlinks.Add _
        Anchor:=Sheets("Home Sheet").Cells(2 + i, 2), _
        Address:="", _
        SubAddress:="'" & ws_name & "'!A1", _
        TextToDisplay:=ws_name
        
    ' Don't add a back button for the home sheet
    If ws_name <> "Home Sheet" Then
        ' If Back button already exists, do not add another
        If Sheets(ws_name).Cells(2, 2) <> "BACK TO HOME SHEET" Then
            ' Add buffer columns/rows for prettiness
            Sheets(ws_name).Columns("A:B").Insert
            Sheets(ws_name).Rows("1:2").Insert Shift:=xlDown
            ' Add back button to home sheet
            Sheets(ws_name).Hyperlinks.Add _
                Anchor:=Sheets(ws_name).Cells(2, 2), _
                Address:="", _
                SubAddress:="'Home Sheet'!B2", _
                TextToDisplay:="BACK TO HOME SHEET"
                
            Sheets(ws_name).Cells.EntireColumn.AutoFit
        End If
    End If
Next i

' Autofit home sheet navigation column
Sheets("Home Sheet").Columns("B").AutoFit

End Sub
