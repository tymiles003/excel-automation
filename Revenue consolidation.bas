Attribute VB_Name = "Module1"
'Revenue Forecast Consolidation
'Developed by Pankaj Kumar
'VBA code to consolidate data from multiple revenue forecast files
'Last updated- 15/9/2017

Public Sub Consolidate_Rev_Forecast()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
mysheets = Array("Modify Data", "FTE Forecast- 2017", "Rev Forecast Committed", "Passthrough Revenue", "Opportunities Included", "Revenue Forecast Final", "QoQ Details", "MoM Details with Location", "FTE Forecast- 2018", "Rev Forecast- 2018")
Call clr_sheet(mysheets)

    
    Dim wbk As Workbook
    Dim myPath As String
    Dim myFile As String
    Dim FolderName As String
    
    'dialog box for selecting folder containing files
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
        On Error Resume Next
        FolderName = .SelectedItems(1)
        Err.Clear
        On Error GoTo 0
    End With
    myPath = FolderName
    If Right(myPath, 1) <> "\" Then myPath = myPath & "\"
    myFile = Dir(myPath & "*.xls*")
    
    'loop through all files in the folder
Do While Len(myFile) > 0
    Set wbk = Workbooks.Open(myPath & myFile)
    mysheets = Array("Modify Data", "FTE Forecast- 2017", "Rev Forecast Committed", "Passthrough Revenue", "Opportunities Included", "Revenue Forecast Final", "QoQ Details", "MoM Details with Location", "FTE Forecast- 2018", "Rev Forecast- 2018")

    'copies and pastes the data
    Dim this_sht As Worksheet
    For x = 0 To 9
    Set this_sht = wbk.Sheets(mysheets(x))

    'last row & col for selected file
    lr = this_sht.Cells(Rows.Count, 1).End(xlUp).Row
    lc = this_sht.UsedRange.Columns.Count - 50

    'last row for consolidated file
    next_row = ThisWorkbook.Sheets(mysheets(x)).Cells(Rows.Count, 1).End(xlUp).Row + 1

    'address for cells
    strt_range = Cells(5, 1).Address()
    end_range = Cells(lr, lc - 1).Address()
    strt_range_key = Cells(5, lc).Address()
    end_range_key = Cells(lr, lc).Address()

    'paste with links
    this_sht.Range(strt_range & ":" & end_range).Copy
    Windows("Consolidated Rev Forecast.xlsb").Activate
    ThisWorkbook.Sheets(mysheets(x)).Cells(next_row, 1).PasteSpecial Paste:=xlPasteAll
    this_sht.Range(strt_range_key & ":" & end_range_key).Copy
    Windows("Consolidated Rev Forecast.xlsb").Activate
    ThisWorkbook.Sheets(mysheets(x)).Cells(next_row, lc).PasteSpecial Paste:=xlPasteValues
    Next x
    
    Application.CutCopyMode = False
    wbk.Close True
    myFile = Dir
Loop

mysheets = Array("Modify Data", "FTE Forecast- 2017", "Rev Forecast Committed", "Passthrough Revenue", "Opportunities Included", "Revenue Forecast Final", "QoQ Details", "MoM Details with Location", "FTE Forecast- 2018", "Rev Forecast- 2018")
Call check_sheet(mysheets)

ThisWorkbook.Sheets("Modify Data").Select
Range("A4").Select

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

'break links
Dim Links As Variant
Links = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
For i = 1 To UBound(Links)
ActiveWorkbook.BreakLink _
    Name:=Links(i), _
    Type:=xlLinkTypeExcelLinks
Next i

End Sub

Function check_sheet(mysheets As Variant)
'checks the contents of the final sheets
Dim con_rev_final As Worksheet
For x = 0 To 9
Set con_rev_final = ThisWorkbook.Sheets(mysheets(x))
lr_final = con_rev_final.Cells(Rows.Count, 1).End(xlUp).Row
lc_final = con_rev_final.UsedRange.Columns.Count
lst_cell_final = Cells(lr_final, lc_final).Address()
lst_cell_col1 = Cells(lr_final, 1).Address()

con_rev_final.Select
Range("a4:" & lst_cell_final).AutoFilter Field:=lc_final, Criteria1:=Array("0", "-"), Operator:=xlFilterValues
If Cells(Rows.Count, 1).End(xlUp).Row = 4 Then
Range("A4").AutoFilter
Range("A4").Select
Else
Range("a5:" & lst_cell_col1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
Range("A4").AutoFilter
End If

'sorting on the basis of client name & sort key
cell_sort_key = Cells(5, lc_final - 2).Address()
Range("A5:" & lst_cell_final).Select
Selection.Sort Key1:=Range("B5"), Order1:=xlAscending, Key2:=Range(cell_sort_key) _
        , Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
Next x
End Function

Function clr_sheet(mysheets As Variant)
'clear the contents of the sheets
Dim con_rev As Worksheet
For x = 0 To 9
Set con_rev = ThisWorkbook.Sheets(mysheets(x))
clr_row = con_rev.Cells(Rows.Count, 1).End(xlUp).Row
clr_col = con_rev.UsedRange.Columns.Count
lst_cell = Cells(clr_row, clr_col).Address()
con_rev.Range("A5:XFD1048576").Interior.Color = RGB(255, 255, 255)
con_rev.Range("A5:XFD1048576").Borders(xlEdgeLeft).LineStyle = xlNone
con_rev.Range("A5:XFD1048576").Borders(xlEdgeTop).LineStyle = xlNone
con_rev.Range("A5:XFD1048576").Borders(xlEdgeRight).LineStyle = xlNone
con_rev.Range("A5:XFD1048576").Borders(xlEdgeBottom).LineStyle = xlNone
con_rev.Range("A5:XFD1048576").Borders(xlInsideVertical).LineStyle = xlNone
con_rev.Range("A5:XFD1048576").Borders(xlInsideHorizontal).LineStyle = xlNone
con_rev.Range("A5:XFD1048576").Borders(xlDiagonalUp).LineStyle = xlNone
con_rev.Range("A5:XFD1048576").Borders(xlDiagonalDown).LineStyle = xlNone
If clr_row <> 4 Then
con_rev.Range("a5:" & lst_cell).ClearContents
End If
Next x
End Function

