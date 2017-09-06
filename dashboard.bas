Attribute VB_Name = "Module1"
'Developed by Pankaj Kumar
'VBA code to import data from revenue consolidated file and prepare revenue dashboard
'Last updated- 3/9/2017

Public Sub Import_Data()

'data sheet and dashboard sheet
Dim data As Worksheet
Set data = ThisWorkbook.Sheets("Data")
Dim dashboard As Worksheet
Set dashboard = ThisWorkbook.Sheets("Dashboard")

'lastrow for clearing previous contents in data sheet
lr = data.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
If lr <> 4 Then
data.range("A1:AO" & lr).ClearContents
End If

'revenue consolidated file
Dim cons_rev As Workbook
Dim myFile As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
    
'dialog box for selecting consolidated revenue file
With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Please select the consolidated forecast file."
        .Filters.Clear
        .Filters.Add "Excel 2010", "*.xls?"
        .Show
        On Error Resume Next
        myFile = .SelectedItems(1)
        Err.Clear
        On Error GoTo 0
End With
'copy the contents from raw data and historic data sheets and paste to data
Set cons_rev = Workbooks.Open(myFile)
cons_rev.Sheets("RawData").range("A1:Z350").Copy
Windows("Revenue Dashboard.xlsb").Activate
data.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
data.Visible = xlSheetHidden
Application.CutCopyMode = False
cons_rev.Close True

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Public Sub AllClients()
'macros for all clients
Dim cht_old As Object
For Each cht_old In ActiveSheet.ChartObjects
cht_old.Delete
Next

'take combo input
chart_no = Sheets("Dashboard").range("ba2")
Select Case chart_no
Case 1  'vertical view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k2:n9", "a2:d9", "a12:d19", "f2:i9", "f12:i19")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)

Case 2  'ncr view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k22:n29", "a22:d29", "a32:d39", "f22:i29", "f32:i39")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)
    
Case 3  'delivery-region view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k42:n47", "a42:d47", "a50:d55", "f42:i47", "f50:i55")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)
    
Case 4  'category view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k58:n61", "a58:d61", "a65:d68", "f58:i61", "f65:i68")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)
    
Case 5  'QoQ & MoM views
    lft = Array(20, 520)
    tp = Array(100, 100)
    src = Array("e71:g75", "i71:k83")
    titl = Array("QoQ View", "MoM View")
    Call lin_cht(lft, tp, src, titl)
    
Case 6  'Top clients revenue
    lft = 20
    tp = 100
    src = "a71:c91"
    titl = "Top Revenue Clients"
    Call col_cht2(lft, tp, src, titl)
    
Case 7  'Client fragmentation
    lft = 200
    tp = 100
    src = "i86:m92"
    titl = "Client Fragmentation"
    Call col_cht3(lft, tp, src, titl)
    
Case 8  'top hits & misses
    lft = Array(120, 120)
    tp = Array(100, 420)
    src = Array("a93:d103", "f93:i103")
    titl = Array("Top Hits", "Top Misses")
    Call col_cht4(lft, tp, src, titl)

Case 9  'historic data of top clients
    lft = 20
    tp = 100
    src = "o71:s91"
    titl = "Historic Data"
    Call col_cht5(lft, tp, src, titl)
    
End Select
End Sub

Function col_cht(lft, tp, src, titl As Variant)
'column charts function- vertical, ncr, region, category
Dim new_cht As Object
For x = 0 To 4
Set new_cht = ActiveSheet.ChartObjects.Add(Left:=lft(x), Width:=360, Top:=tp(x), Height:=216)
With new_cht
.Chart.SetSourceData Source:=Worksheets("Data").range(src(x))
.Chart.Type = xlColumn
.Chart.SetElement (msoElementLegendBottom)
.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
.Chart.SetElement (msoElementDataLabelOutSideEnd)
.Chart.SetElement (msoElementPrimaryValueAxisNone)
.Chart.SeriesCollection(1).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(1).DataLabels.Orientation = xlUpward
.Chart.SeriesCollection(2).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(2).DataLabels.Orientation = xlUpward
.Chart.SeriesCollection(3).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(3).DataLabels.Orientation = xlUpward
.Chart.HasTitle = True
.Chart.ChartTitle.Text = titl(x)
.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(79, 129, 189)
.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(228, 108, 10)
.Chart.SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(127, 127, 127)
.Chart.ChartGroups(1).Overlap = -25
End With
Next x
End Function
Function lin_cht(lft, tp, src, titl As Variant)
'line charts function- mom, qoq
Dim new_cht As Object
For x = 0 To 1
Set new_cht = ActiveSheet.ChartObjects.Add(Left:=lft(x), Width:=360, Top:=tp(x), Height:=216)
With new_cht
.Chart.SetSourceData Source:=Worksheets("Data").range(src(x))
.Chart.Type = xlLine
.Chart.SetElement (msoElementLegendBottom)
.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
.Chart.SetElement (msoElementDataLabelTop)
.Chart.SetElement (msoElementPrimaryValueAxisNone)
.Chart.SeriesCollection(1).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(1).DataLabels.Position = xlLabelPositionBelow
.Chart.SeriesCollection(2).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(2).DataLabels.Position = xlLabelPositionTop
.Chart.HasTitle = True
.Chart.ChartTitle.Text = titl(x)
.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(79, 129, 189)
.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(228, 108, 10)
End With
Next x
End Function
Function col_cht2(lft, tp, src, titl As Variant)
'column charts function- top clients
Dim new_cht As Object
Set new_cht = ActiveSheet.ChartObjects.Add(Left:=lft, Width:=938, Top:=tp, Height:=288)
With new_cht
.Chart.SetSourceData Source:=Worksheets("Data").range(src)
.Chart.Type = xlColumn
.Chart.SetElement (msoElementLegendBottom)
.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
.Chart.SetElement (msoElementDataLabelOutSideEnd)
.Chart.SetElement (msoElementPrimaryValueAxisNone)
.Chart.SeriesCollection(1).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(1).DataLabels.Orientation = xlUpward
.Chart.SeriesCollection(2).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(2).DataLabels.Orientation = xlUpward
.Chart.HasTitle = True
.Chart.ChartTitle.Text = titl
.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(79, 129, 189)
.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(228, 108, 10)
.Chart.ChartGroups(1).Overlap = 0
End With
End Function
Function col_cht3(lft, tp, src, titl As Variant)
'column charts function- client fragmentation
Dim new_cht As Object
Set new_cht = ActiveSheet.ChartObjects.Add(Left:=lft, Width:=576, Top:=tp, Height:=288)
With new_cht
.Chart.SetSourceData Source:=Worksheets("Data").range(src)
.Chart.Type = xlColumn
.Chart.SetElement (msoElementLegendBottom)
.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
.Chart.SetElement (msoElementDataLabelOutSideEnd)
.Chart.SetElement (msoElementPrimaryValueAxisNone)
.Chart.SeriesCollection(1).DataLabels.NumberFormat = "0;0;"
.Chart.SeriesCollection(2).DataLabels.NumberFormat = "0;0;"
.Chart.SeriesCollection(3).DataLabels.NumberFormat = "0;0;"
.Chart.SeriesCollection(4).DataLabels.NumberFormat = "0;0;"
.Chart.HasTitle = True
.Chart.ChartTitle.Text = titl
.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(79, 129, 189)
.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(228, 108, 10)
.Chart.SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(127, 127, 127)
.Chart.SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(217, 150, 148)
.Chart.ChartGroups(1).Overlap = -25
End With
End Function
Function col_cht4(lft, tp, src, titl As Variant)
'column charts function- top hits & misses
Dim new_cht As Object
For x = 0 To 1
Set new_cht = ActiveSheet.ChartObjects.Add(Left:=lft(x), Width:=720, Top:=tp(x), Height:=288)
With new_cht
.Chart.SetSourceData Source:=Worksheets("Data").range(src(x))
.Chart.Type = xlColumn
.Chart.SetElement (msoElementLegendBottom)
.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
.Chart.SetElement (msoElementDataLabelOutSideEnd)
.Chart.SetElement (msoElementPrimaryValueAxisNone)
.Chart.SeriesCollection(1).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(1).DataLabels.Orientation = xlUpward
.Chart.SeriesCollection(2).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(2).DataLabels.Orientation = xlUpward
.Chart.SeriesCollection(3).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(3).DataLabels.Orientation = xlUpward
.Chart.HasTitle = True
.Chart.ChartTitle.Text = titl(x)
.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(79, 129, 189)
.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(228, 108, 10)
.Chart.SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(217, 150, 148)
.Chart.ChartGroups(1).Overlap = -25
End With
Next x
End Function
Function col_cht5(lft, tp, src, titl As Variant)
'column charts function- historic data
Dim new_cht As Object
Set new_cht = ActiveSheet.ChartObjects.Add(Left:=lft, Width:=938, Top:=tp, Height:=288)
With new_cht
.Chart.SetSourceData Source:=Worksheets("Data").range(src)
.Chart.Type = xlColumn
.Chart.SetElement (msoElementLegendBottom)
.Chart.SetElement (msoElementPrimaryValueGridLinesNone)
.Chart.SetElement (msoElementDataLabelOutSideEnd)
.Chart.SetElement (msoElementPrimaryValueAxisNone)
.Chart.SeriesCollection(1).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(1).DataLabels.Orientation = xlUpward
.Chart.SeriesCollection(2).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(2).DataLabels.Orientation = xlUpward
.Chart.SeriesCollection(3).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(3).DataLabels.Orientation = xlUpward
.Chart.SeriesCollection(4).DataLabels.NumberFormat = "0.0,,;0.0,,;"
.Chart.SeriesCollection(4).DataLabels.Orientation = xlUpward
.Chart.HasTitle = True
.Chart.ChartTitle.Text = titl
.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(79, 129, 189)
.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(228, 108, 10)
.Chart.SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(127, 127, 127)
.Chart.SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(217, 150, 148)
.Chart.ChartGroups(1).Overlap = -15
End With
End Function
Public Sub Ihcm()
'macros for ihcm
Dim cht_old As Object
For Each cht_old In ActiveSheet.ChartObjects
cht_old.Delete
Next

'take combo input
chart_no = Sheets("Dashboard").range("ba2")
Select Case chart_no
Case 1  'vertical view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k109:n116", "a109:d116", "a119:d126", "f109:i116", "f119:i126")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)

Case 2  'ncr view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k129:n136", "a129:d136", "a139:d146", "f129:i136", "f139:i146")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)
    
Case 3  'delivery-region view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k149:n154", "a149:d154", "a157:d162", "f149:i154", "f157:i162")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)
    
Case 4  'category view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k165:n168", "a165:d168", "a172:d175", "f165:i168", "f172:i175")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)
    
Case 5  'QoQ & MoM views
    lft = Array(20, 520)
    tp = Array(100, 100)
    src = Array("e178:g182", "i178:k190")
    titl = Array("QoQ View", "MoM View")
    Call lin_cht(lft, tp, src, titl)
    
Case 6  'Top clients revenue
    lft = 20
    tp = 100
    src = "a178:c198"
    titl = "Top Revenue Clients"
    Call col_cht2(lft, tp, src, titl)
    
Case 7  'Client fragmentation
    lft = 200
    tp = 100
    src = "i193:m199"
    titl = "Client Fragmentation"
    Call col_cht3(lft, tp, src, titl)
    
Case 8  'top hits & misses
    lft = Array(120, 120)
    tp = Array(100, 420)
    src = Array("a200:d210", "f200:i210")
    titl = Array("Top Hits", "Top Misses")
    Call col_cht4(lft, tp, src, titl)

Case 9  'historic data of top clients
    lft = 20
    tp = 100
    src = "o178:s198"
    titl = "Historic Data"
    Call col_cht5(lft, tp, src, titl)
    
End Select
End Sub

Public Sub Nonihcm()
'macros for non-ihcm
Dim cht_old As Object
For Each cht_old In ActiveSheet.ChartObjects
cht_old.Delete
Next

'take combo input
chart_no = Sheets("Dashboard").range("ba2")
Select Case chart_no
Case 1  'vertical view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k216:n223", "a216:d223", "a226:d233", "f216:i223", "f226:i233")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)

Case 2  'ncr view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k236:n243", "a236:d243", "a246:d253", "f236:i243", "f246:i253")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)
    
Case 3  'delivery-region view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k256:n261", "a256:d261", "a264:d269", "f256:i261", "f264:i269")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)
    
Case 4  'category view
    lft = Array(20, 520, 20, 520, 20)
    tp = Array(100, 100, 350, 350, 600)
    src = Array("k272:n275", "a272:d275", "a279:d282", "f272:i275", "f279:i282")
    titl = Array("FY' 17", "Q1' 17", "Q2' 17", "Q3' 17", "Q4' 17")
    Call col_cht(lft, tp, src, titl)
    
Case 5  'QoQ & MoM views
    lft = Array(20, 520)
    tp = Array(100, 100)
    src = Array("e285:g289", "i285:k297")
    titl = Array("QoQ View", "MoM View")
    Call lin_cht(lft, tp, src, titl)
    
Case 6  'Top clients revenue
    lft = 20
    tp = 100
    src = "a285:c305"
    titl = "Top Revenue Clients"
    Call col_cht2(lft, tp, src, titl)
    
Case 7  'Client fragmentation
    lft = 200
    tp = 100
    src = "i300:m306"
    titl = "Client Fragmentation"
    Call col_cht3(lft, tp, src, titl)
    
Case 8  'top hits & misses
    lft = Array(120, 120)
    tp = Array(100, 420)
    src = Array("a307:d317", "f307:i317")
    titl = Array("Top Hits", "Top Misses")
    Call col_cht4(lft, tp, src, titl)

Case 9  'historic data of top clients
    lft = 20
    tp = 100
    src = "o285:s305"
    titl = "Historic Data"
    Call col_cht5(lft, tp, src, titl)
    
End Select
End Sub



