Attribute VB_Name = "Module1"
Public Sub Pull_CF()


Sheets("Revenue Forecast CF").Select
lr = Sheets("Revenue Forecast CF").UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
If lr <> 1 Then

With Sheets("Revenue Forecast CF").UsedRange
        .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0).ClearContents
    End With
    End If
    
    Dim wbk As Workbook
    Dim myPath As String
    Dim myFile As String
    Dim FolderName As String
     
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
          
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
        Lastrow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    f = Lastrow + 1
    r = Lastrow + 1
    lr = 315
    lst = 316
    Do While Len(myFile) > 0
Set wbk = Workbooks.Open(myPath & myFile)
           
    ' copy "total revenue current year"
    wbk.Sheets("Output ").Range("K14:AA329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D10")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    r = r + lr + 1
    
    ' copy "total revenue coming years"
    wbk.Sheets("Output ").Range("AB14:AV329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D10")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    r = r + lr + 1
           
    ' copy "revenue from operations current year"
    wbk.Sheets("Output ").Range("AZ14:BP329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D2")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    r = r + lr + 1
    
    ' copy "revenue from operations coming years"
    wbk.Sheets("Output ").Range("BQ14:CK329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D2")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    r = r + lr + 1
    
    ' copy "training/recruitment revenue current year"
    wbk.Sheets("Output ").Range("EE14:EU329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D3")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    r = r + lr + 1
    
    ' copy "training/recruitment revenue coming years"
    wbk.Sheets("Output ").Range("EV14:FP329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D3")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    r = r + lr + 1
    
    ' copy "consulting/migrations revenue current year"
    wbk.Sheets("Output ").Range("FU14:GK329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D4")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    
    ' copy "consulting/migrations revenue coming years"
    wbk.Sheets("Output ").Range("GL14:HF329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D4")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    
    ' copy "recoveries (billable expenses) current year"
    wbk.Sheets("Output ").Range("HK14:IA329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D5")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    
    ' copy "recoveries (billable expenses) coming years"
    wbk.Sheets("Output ").Range("IB14:IV329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D5")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    
    ' copy "less: stock compensation expense current year"
    wbk.Sheets("Output ").Range("JA14:JQ329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D6")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    
    ' copy "less: stock compensation expense coming years"
    wbk.Sheets("Output ").Range("JR14:KL329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D6")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    
    ' copy "less: service credits current year"
    wbk.Sheets("Output ").Range("KQ14:LG329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D7")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    
    ' copy "less: service credits coming years"
    wbk.Sheets("Output ").Range("LH14:MB329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D7")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    
    ' copy "less: CPC fees current year"
    wbk.Sheets("Output ").Range("MG14:MW329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D8")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    
    ' copy "less: CPC fees coming years"
    wbk.Sheets("Output ").Range("MX14:NR329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = wbk.Sheets("Output ").Range("D8")
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    
    ' copy "weighted FTEs current year"
    wbk.Sheets("FTE").Range("AHG13:AHW219,AHG442:AHW550").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = "W-FTEs"
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    r = r + lr + 1
    
    ' copy "weighted FTEs coming years"
    wbk.Sheets("FTE").Range("AHX13:AIR219,AHX442:AIR550").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = "W-FTEs"
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    r = r + lr + 1
    
    ' copy "COLA% Working current year"
    wbk.Sheets("COLA Working").Range("BA13:BQ328").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = "COLA%"
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    r = r + lr + 1
    
     ' copy "COLA% Working coming years"
    wbk.Sheets("COLA Working").Range("BR13:CL328").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = "COLA%"
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    r = r + lr + 1
    
    ' copy "COLA USD Working current year"
    wbk.Sheets("COLA Working").Range("CQ13:DG328").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = "COLA$$"
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2018"
    r = r + lr + 1
    
    ' copy "COLA USD Working coming years"
    wbk.Sheets("COLA Working").Range("DH13:EB328").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Cells(r, 13).PasteSpecial Paste:=xlPasteValues
    Sheets("Revenue Forecast CF").Range(Cells(r, 35), Cells(r + lr, 35)) = "COLA$$"
    Sheets("Revenue Forecast CF").Range(Cells(r, 34), Cells(r + lr, 34)) = "2019"
    r = r + lr + 1
    
    ' copy common fields"
    wbk.Sheets("Output ").Range("A14:J329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Range(Cells(f, 3), Cells(r - 1, 3)).PasteSpecial Paste:=xlPasteValues
    wbk.Sheets("Output ").Range("NU14:NU329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Range(Cells(f, 36), Cells(r - 1, 36)).PasteSpecial Paste:=xlPasteValues
    wbk.Sheets("Output ").Range("AY14:AY329").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Range(Cells(f, 37), Cells(r - 1, 37)).PasteSpecial Paste:=xlPasteValues
    wbk.Sheets("FTE").Range("H13:H328").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Range(Cells(f, 38), Cells(r - 1, 38)).PasteSpecial Paste:=xlPasteValues
    wbk.Sheets("FTE").Range("J13:J328").Copy
    Windows("FPnA Revenue Dashboard 1.2.xlsb").Activate
    Sheets("Revenue Forecast CF").Range(Cells(f, 39), Cells(r - 1, 39)).PasteSpecial Paste:=xlPasteValues
    
    Sheets("Revenue Forecast CF").Range(Cells(f, 1), Cells(r - 1, 1)) = wbk.Sheets("Output ").Range("B2")
    Sheets("Revenue Forecast CF").Range(Cells(f, 2), Cells(r - 1, 2)) = wbk.Sheets("Output ").Range("B3")
    Sheets("Revenue Forecast CF").Range(Cells(f, 40), Cells(r - 1, 40)) = "CF"
    f = r
    
    
    Application.CutCopyMode = False
    wbk.Close True
    myFile = Dir
Loop

r = r
'delete empty rows
Sheets("Revenue Forecast CF").Select
Range("A1").Select
Selection.AutoFilter
Sheets("Revenue Forecast CF").Range(Cells(1, 1), Cells(r, 40)).AutoFilter Field:=3, Criteria1:=Array( _
        "0", "Process Names", "="), Operator:=xlFilterValues
Range(Cells(2, 1), Cells(r, 1)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
Range("A1").Select
Selection.AutoFilter


Range("A1").Select


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
End Sub
