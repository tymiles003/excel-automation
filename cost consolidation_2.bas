Attribute VB_Name = "Module1"
'Developed by Pankaj Kumar
'macro for extracting budget from different cost template files, and consolidate them rowwise
'last updated- 4/10/2017

Public Sub consolidate_budget_cost()
Dim bpo_budget As Worksheet
Set bpo_budget = ThisWorkbook.Sheets("BPO Budget Dump")

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

'lastrow for clearing previous contents
lr = bpo_budget.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
If lr <> 2 Then
bpo_budget.Range("a3:akz" & lr).ClearContents
End If

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
    
    'loop through all cost files in the folder
Do While Len(myFile) > 0
    Set wbk = Workbooks.Open(myPath & myFile)
    
    'array of sheets from which data is to extracted- sht_list(counter)
    Dim sht_list() As Variant
    sht_count = wbk.Worksheets.Count - 6
    ReDim sht_list(sht_count)
    sht_indx = 5
    For counter = 0 To sht_count
        sht_list(counter) = wbk.Worksheets(sht_indx).Name
        sht_indx = sht_indx + 1
    Next counter
    
    'loop through the array of sheets
    Dim this_sht As Worksheet
    For x = 0 To counter - 1
    Set this_sht = wbk.Sheets(sht_list(x))
    
    'last row for consolidated file
    next_row_cons = bpo_budget.Cells(Rows.Count, 3).End(xlUp).Row + 1
        
    'select data- copy & paste
    bpo_budget.Range("C" & next_row_cons) = this_sht.Range("B2")
    bpo_budget.Range("F" & next_row_cons) = this_sht.Range("B3")
    bpo_budget.Range("H" & next_row_cons) = this_sht.Range("B4")
    
    next_col_cons = 24
    'revenue
    For row_this_sht = 8 To 14
    this_sht.Range("CB" & row_this_sht & ":" & "CN" & row_this_sht).Copy
    Windows("Consolidated BPO F&A_v0.xlsb").Activate
    bpo_budget.Cells(next_row_cons, next_col_cons).PasteSpecial Paste:=xlPasteValues
    next_col_cons = next_col_cons + 14
    Next row_this_sht
    'fte
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB129:CN129").Value
    next_col_cons = next_col_cons + 14
    'headcount
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB133:CN133").Value
    next_col_cons = next_col_cons + 14
    'mei
    next_col_cons = next_col_cons + 14
    'seats
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB155:CN155").Value
    next_col_cons = next_col_cons + 14
    'seat utilization
    next_col_cons = next_col_cons + 14
    'expenses- operations
    For row_this = 20 To 34
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB" & row_this & ":CN" & row_this).Value
    next_col_cons = next_col_cons + 14
    Next row_this
    'total expenses
    next_col_cons = next_col_cons + 14
    'telecom
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB48:CN48").Value
    next_col_cons = next_col_cons + 14
    'other telecom cost
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB54:CN54").Value
    next_col_cons = next_col_cons + 14
    'total telecom
    next_col_cons = next_col_cons + 14
    'technology
    For row_this = 49 To 52
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB" & row_this & ":CN" & row_this).Value
    next_col_cons = next_col_cons + 14
    Next row_this
    'total technology
    next_col_cons = next_col_cons + 14
    'transition cost
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB53:CN53").Value
    next_col_cons = next_col_cons + 14
    'allocated technology cost
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB57:CN57").Value
    next_col_cons = next_col_cons + 14
    'allocated telecom cost
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB58:CN58").Value
    next_col_cons = next_col_cons + 14
    'facility ops cost
    For row_this = 66 To 75
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB" & row_this & ":CN" & row_this).Value
    next_col_cons = next_col_cons + 14
    Next row_this
    'facility opex
    next_col_cons = next_col_cons + 14
    'allocated facility cost
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB79:CN79").Value
    next_col_cons = next_col_cons + 14
    'unallocated facility opex
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB120:CN120").Value
    next_col_cons = next_col_cons + 14
    'non-billable t&e
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB86:CN86").Value
    next_col_cons = next_col_cons + 14
    'non-billable travel
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB87:CN87").Value
    next_col_cons = next_col_cons + 14
    'billable t&e
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB91:CN91").Value
    next_col_cons = next_col_cons + 14
    'billable others
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB92:CN92").Value
    next_col_cons = next_col_cons + 14
    'total t&e
    next_col_cons = next_col_cons + 14
    'consulting
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB108:CN108").Value
    next_col_cons = next_col_cons + 14
    'professional fees
    For row_this = 39 To 43
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB" & row_this & ":CN" & row_this).Value
    next_col_cons = next_col_cons + 14
    Next row_this
    'misc expenses
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB105:CN105").Value
    next_col_cons = next_col_cons + 14
    'cost savings
    bpo_budget.Range(Cells(next_row_cons, next_col_cons), Cells(next_row_cons, next_col_cons + 12)).Value = this_sht.Range("CB109:CN109").Value
    next_col_cons = next_col_cons + 14
    'cogs
    next_col_cons = next_col_cons + 14
    'gm
    next_col_cons = next_col_cons + 14
    'agm
    next_col_cons = next_col_cons + 14
    'agm$
    next_col_cons = next_col_cons + 14
    
    'next sheet of the file
    Next x
    
    Application.CutCopyMode = False
    wbk.Close True
    myFile = Dir
    
'next file in the folder
Loop
bpo_budget.Range("B3").Select

last_row_cons = bpo_budget.Cells(Rows.Count, 3).End(xlUp).Row
If last_row_cons <> 2 Then
'columns which have autofill formulas
bpo_budget.Range(Cells(3, 10), Cells(last_row_cons, 22)).Formula = "=X3+AL3+AZ3+BN3+CB3+CP3+DD3"    'total revenue
bpo_budget.Range(Cells(3, 150), Cells(last_row_cons, 162)).Formula = "=IFERROR(EF3/DR3,0)"  'mei average
bpo_budget.Range(Cells(3, 178), Cells(last_row_cons, 190)).Formula = "=IFERROR(EF3/FH3,0)"  'seat utilization
bpo_budget.Range(Cells(3, 402), Cells(last_row_cons, 414)).Formula = "=GJ3+GX3+HL3+HZ3+IN3+JB3+JP3+KD3+KR3+LF3+LT3+MH3+MV3+NJ3+NX3"    'expenses total
bpo_budget.Range(Cells(3, 444), Cells(last_row_cons, 456)).Formula = "=OZ3+PN3"  'telecom total
bpo_budget.Range(Cells(3, 514), Cells(last_row_cons, 526)).Formula = "=QP3+RD3+RR3+SF3"  'technology total
bpo_budget.Range(Cells(3, 710), Cells(last_row_cons, 722)).Formula = "=UX3+VL3+VZ3+WN3+XB3+XP3+YD3+YR3+ZF3+ZT3"  'facility opex
bpo_budget.Range(Cells(3, 808), Cells(last_row_cons, 820)).Formula = "=ABX3+ACZ3+ADN3+ACL3"  'total t&e
bpo_budget.Range(Cells(3, 934), Cells(last_row_cons, 946)).Formula = "=OL3+QB3+ST3+AAH3+AEB3+AEP3+AFD3+AHV3+AIJ3+TH3+TV3+UJ3+AAV3+AFR3+AGF3+AGT3+AHH3" 'cogs
bpo_budget.Range(Cells(3, 948), Cells(last_row_cons, 960)).Formula = "=IFERROR(1-(AIX3/J3),0)"  'gm
bpo_budget.Range(Cells(3, 962), Cells(last_row_cons, 974)).Formula = "=IFERROR(1-((AIX3+ABJ3)/J3),0)"  'agm
bpo_budget.Range(Cells(3, 976), Cells(last_row_cons, 988)).Formula = "=J3-AIX3-ABJ3"  'agm$

'sorting on the basis of client name
Range("A3:AKZ" & last_row_cons).Select
Selection.Sort Key1:=Range("C3"), Order1:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
End If
'home
bpo_budget.Range("B3").Select

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub
