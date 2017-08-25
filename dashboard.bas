Attribute VB_Name = "Module5"

Sub StartDashboard()
' Starts the dashboard and puts sheet title and relevant form controls for the dashboard
' title of the sheet- ie FP&A Dashboard
    Sheets("Dashboard").Select
    Cells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("S1").Select
    ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 387.75, 23.25, 223.2, 50.4 _
        ).Select
    Selection.Name = "Sheet Title"
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "F&A Dashboard"
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 13).ParagraphFormat. _
        FirstLineIndent = 0
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 13).Font
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 32
        .Name = "+mn-lt"
    End With
    ActiveSheet.Shapes.Range(Array("Sheet Title")).Select
    Selection.Cut
    Range("I1").Select
    ActiveSheet.Paste
' Combo box for chart types
    Sheets("Dashboard").Select
    ActiveSheet.DropDowns.Add(726.75, 78, 108, 14.25).Select
    Selection.Name = "ChartTypes"
    Selection.Cut
    Range("B1").Select
    ActiveSheet.Paste
    ActiveSheet.Shapes.Range(Array("ChartTypes")).Select
    With Selection
        .ListFillRange = "RawData!$R4:$R12"
        .LinkedCell = "RawData!R3"
        .DropDownLines = 9
        .Display3DShading = True
    End With
' Show All Clients button
'
    ActiveSheet.Buttons.Add(492, 86.25, 72, 14.4).Select
    Selection.Cut
    Range("B4").Select
    ActiveSheet.Paste
    Selection.Characters.Text = "All Clients"
    Selection.Name = "all"
    With Selection.Characters(Start:=1, Length:=5).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveSheet.Shapes.Range(Array("all")).Select
    Selection.OnAction = "AllClients"

'
    ActiveSheet.Buttons.Add(492, 86.25, 72, 14.4).Select
    Selection.Name = "ichm"
    Selection.Cut
    Range("D4").Select
    ActiveSheet.Paste
    Selection.Characters.Text = "IHCM Clients"
    With Selection.Characters(Start:=1, Length:=5).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveSheet.Shapes.Range(Array("ichm")).Select
    Selection.OnAction = "Ihcm"

'
    ActiveSheet.Buttons.Add(492, 86.25, 72, 14.4).Select
    Selection.Cut
    Range("F4").Select
    ActiveSheet.Paste
    Selection.Characters.Text = "Non-IHCM Clients"
    Selection.Name = "nonichm"
    With Selection.Characters(Start:=1, Length:=5).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveSheet.Shapes.Range(Array("nonichm")).Select
    Selection.OnAction = "NonIhcm"
        
    Sheets("RawData").Visible = True
    Sheets("RawData").Select
    Range("A1:N317").Select
    Selection.NumberFormat = "0.00,,"
    Range("O14:T23").Select
    Selection.NumberFormat = "0.00,,"
    Sheets("RawData").Visible = False
    Sheets("Dashboard").Select
    Range("J8:L17").Select
    Selection.NumberFormat = "0.00,,"
    
 ' dummy charts to be populated while starting dashboard
    Sheets("Dashboard").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
    ActiveSheet.Shapes("q4").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
    ActiveSheet.Shapes("fy").Line.Visible = msoFalse
    Range("R6").Select
End Sub
Sub AllClients()
ActiveSheet.Shapes.Range(Array("q1", "q2", "q3", "q4", "fy")).Select
Selection.Delete
'take combo input
chart_no = Worksheets("RawData").Range("R3")
If chart_no = 1 Then
'vertical view--------------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$2:$N$9")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$2:$D$9")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$12:$D$19")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$2:$I$9")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$12:$I$19")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
ElseIf chart_no = 2 Then '--------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$22:$N$29")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$22:$D$29")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$32:$D$39")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$22:$I$29")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$32:$I$39")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

ElseIf chart_no = 3 Then '--------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$42:$N$47")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$42:$D$47")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$50:$D$55")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$42:$I$47")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$50:$I$55")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
ElseIf chart_no = 4 Then '--------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$58:$N$61")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$58:$D$61")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$65:$D$68")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$58:$I$61")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$65:$I$68")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
ElseIf chart_no = 5 Then '--------------------------------------------------------------
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("B23").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("B24").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
'qoq view
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.Parent.Cut
    Sheets("Dashboard").Select
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
    ActiveChart.SetSourceData Source:=Range("RawData!$E$71:$G$75")
'bold title
    ActiveChart.ChartTitle.Text = "QoQ View"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'data label
    ActiveChart.SetElement (msoElementDataLabelTop)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveSheet.Shapes("q4").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementLegendBottom)
'MOM view
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.Parent.Cut
    Sheets("Dashboard").Select
    Range("L7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
    ActiveChart.SetSourceData Source:=Range("RawData!$I$71:$K$83")
'bold title
    ActiveChart.ChartTitle.Text = "MoM View"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'data label
    ActiveChart.SetElement (msoElementDataLabelTop)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    Selection.Position = xlLabelPositionBelow
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    Selection.Position = xlLabelPositionTop
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementLegendBottom)
    Range("R6").Select
ElseIf chart_no = 6 Then '--------------------------------------------------------------
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
    ActiveSheet.Shapes("q4").Line.Visible = msoFalse
'insert column chart- top 20 clients
    ActiveSheet.Shapes.AddChart2(322, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C8").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
    ActiveSheet.Shapes("fy").Height = 252
    ActiveSheet.Shapes("fy").Width = 813.6
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$71:$C$91")
'chart title
    ActiveChart.ChartTitle.Text = "Top 20 Clients"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'data label
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    ActiveChart.SetElement (msoElementLegendBottom)
    Range("R6").Select
ElseIf chart_no = 7 Then '--------------------------------------------------------------
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C31").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C32").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C33").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
    ActiveSheet.Shapes("q4").Line.Visible = msoFalse
' chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("H10").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("Summary!$E$67:$H$73")
'chart title
    ActiveChart.ChartTitle.Text = "Client fragmentation View"
'data label
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementDataLabelShow)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.ChartGroups(1).Overlap = 0
    Range("R6").Select
ElseIf chart_no = 8 Then '--------------------------------------------------------------
'column chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("D7").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$93:$D$103")
    ActiveChart.ChartTitle.Text = "Top Hits"
'data label
    ActiveSheet.Shapes("fy").Width = 720
    ActiveSheet.Shapes("fy").Height = 252
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
'column chart- top misses
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("D24").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$93:$I$103")
    ActiveChart.ChartTitle.Text = "Top Misses"
'data label
    ActiveSheet.Shapes("q4").Width = 720
    ActiveSheet.Shapes("q4").Height = 252
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("q4").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C50").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C50").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C50").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
    Range("R6").Select
ElseIf chart_no = 9 Then '--------------------------------------------------------------
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
' Combo box for chart types in historic data
    Sheets("Dashboard").Select
    ActiveSheet.DropDowns.Add(726.75, 78, 115.2, 14.4).Select
    Selection.Name = "q4"
    Selection.Cut
    Range("F1").Select
    ActiveSheet.Paste
    ActiveSheet.Shapes.Range(Array("q4")).Select
    With Selection
        .ListFillRange = "'Historic Data'!$A$5:$A$150"
        .LinkedCell = "'Historic Data'!$Q$5"
        .DropDownLines = 25
        .Display3DShading = True
    End With
    'column chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("G8").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("'Historic Data'!$P$7:$Q$11")
'data label
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Range("R6").Select
End If
End Sub
Sub Ihcm()
ActiveSheet.Shapes.Range(Array("q1", "q2", "q3", "q4", "fy")).Select
Selection.Delete
'take combo input
chart_no = Worksheets("RawData").Range("R3")
If chart_no = 1 Then
'vertical view--------------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$109:$N$116")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$109:$D$116")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$119:$D$127")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$109:$I$116")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$119:$I$127")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
ElseIf chart_no = 2 Then '--------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$129:$N$136")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$129:$D$136")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$139:$D$146")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$129:$I$136")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$139:$I$146")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

ElseIf chart_no = 3 Then '--------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$149:$N$154")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$149:$D$154")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$157:$D$162")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$149:$I$154")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$157:$I$162")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
ElseIf chart_no = 4 Then '--------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$165:$N$168")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$165:$D$168")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$172:$D$175")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$165:$I$168")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$172:$I$175")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
ElseIf chart_no = 5 Then '--------------------------------------------------------------
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("B23").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("B24").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
'qoq view
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.Parent.Cut
    Sheets("Dashboard").Select
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
    ActiveChart.SetSourceData Source:=Range("RawData!$E$178:$G$182")
'bold title
    ActiveChart.ChartTitle.Text = "QoQ View"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'data label
    ActiveChart.SetElement (msoElementDataLabelTop)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveSheet.Shapes("q4").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementLegendBottom)
'MOM view
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.Parent.Cut
    Sheets("Dashboard").Select
    Range("L7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
    ActiveChart.SetSourceData Source:=Range("RawData!$I$178:$K$190")
'bold title
    ActiveChart.ChartTitle.Text = "MoM View"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'data label
    ActiveChart.SetElement (msoElementDataLabelTop)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    Selection.Position = xlLabelPositionBelow
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    Selection.Position = xlLabelPositionTop
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementLegendBottom)
    Range("R6").Select
ElseIf chart_no = 6 Then '--------------------------------------------------------------
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
    ActiveSheet.Shapes("q4").Line.Visible = msoFalse
'insert column chart- top 20 clients
    ActiveSheet.Shapes.AddChart2(322, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C8").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
    ActiveSheet.Shapes("fy").Height = 252
    ActiveSheet.Shapes("fy").Width = 813.6
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$178:$C$198")
'chart title
    ActiveChart.ChartTitle.Text = "Top 20 Clients"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'data label
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.SetElement (msoElementLegendBottom)
    Range("R6").Select
ElseIf chart_no = 7 Then '--------------------------------------------------------------
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C31").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C32").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C33").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
    ActiveSheet.Shapes("q4").Line.Visible = msoFalse
' chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("H10").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("Summary!$E$220:$H$226")
'chart title
    ActiveChart.ChartTitle.Text = "Client fragmentation View"
'data label
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementDataLabelShow)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.ChartGroups(1).Overlap = 0
    Range("R6").Select
ElseIf chart_no = 8 Then '--------------------------------------------------------------
'column chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("D7").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$200:$D$210")
    ActiveChart.ChartTitle.Text = "Top Hits"
'data label
    ActiveSheet.Shapes("fy").Width = 720
    ActiveSheet.Shapes("fy").Height = 252
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
'column chart- top misses
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("D24").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$200:$I$210")
    ActiveChart.ChartTitle.Text = "Top Misses"
'data label
    ActiveSheet.Shapes("q4").Width = 720
    ActiveSheet.Shapes("q4").Height = 252
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("q4").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabels.Orientation = xlUpward
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C50").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C50").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C50").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
    Range("R6").Select
ElseIf chart_no = 9 Then '--------------------------------------------------------------
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
' Combo box for chart types in historic data
    Sheets("Dashboard").Select
    ActiveSheet.DropDowns.Add(726.75, 78, 115.2, 14.4).Select
    Selection.Name = "q4"
    Selection.Cut
    Range("F4").Select
    ActiveSheet.Paste
    ActiveSheet.Shapes.Range(Array("q4")).Select
    With Selection
        .ListFillRange = "'Historic Data'!$J$5:$J$150"
        .LinkedCell = "'Historic Data'!$Q$15"
        .DropDownLines = 25
        .Display3DShading = True
    End With
    'column chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("G8").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("'Historic Data'!$P$18:$Q$22")
'data label
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Range("R6").Select
End If
End Sub
Sub NonIhcm()
ActiveSheet.Shapes.Range(Array("q1", "q2", "q3", "q4", "fy")).Select
Selection.Delete
chart_no = Worksheets("RawData").Range("R3")
If chart_no = 1 Then
'vertical view--------------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$216:$N$223")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$216:$D$223")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$226:$D$233")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$216:$I$223")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$226:$I$233")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
ElseIf chart_no = 2 Then '--------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$236:$N$243")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$236:$D$243")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$246:$D$253")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$236:$I$243")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$246:$I$253")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

ElseIf chart_no = 3 Then '--------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$256:$N$261")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$256:$D$261")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$264:$D$269")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$256:$I$261")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$264:$I$269")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
ElseIf chart_no = 4 Then '--------------------------------------------------------------
'insert column chart- fy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$K$272:$N$275")
'chart title
    ActiveChart.ChartTitle.Text = "FY' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q1
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$272:$D$275")
'chart title
    ActiveChart.ChartTitle.Text = "Q1' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
    
'insert column chart- q2
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$269:$D$282")
'chart title
    ActiveChart.ChartTitle.Text = "Q2' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q3
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("M22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$272:$I$275")
'chart title
    ActiveChart.ChartTitle.Text = "Q3' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select

'insert column chart- q4
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
'postion of chart
    ActiveChart.Parent.Cut
    Range("B37").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$269:$I$282")
'chart title
    ActiveChart.ChartTitle.Text = "Q4' 17"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
'reducing gaps between the two series lines
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    Selection.NumberFormat = "0.00,,"
'no vertical axis
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    Range("R6").Select
ElseIf chart_no = 5 Then '--------------------------------------------------------------
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("B22").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("B23").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("B24").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
'qoq view
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.Parent.Cut
    Sheets("Dashboard").Select
    Range("B7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
    ActiveChart.SetSourceData Source:=Range("RawData!$E$285:$G$289")
'bold title
    ActiveChart.ChartTitle.Text = "QoQ View"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'data label
    ActiveChart.SetElement (msoElementDataLabelTop)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveSheet.Shapes("q4").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementLegendBottom)
'MOM view
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.Parent.Cut
    Sheets("Dashboard").Select
    Range("L7").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "fy"
    ActiveChart.SetSourceData Source:=Range("RawData!$I$285:$K$297")
'bold title
    ActiveChart.ChartTitle.Text = "MoM View"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'data label
    ActiveChart.SetElement (msoElementDataLabelTop)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    Selection.Position = xlLabelPositionBelow
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,,"
    Selection.Position = xlLabelPositionTop
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementLegendBottom)
    Range("R6").Select
ElseIf chart_no = 6 Then '--------------------------------------------------------------
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
    ActiveSheet.Shapes("q4").Line.Visible = msoFalse
'insert column chart- top 20 clients
    ActiveSheet.Shapes.AddChart2(322, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C8").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
    ActiveSheet.Shapes("fy").Height = 252
    ActiveSheet.Shapes("fy").Width = 813.6
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$285:$C$305")
'chart title
    ActiveChart.ChartTitle.Text = "Top 20 Clients"
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'data label
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.SetElement (msoElementLegendBottom)
    Range("R6").Select
ElseIf chart_no = 7 Then '--------------------------------------------------------------
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C31").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C32").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C33").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q4"
    ActiveSheet.Shapes("q4").Line.Visible = msoFalse
' chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("H10").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("Summary!$E$374:$H$380")
'chart title
    ActiveChart.ChartTitle.Text = "Client fragmentation View"
'data label
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementDataLabelShow)
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.ChartGroups(1).Overlap = 0
    Range("R6").Select
ElseIf chart_no = 8 Then '--------------------------------------------------------------
'column chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("D7").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$A$307:$D$317")
    ActiveChart.ChartTitle.Text = "Top Hits"
'data label
    ActiveSheet.Shapes("fy").Width = 720
    ActiveSheet.Shapes("fy").Height = 252
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
'column chart- top misses
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("D24").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "q4"
'data source
    ActiveChart.SetSourceData Source:=Range("RawData!$F$307:$I$317")
    ActiveChart.ChartTitle.Text = "Top Misses"
'data label
    ActiveSheet.Shapes("q4").Width = 720
    ActiveSheet.Shapes("q4").Height = 252
    ActiveChart.SetElement (msoElementLegendBottom)
    ActiveSheet.Shapes("q4").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    Selection.Format.TextFrame2.Orientation = msoTextOrientationUpward
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Selection.Orientation = xlUpward
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabels.Orientation = xlUpward
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C50").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C50").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C50").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
    Range("R6").Select
ElseIf chart_no = 9 Then '--------------------------------------------------------------
' dummy
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q1"
    ActiveSheet.Shapes("q1").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q2"
    ActiveSheet.Shapes("q2").Line.Visible = msoFalse
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("C30").Select
    ActiveSheet.Paste
    ActiveChart.Parent.Name = "q3"
    ActiveSheet.Shapes("q3").Line.Visible = msoFalse
' Combo box for chart types in historic data
    Sheets("Dashboard").Select
    ActiveSheet.DropDowns.Add(726.75, 78, 115.2, 14.4).Select
    Selection.Name = "q4"
    Selection.Cut
    Range("F4").Select
    ActiveSheet.Paste
    ActiveSheet.Shapes.Range(Array("q4")).Select
    With Selection
        .ListFillRange = "'Historic Data'!$K$5:$K$150"
        .LinkedCell = "'Historic Data'!$Q$25"
        .DropDownLines = 25
        .Display3DShading = True
    End With
    'column chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.Parent.Cut
    Range("G8").Select
    ActiveSheet.Paste
'chart name
    ActiveChart.Parent.Name = "fy"
'data source
    ActiveChart.SetSourceData Source:=Range("'Historic Data'!$P$28:$Q$32")
'data label
    ActiveSheet.Shapes("fy").Line.Visible = msoTrue
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.00,, "
    Range("R6").Select
End If
End Sub
Sub PrepareSummary()
'
' populates array formula for top hits & losses
'
    Sheets("Historic Data").Visible = True
    Sheets("Historic Data").Select
    Range("J5").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX(C[-9],SMALL(IF(C[-7]=""IHCM"",ROW(C[-9])),ROW(R[-4])),0),"""")"
    Selection.AutoFill Destination:=Range("J5:J117"), Type:=xlFillDefault
    Range("J5:J117").Select
    Range("K5").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX(C[-10],SMALL(IF(C[-8]=""NON-IHCM"",ROW(C[-10])),ROW(R[-4])),0),"""")"
    Selection.AutoFill Destination:=Range("K5:K117"), Type:=xlFillDefault
    Range("K5:K117").Select
    Sheets("Historic Data").Visible = False
    
    Sheets("Summary").Select
    Range("A90").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX('Revenue Forecast Final'!C[1],SMALL(IF('Revenue Forecast Final'!C[38]=""Top hits"",ROW('Revenue Forecast Final'!C[1])),ROW(R[-89])),0),"""")"
    Selection.AutoFill Destination:=Range("A90:A119"), Type:=xlFillDefault
    
    Range("A122").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX('Revenue Forecast Final'!C[1],SMALL(IF('Revenue Forecast Final'!C[38]=""Top misses"",ROW('Revenue Forecast Final'!C[1])),ROW(R[-121])),0),"""")"
    Selection.AutoFill Destination:=Range("A122:A151"), Type:=xlFillDefault
    
    Range("A243").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX('Revenue Forecast Final'!C[1],SMALL(IF('Revenue Forecast Final'!C[64]=""Top hits_IHCM"",ROW('Revenue Forecast Final'!C[1])),ROW(R[-242])),0),"""")"
    Selection.AutoFill Destination:=Range("A243:A272"), Type:=xlFillDefault
    Range("A243:A272").Select
    
    Range("A275").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX('Revenue Forecast Final'!C[1],SMALL(IF('Revenue Forecast Final'!C[64]=""Top misses_IHCM"",ROW('Revenue Forecast Final'!C[1])),ROW(R[-274])),0),"""")"
    Selection.AutoFill Destination:=Range("A275:A304"), Type:=xlFillDefault
    Range("A275:A304").Select
    
    Range("A397").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX('Revenue Forecast Final'!C[1],SMALL(IF('Revenue Forecast Final'!C[89]=""Top hits_NON-IHCM"",ROW('Revenue Forecast Final'!C[1])),ROW(R[-396])),0),"""")"
    Selection.AutoFill Destination:=Range("A397:A426"), Type:=xlFillDefault
    Range("A397:A426").Select
    
    Range("A429").Select
    Selection.FormulaArray = _
        "=IFERROR(INDEX('Revenue Forecast Final'!C[1],SMALL(IF('Revenue Forecast Final'!C[89]=""Top misses_NON-IHCM"",ROW('Revenue Forecast Final'!C[1])),ROW(R[-428])),0),"""")"
    Selection.AutoFill Destination:=Range("A429:A458"), Type:=xlFillDefault
    Range("A429:A458").Select
End Sub


