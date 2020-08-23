Attribute VB_Name = "Module6"
Sub load_data()
'
'
    Sheets("Data").Select
    Columns("A:H").Select
    Selection.ClearContents
    Range("A15").Select
    Sheets("East").Select
    Range("A1:A2").Select
    Selection.Copy
    Sheets("Data").Select
    Range("A15").Select
    ActiveSheet.Paste
    Range("B15").Select
    Sheets("East").Select
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Sheets("Data").Select
    Range("B15").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Sheets("West").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("A1").Select
    Sheets("Data").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Sheets("South").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("A1").Select
    Sheets("Data").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Sheets("North").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("A1").Select
    Sheets("Data").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.FormulaR1C1 = "2"
    Range("A16").Select
    Selection.AutoFill Destination:=Range(Selection, Selection.End(xlDown)), Type:=xlFillSeries
    Range("A15").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Selection, Selection.End(xlToRight)), , xlYes).Name = _
        "My table"
    ActiveSheet.Shapes.Range(Array("data")).Select
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    Dim ch As Shape
    Range("A1").Select
    Range("My_table[[#Headers],[Qty]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("data!$F$15:$H$62")
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(3).ChartType = xlLine
    ActiveChart.FullSeriesCollection(3).AxisGroup = 1
    ActiveChart.FullSeriesCollection(1).AxisGroup = 2
    ActiveChart.FullSeriesCollection(3).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).ChartType = xlLine
    ActiveChart.FullSeriesCollection(2).ChartType = xlAreaStacked
    Set ch = ActiveSheet.Shapes(1)
    ch.Name = "My chart"
    ActiveSheet.Shapes("My chart").IncrementLeft -314.25
    ActiveSheet.Shapes("My chart").IncrementTop -97.1249606299
    ActiveSheet.Shapes("My chart").ScaleHeight 0.96875, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("My chart").ScaleWidth 1.7895833333, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.ChartObjects("My chart").Activate
    ActiveSheet.ChartObjects("My chart").Activate
    ActiveChart.ChartTitle.Select
    Selection.Delete
    Application.Run "'Dashboard-2 Project_stable_3.6.xlsm'!load_dashboard"
    Range("S26").Select
End Sub
Sub Slicers()
Attribute Slicers.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
    Application.Run "'Dashboard-2 Project_stable_3.6.xlsm'!load_data"
    Range("My_table[#Headers]").Select
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("My table"), "Date"). _
        Slicers.Add ActiveSheet, , "Date", "Date", 33.75, 357, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("My table"), "Product" _
        ).Slicers.Add ActiveSheet, , "Product", "Product", 71.25, 394.5, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("My table"), "Name"). _
        Slicers.Add ActiveSheet, , "Name", "Name", 108.75, 432, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("My table"), "Area"). _
        Slicers.Add ActiveSheet, , "Area", "Area", 146.25, 469.5, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("My table"), "Qty"). _
        Slicers.Add ActiveSheet, , "Qty", "Qty", 183.75, 507, 144, 198.75
    ActiveSheet.Shapes.Range(Array("Qty")).Select
    ActiveSheet.Shapes("Qty").IncrementLeft 336.75
    ActiveSheet.Shapes("Qty").IncrementTop -181.5
    ActiveSheet.Shapes("Qty").ScaleWidth 0.578125, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes.Range(Array("Date")).Select
    ActiveSheet.Shapes("Date").IncrementLeft 401.25
    ActiveSheet.Shapes("Date").IncrementTop -32.25
    ActiveSheet.Shapes("Date").ScaleWidth 1.7708333333, msoFalse, _
        msoScaleFromBottomRight
    ActiveWorkbook.SlicerCaches("Slicer_Date").Slicers("Date").NumberOfColumns = 2
    ActiveWorkbook.SlicerCaches("Slicer_Date").Slicers("Date").NumberOfColumns = 3
    ActiveWorkbook.SlicerCaches("Slicer_Date").Slicers("Date").NumberOfColumns = 4
    ActiveWorkbook.SlicerCaches("Slicer_Date").Slicers("Date").NumberOfColumns = 5
    ActiveWorkbook.SlicerCaches("Slicer_Date").Slicers("Date").NumberOfColumns = 4
    ActiveWorkbook.SlicerCaches("Slicer_Date").Slicers("Date").NumberOfColumns = 3
    ActiveSheet.Shapes("Date").ScaleHeight 0.3622641509, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Date").IncrementLeft -11.25
    ActiveSheet.Shapes("Date").IncrementTop 122.25
    ActiveSheet.Shapes.Range(Array("Area")).Select
    ActiveSheet.Shapes("Area").IncrementLeft 174.75
    ActiveSheet.Shapes("Area").IncrementTop -137.25
    ActiveSheet.Shapes("Area").ScaleWidth 1.7395833333, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Area").IncrementLeft 7.5
    ActiveSheet.Shapes("Area").IncrementTop -5.25
    ActiveSheet.Shapes("Area").ScaleWidth 1.0239520958, msoFalse, _
        msoScaleFromBottomRight
    ActiveWorkbook.SlicerCaches("Slicer_Area").Slicers("Area").NumberOfColumns = 2
    ActiveWorkbook.SlicerCaches("Slicer_Area").Slicers("Area").NumberOfColumns = 3
    ActiveWorkbook.SlicerCaches("Slicer_Area").Slicers("Area").NumberOfColumns = 4
    ActiveWorkbook.SlicerCaches("Slicer_Area").Slicers("Area").Style = _
        "SlicerStyleLight6"
    ActiveWorkbook.SlicerCaches("Slicer_Area").Slicers("Area").Style = _
        "SlicerStyleDark6"
    ActiveSheet.Shapes("Area").ScaleHeight 0.2339622642, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Qty")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Qty").Slicers("Qty").Style = _
        "SlicerStyleDark5"
    ActiveSheet.Shapes.Range(Array("Name")).Select
    ActiveSheet.Shapes("Name").IncrementLeft 213.75
    ActiveSheet.Shapes("Name").IncrementTop -54
    ActiveSheet.Shapes("Name").ScaleWidth 1.7760416667, msoFalse, _
        msoScaleFromTopLeft
    ActiveWorkbook.SlicerCaches("Slicer_Name").Slicers("Name").NumberOfColumns = 2
    ActiveWorkbook.SlicerCaches("Slicer_Name").Slicers("Name").NumberOfColumns = 3
    ActiveWorkbook.SlicerCaches("Slicer_Name").Slicers("Name").Style = _
        "SlicerStyleDark4"
    ActiveSheet.Shapes("Name").ScaleHeight 0.3283018868, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Name").ScaleHeight 1.091954023, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Date")).Select
    ActiveSheet.Shapes("Date").IncrementLeft 9.75
    ActiveSheet.Shapes("Date").IncrementTop 3.75
    ActiveSheet.Shapes("Date").IncrementLeft -7.5
    ActiveSheet.Shapes("Date").IncrementTop 102.75
    ActiveSheet.Shapes.Range(Array("Product")).Select
    ActiveSheet.Shapes("Product").IncrementLeft 249.75
    ActiveSheet.Shapes("Product").IncrementTop 60
    ActiveSheet.Shapes("Product").ScaleWidth 1.7760416667, msoFalse, _
        msoScaleFromTopLeft
    ActiveWorkbook.SlicerCaches("Slicer_Product").Slicers("Product"). _
        NumberOfColumns = 2
    ActiveWorkbook.SlicerCaches("Slicer_Product").Slicers("Product"). _
        NumberOfColumns = 3
    ActiveWorkbook.SlicerCaches("Slicer_Product").Slicers("Product"). _
        NumberOfColumns = 4
    ActiveWorkbook.SlicerCaches("Slicer_Product").Slicers("Product"). _
        NumberOfColumns = 3
    ActiveSheet.Shapes("Product").ScaleHeight 0.3433962264, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Product").ScaleHeight 1.1318681319, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Product").ScaleHeight 0.9417475728, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Date")).Select
    ActiveSheet.Shapes("Date").IncrementLeft 6
    ActiveSheet.Shapes("Date").IncrementTop -22.5
    ActiveSheet.Shapes.Range(Array("Qty")).Select
    ActiveSheet.Shapes("Qty").ScaleHeight 1.3924528302, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Product")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Product").Slicers("Product").Style = _
        "SlicerStyleDark2"
    ActiveSheet.Shapes.Range(Array("Date")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Date").Slicers("Date").Style = _
        "SlicerStyleDark3"
    Range("N15").Select
End Sub
