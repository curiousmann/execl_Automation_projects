Attribute VB_Name = "Module7"
Sub load_dashboard()
Attribute load_dashboard.VB_ProcData.VB_Invoke_Func = " \n14"
'
' load_dashboard Macro
' use this 'only' if the butten is deleted

'
    Sheets("East").Select
    ActiveWindow.SmallScroll Down:=40
    ActiveSheet.Shapes.Range(Array("data")).Select
    Selection.Copy
    Selection.Copy
    Sheets("data").Select
    Range("I18").Select
    ActiveSheet.Paste
    Selection.ShapeRange.IncrementLeft 20.25
    Selection.ShapeRange.IncrementTop -9
    Sheets("South").Select
    ActiveWindow.SmallScroll Down:=40
    Sheets("West").Select
    ActiveWindow.SmallScroll Down:=30
    ActiveSheet.Shapes.Range(Array("Slicers_butten")).Select
    Selection.Copy
    Sheets("data").Select
    Range("K17").Select
    ActiveSheet.Paste
    Selection.ShapeRange.IncrementLeft 2.25
    Selection.ShapeRange.IncrementTop 6
    Range("S27").Select
    Sheets("East").Select
    ActiveWindow.SmallScroll Down:=-65
    Range("A1").Select
    Sheets("West").Select
    ActiveWindow.SmallScroll Down:=-70
    Range("A1").Select
    Sheets("South").Select
    ActiveWindow.SmallScroll Down:=-110
    Sheets("data").Select
    Range("S26").Select
End Sub
