Attribute VB_Name = "Module1"
Sub MakeTPSPivot()

Dim pt As PivotTable
Dim cacheOfpt As PivotCache  'source data for pt
Dim pf As PivotField
Dim pi As PivotItem
Dim WSD As Worksheet
Dim PRange As Range
Dim wsSheet1 As Worksheet

'try to rename current workbook 'Data' if it is not already
On Error Resume Next
Set wsSheet1 = Sheets("Data")
On Error GoTo 0
If Not wsSheet1 Is Nothing Then
	Sheets("Data").Select ' does exist	
Else
	ActiveSheet.Name = "Data" 'does not exist
End If

'get data rage size for later
Set WSD = Worksheets("Data")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

'try to create workbook named 'Result' if it does not already exist
Dim wsSheet As Worksheet
On Error Resume Next
Set wsSheet = Sheets("Result")
On Error GoTo 0
If Not wsSheet Is Nothing Then
	Sheets("Result").Select ' does exist
Else
	Sheets.Add.Name = "Result" 'does not exist
End If

'try to delete previous pivot table if one exists
On Error Resume Next
ActiveSheet.PivotTables("MyPT").TableRange2.Clear 'delete old pt

'set cache of pt
Sheets("Data").Select
Set cacheOfpt = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=PRange)

'creat pt
Sheets("Result").Select
Set pt = ActiveSheet.PivotTables.Add(cacheOfpt, Range("a1"), "MyPT")

'put fields in

With pt
	'add fields
	.PivotFields(" Time").Orientation = xlRowField
	.PivotFields(" Node").Orientation = xlPageField
	.PivotFields(" Avg TPS").Orientation = xlDataField
	.PivotFields(" Max TPS").Orientation = xlDataField
	.PivotFields(" Min TPS").Orientation = xlDataField
End With

'display only nodes that are needed
Set pf = pt.PivotFields(" Node")
With pf
	For Each pi In pf.PivotItems
		If pi.Name Like " AES_*" Then
		pi.Visible = True
		Else
		pi.Visible = False
		End If
	Next pi
End With

'rename columns
Sheets("Result").Range("A4").Value = "Intervals"
Sheets("Result").Range("B4").Value = "Average"
Sheets("Result").Range("C4").Value = "Max TPS"
Sheets("Result").Range("D4").Value = "Min TPS"

'set max, average and min settings    
With ActiveSheet.PivotTables("MyPT").PivotFields("Average")
	.Caption = "Average TPS"
    .Function = xlAverage
End With
With ActiveSheet.PivotTables("MyPT").PivotFields("Max TPS")
    .Caption = "Max TPS"
    .Function = xlMax
End With
With ActiveSheet.PivotTables("MyPT").PivotFields("Min TPS")
    .Caption = "Min TPS"
    .Function = xlMin
End With

'format number columns
Columns(2).NumberFormat = "0.00"
Columns(3).NumberFormat = "0.00"

'get range data of pivot table
Set WSD = Worksheets("Result")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

'create line chart
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=PRange
ActiveChart.ChartType = xlLineMarkers
With ActiveChart.Parent
         .Height = 350 ' resize
         .Width = 500  ' resize
         .Top = 10    ' reposition
         .Left = 250   ' reposition
End With

End Sub


