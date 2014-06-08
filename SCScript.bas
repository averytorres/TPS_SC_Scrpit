Attribute VB_Name = "Module3"
Sub MakeSCPivot()

Dim pt As PivotTable
Dim cacheOfpt As PivotCache  'source data for pt
Dim pf As PivotField
Dim pi As PivotItem
Dim WSD As Worksheet
Dim PRange As Range
Dim wsSheet1 As Worksheet

'try to rename the workbook with data into 'Data', if it already exists use it
On Error Resume Next
Set wsSheet1 = Sheets("Data")
On Error GoTo 0
If Not wsSheet1 Is Nothing Then
	Sheets("Data").Select ' does exist
Else
	ActiveSheet.Name = "Data" 'does not exist
End If

'get rage of used cells in data workbook
Set WSD = Worksheets("Data")
FinalRow = WSD.Cells(Application.Rows.Count, 1).End(xlUp).Row
FinalCol = WSD.Cells(1, Application.Columns.Count). _
    End(xlToLeft).Column
Set PRange = WSD.Cells(1, 1).Resize(FinalRow, FinalCol)

'try to create a workbok named 'Result' if it already exists use it
Dim wsSheet As Worksheet
On Error Resume Next
Set wsSheet = Sheets("Result")
On Error GoTo 0
If Not wsSheet Is Nothing Then
	Sheets("Result").Select ' does exist
Else
	Sheets.Add.Name = "Result" 'does not exist
End If

'delete old pivot table if it exists
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
	.PivotFields(" Status").Orientation = xlColumnField
	.PivotFields(" Count").Orientation = xlDataField
End With

'set correct nodes to be displayed
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
Sheets("Result").Range("A3").Value = "Total Transactions"
Sheets("Result").Range("A4").Value = "Itervals"
Sheets("Result").Range("B4").Value = "Success"
Sheets("Result").Range("C4").Value = "Error"
Sheets("Result").Range("D4").Value = "Total Transactions"
Sheets("Result").Select

'get pivot table size
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
