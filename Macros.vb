
Sub Formato_Hoja_Report()
  'Cambia la apariencia de la hoja y los colores de las celdas
  '
    Range("Table1[[#Headers],[Register]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
End Sub

Sub Insertar_Tabla_Dinamica()
'Macro utilizada en la hoja Pivot Table para generar una tabla dinámica.

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Table1[#All]", Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:="Pivot Table!R4C2", TableName:="PivotTable1", _
        DefaultVersion:=xlPivotTableVersion12
    Sheets("Pivot Table").Select
    Cells(4, 2).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Amount"), "Sum of Amount", xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Banc")
        .Orientation = xlColumnField
        .Position = 1
    End With
End Sub
