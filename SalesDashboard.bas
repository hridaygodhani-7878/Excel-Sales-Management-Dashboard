Option Explicit

Sub SlicerConnection()

If Sheet1.Range("A1").Value = True Then

    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable3"))
Else
              
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable3"))
End If

If Sheet1.Range("D1").Value = True Then

    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable4"))
Else
              
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable4"))
End If

If Sheet1.Range("G1").Value = True Then

    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable5"))
Else
              
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable5"))
End If

If Sheet1.Range("J1").Value = True Then

    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable6"))
Else
              
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable6"))
End If
End Sub

