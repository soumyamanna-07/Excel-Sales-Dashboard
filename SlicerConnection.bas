Attribute VB_Name = "SlicerConnection"
Option Explicit

' ---------------------------------------------------------------
' Connects / disconnects a slicer from pivot tables based on
' checkbox-linked cells on Sheet2.
'
' Linked cells:  A1 -> PivotTable5
'                D1 -> PivotTable6
'                G1 -> PivotTable7
'                J1 -> PivotTable8
'
' Change the "Set ws" line below to the sheet holding the pivots.
' ---------------------------------------------------------------

Sub SlicerConnection()

    Dim ws As Worksheet
    Dim sc As SlicerCache
    Dim linkCells As Variant
    Dim pivotNames As Variant
    Dim i As Long

    Set ws = Sheet1                                   ' sheet with the pivot tables
    Set sc = ActiveWorkbook.SlicerCaches("Slicer_Region")

    linkCells = Array("A1", "D1", "G1", "J1")
    pivotNames = Array("PivotTable5", "PivotTable6", "PivotTable7", "PivotTable8")

    Application.ScreenUpdating = False

    For i = LBound(linkCells) To UBound(linkCells)

        On Error Resume Next

        If Sheet2.Range(linkCells(i)).Value = True Then
            sc.PivotTables.AddPivotTable ws.PivotTables(pivotNames(i))
        Else
            sc.PivotTables.RemovePivotTable ws.PivotTables(pivotNames(i))
        End If

        On Error GoTo 0

    Next i

    Application.ScreenUpdating = True

End Sub


' ---------------------------------------------------------------
' Helper: prints every pivot table and slicer cache name to the
' Immediate Window (Ctrl + G). Run this to find the exact names
' to use above.
' ---------------------------------------------------------------

Sub ListPivotsAndSlicers()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sc As SlicerCache

    Debug.Print "--- PIVOT TABLES ---"
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            Debug.Print ws.Name & " | " & pt.Name & " | " & pt.TableRange1.Address
        Next pt
    Next ws

    Debug.Print "--- SLICER CACHES ---"
    For Each sc In ActiveWorkbook.SlicerCaches
        Debug.Print sc.Name
    Next sc

End Sub
