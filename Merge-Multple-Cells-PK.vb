'Macro que permite fusionar varias celdas de diferentes valores, respetando estas 
'diferencias. 

Option Explicit

Sub MergeSameCells()

Application.DisplayAlerts = False

Dim rng As Range
MergeCells:

For Each rng In Selection
    If rng.Value = rng.Offset(1, 0).Value And rng.Value <> "" Then
        Range(rng, rng.Offset(1, 0)).Merge
        GoTo MergeCells
        Range(rng, rng.Offset(1, 0)).Merge
        GoTo MergeCells
    End If

Next
End Sub