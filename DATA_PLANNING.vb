Sub DATA_PLANNING()
Dim i As Long, LastRowIput As Long, RowInput As Integer, RowPlanningFAM As Integer, LastRowPlanning As Integer, RowPlanningMP As Integer
RowInput = 2
RowPlanningFAM = 10
RowPlanningMP = 10

Application.ScreenUpdating = False
    
''Crear una copia de la hoja TEMPLATE y cambiarle el nombre a PLANNING
Worksheets("TEMPLATE").Copy After:=Worksheets("TEMPLATE")
ActiveSheet.name = "PLANNING_FAM"

Worksheets("TEMPLATE").Copy After:=Worksheets("TEMPLATE")
ActiveSheet.name = "PLANNING_MP"

''Copiar los datos de la hoja INPUT a la hoja PLANNING
With Worksheets("INPUT")
    LastRowInput = .Range("A1").End(xlDown).row
End With

Do While RowInput < LastRowInput
    If Worksheets("INPUT").Cells(RowInput, "L").Value = "FAM_THE" Or Worksheets("INPUT").Cells(RowInput, "L").Value = "FAM_AIT" Then
        Call COPY_DATA(RowInput, RowPlanningFAM, "PLANNING_FAM")
        RowPlanningFAM = RowPlanningFAM + 1
    Else:
        Call COPY_DATA(RowInput, RowPlanningMP, "PLANNING_MP")
        RowPlanningMP = RowPlanningMP + 1
    End If
    
    RowInput = RowInput + 1
Loop
    
Application.ScreenUpdating = True


End Sub

''Macro para copiar los valores de la hoja Input a la hoja Planning
Sub COPY_DATA(RowInput As Integer, RowPlannnig As Integer, Sheet As String)

Dim A As Variant, B As Variant, i As Integer

A = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
B = Array(1, 2, 3, 4, 5, 7, 8, 9, 12)

For i = 0 To 8
    Worksheets(Sheet).Cells(RowPlannnig, A(i)) = Worksheets("INPUT").Cells(RowInput, B(i))
Next i



End Sub