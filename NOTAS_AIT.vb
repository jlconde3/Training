Sub NOTAS_AIT()

Dim i As Integer, j As Integer, LastStudent As Integer, billet_number_PEAIT, excercise_PEAIT, billet_number_AIT, excercise_AIT As Range



LastStudent = Worksheets("PLAN_ESTUDIO_AIT").Range("A" & Rows.Count).End(xlUp).Row

For i = 2 To LastStudent
    For j = 3 To 8
        billet_number_PEAIT = Worksheets("PLAN_ESTUDIO_AIT").Cells(i, 1)
        excercise_PEAIT = Worksheets("PLAN_ESTUDIO_AIT").Cells(i, j)
        If Not excercise_PEAIT = "REVIEW" Then
            If Not excercise_PEAIT = "NAVANTIA SURVEY" Then
                If Not excercise_PEAIT = "GHENOVA SURVEY" Then
                    With Worksheets("NOTAS_AIT").Range("A1:A444")
                    Set billet_number_AIT = .Find(billet_number_PEAIT)
                        row_AIT = Range(billet_number_AIT.Address).Row
                    End With
                    With Worksheets("NOTAS_AIT").Range("A1:CF1")
                    Set excercise_AIT = .Find(excercise_PEAIT)
                    colum_AIT = Range(excercise_AIT.Address).Column
                    End With
                    If IsEmpty(Worksheets("NOTAS_AIT").Cells(row_AIT, colum_AIT)) = True Then
                        Worksheets("NOTAS_AIT").Cells(row_AIT, colum_AIT) = Worksheets("PLAN_ESTUDIO_AIT").Cells(i, j + 6)
                    End If
                End If
            End If
        End If
    Next j
Next i

End Sub