# Excel_VBA
#Soluções em VBA planilha Método GTD


Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("0-Coisas")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row ' Encontrar a última linha preenchida na coluna D
    
    Dim changedCell As Range
    For Each changedCell In Target
        If Not Intersect(changedCell, ws.Range("D1:D" & lastRow)) Is Nothing Then
            ' Se a célula em D for alterada para "Não", preenche E e F com "-"
            If changedCell.Value = "Não" Then
                ws.Cells(changedCell.Row, 5).Value = "-"
                ws.Cells(changedCell.Row, 6).Value = "-"
                MsgBox "Colocar na espera?"
                ws.Cells(changedCell.Row, 8).Select
            End If
        ElseIf Not Intersect(changedCell, ws.Range("E1:E" & lastRow)) Is Nothing Then
            ' Se a célula em E for alterada para "Sim", preenche J, L, N com "-"
            If changedCell.Value = "Sim" Then
                ws.Cells(changedCell.Row, 10).Value = "-"
                ws.Cells(changedCell.Row, 12).Value = "-"
                ws.Cells(changedCell.Row, 14).Value = "-"
            End If
        ElseIf Not Intersect(changedCell, ws.Range("F1:F" & lastRow)) Is Nothing Then
            ' Se a célula em F for alterada para "Sim", mostra MsgBox e posiciona na célula em G
            If changedCell.Value = "Sim" Then
                MsgBox "Qual o próximo passo?"
                ws.Cells(changedCell.Row, 7).Select
                ws.Cells(changedCell.Row, 8).Value = "-"
            End If
        ElseIf Not Intersect(changedCell, ws.Range("H1:H" & lastRow)) Is Nothing Then
            ' Se a célula em G for alterada para "Incubar", "Eliminar" ou "Arquivar"
            Select Case changedCell.Value
                Case "Incubar"
                    ws.Cells(changedCell.Row, 12).Value = "-"
                    ws.Cells(changedCell.Row, 14).Value = "-"
                    ws.Cells(changedCell.Row, 10).Value = "Sim"
                Case "Eliminar"
                    ws.Cells(changedCell.Row, 10).Value = "-"
                    ws.Cells(changedCell.Row, 14).Value = "-"
                    ws.Cells(changedCell.Row, 12).Value = "Sim"
                Case "Arquivar"
                    ws.Cells(changedCell.Row, 10).Value = "-"
                    ws.Cells(changedCell.Row, 12).Value = "-"
                    ws.Cells(changedCell.Row, 14).Value = "Sim"
            End Select
        End If
    Next changedCell
End Sub
