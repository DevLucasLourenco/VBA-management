Sub solicitante_att()


    Dim answer As Integer
    answer = MsgBox("Você quer executar a tarefa em SOLICITANTE?", vbYesNo)
    
    If answer = vbYes Then

        Columns("H:H").Select
        Selection.End(xlDown).Select
        Selection.End(xlDown).Select
        Selection.End(xlDown).Select
        Selection.End(xlDown).Select
        Selection.End(xlDown).Select
        Selection.End(xlDown).Select
        Selection.End(xlDown).Select
        Selection.End(xlUp).Select
        Range(Selection, Selection.End(xlUp)).Select
        Selection.FillDown
        Selection.End(xlUp).Select
        Selection.End(xlDown).Select
        
    Else
    
        MsgBox "Você selecionou NÃO!"
    End If
    
End Sub

