Sub hr_inicial()


    Dim answer As Integer
    answer = MsgBox("Você quer executar a tarefa em HORA INICIAL?", vbYesNo)
    
    If answer = vbYes Then
        Columns("F:F").Select
        Selection.End(xlDown).Select
        Selection.End(xlDown).Select
        Selection.End(xlDown).Select
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
        
        MsgBox "Você Selecionou NÃO!"
    End If
    
End Sub
Sub hr_final()

     Dim answer As Integer
    answer = MsgBox("Você quer executar a tarefa em HORA FINAL?", vbYesNo)
    
    If answer = vbYes Then
    
        Columns("G:G").Select
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
    
        MsgBox "Você Selecionou NÃO!"
    End If
    
End Sub

