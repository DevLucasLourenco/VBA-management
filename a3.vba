Sub preencher_data()

    
     Dim answer As Integer
    answer = MsgBox("Você quer executar a tarefa em PREENCHER DATA?", vbYesNo)
    
    If answer = vbYes Then

        Dim selCell As Range
        Set selCell = Selection
        
        Dim copyValue As Variant
        copyValue = selCell.Value
        
        Dim copyCount As Variant
        copyCount = InputBox("Insira a quantidade de células para copiar e colar:")
        
        If Not IsNumeric(copyCount) Or copyCount < 1 Then
            MsgBox "Quantidade invï¿½lida! Insira um valor numï¿½rico maior que zero."
            Exit Sub
        End If
        
        Dim pasteRange As Range
        Set pasteRange = Range(selCell.Offset(1, 0), selCell.Offset(CInt(copyCount), 0))
        pasteRange.Value = copyValue
        
    Else
    
        MsgBox "Você selecionou NÃO!"
    End If
    
End Sub


