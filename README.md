# **Excel Avançado - VBA HHT**

👋 **Situação Problema**

Você é responsável por gerenciar os treinamentos de uma empresa e precisa controlar as cargas horárias de cada treinamento por setor. O objetivo é criar uma planilha onde você possa cadastrar as informações dos treinamentos e, em seguida, calcular a carga horária total por setor.

1. Abra o arquivo atividade_hht.xlsx

2. Crie uma planilha e a renomeie para **Resultado**

3. Abra o Visual Basic Editor (VBE)

4. Insira um novo módulo e cole o seguinte código:

```
Sub CalcularCargaHorariaTotal()
    Dim wsTreinamentos As Worksheet
    Dim wsResultado As Worksheet
    Dim ultimaLinhaTreinamentos As Long
    Dim i As Long
    Dim setor As String
    Dim cargaHoraria As Double
    Dim participantes As Long
    Dim resultadoDict As Object
    Dim chave As Variant ' Alteração aqui para evitar o erro
    
    ' Configurar as planilhas
    Set wsTreinamentos = ThisWorkbook.Sheets("Treinamentos")
    Set wsResultado = ThisWorkbook.Sheets("Resultado")
    Set resultadoDict = CreateObject("Scripting.Dictionary")
    
    ' Limpar resultados anteriores
    wsResultado.Range("A2:B" & wsResultado.Cells(wsResultado.Rows.Count, 1).End(xlUp).Row).ClearContents
    
    ' Encontrar a última linha da planilha de treinamentos
    ultimaLinhaTreinamentos = wsTreinamentos.Cells(wsTreinamentos.Rows.Count, 1).End(xlUp).Row
    
    ' Loop pelos dados dos treinamentos
    For i = 2 To ultimaLinhaTreinamentos
        setor = wsTreinamentos.Cells(i, 4).Value ' Setor
        cargaHoraria = wsTreinamentos.Cells(i, 2).Value * wsTreinamentos.Cells(i, 3).Value ' Carga Horária * Participantes
        
        ' Acumular carga horária total por setor
        If resultadoDict.Exists(setor) Then
            resultadoDict(setor) = resultadoDict(setor) + cargaHoraria
        Else
            resultadoDict.Add setor, cargaHoraria
        End If
    Next i
    
    ' Preencher a planilha de resultados
    Dim linhaResultado As Long
    linhaResultado = 2 ' Começar na linha 2
    For Each chave In resultadoDict.Keys ' Alteração aqui
        wsResultado.Cells(linhaResultado, 1).Value = chave
        wsResultado.Cells(linhaResultado, 2).Value = resultadoDict(chave)
        linhaResultado = linhaResultado + 1
    Next chave
    
    MsgBox "Cálculo de carga horária total concluído!"
End Sub
```

5. Vincule a macro CalcularCargaHorariaTotal a um botão.
6. Teste a execução da macro com novos registros.
7. Salve o arquivo com a extensão apropriada.
8. Envie o arquivo para o google sala de aula.

