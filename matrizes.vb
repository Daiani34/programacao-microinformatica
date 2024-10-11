Sub exemplo1()
    'Exemplo de vetor dinâmico
    Dim vetor() As Integer
    Dim tam As Integer, i As Integer
    tam = InputBox("Qual é o tamanho")
    ReDim vetor(tam - 1)
    For i = 0 To tam - 1
        vetor(i) = InputBox("Digite um valor")
    Next
    'somar todos os elemento
    Dim soma As Integer
    For i = 0 To tam - 1
        soma = soma + vetor(i)
    Next
    MsgBox "A soma de todos os elementos é: " & soma
End Sub
 
Sub exemplo2()
    Dim casais(2, 1) As String
    Dim i As Integer, j As Integer
    For i = 0 To 2 'total de linhas 3
        For j = 0 To 1 ' total de colunas 2
            casais(i, j) = InputBox("Digite o nome do parceiro " & (j + 1))
        Next
    Next
    'Apresento os casais para dança
    Dim nomes As String
    For i = 0 To 2
        For j = 0 To 1
            nomes = nomes & casais(i, j) & "   "
        Next
        MsgBox "Casal: " & nomes
        nomes = ""
    Next
End Sub
 
Sub exemplo3()
    'calcular a media ponderada de 03 notas de 04 alunos e exibir resultados
    Dim alunos(3) As String, resultados(3) As String
    Dim notas(3, 3) As Single
    Dim i As Integer, j As Integer
    Dim media As Single
    'Receber os dados e calcular a media
    For i = 0 To 3
        alunos(i) = InputBox("Digite o nome do aluno")
        For j = 0 To 2
            notas(i, j) = InputBox("Digite a nota " & (j + 1))
        Next
        notas(i, 3) = (notas(i, 0) * 2 + notas(i, 1) * 2 + notas(i, 2) * 3) / 7
        resultados(i) = IIf(notas(i, 3) >= 6, "Aprovado", "Reprovado")
    Next
    'Gerar o relatorio no debug.print
    Dim mensagem As String
    For i = 0 To 3
        mensagem = ""
        mensagem = alunos(i) & " foi " & resultados(i) & " sua media " & _
            Format(notas(i, 3), "#.#0")
        Debug.Print mensagem
    Next
End Sub
 
Sub exemplo4()
    Dim nota As Single
    nota = 7.8
    MsgBox Format(nota, "#.#0")
End Sub