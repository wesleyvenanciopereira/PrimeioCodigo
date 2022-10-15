Attribute VB_Name = "Módulo1"
Sub primeiro()
'O comando DIM(Dimension) é utilizado para declarar variavel
'A variavel nome foi tipada como String(texto)
Dim nome As String
'O comando InputBox abre uma caixa de entrada de dados
'Assim o usuário digita o nome e aloca na variavel nome
nome = InputBox("Digite o seu nome")
'O comando Range permite selecionar uma célula na planilha do excel. Assim selecionamos a célula A1 e adicionamos o valor que foi digitado na caixa de entrada usando a variável nome
Range("A1").Value = nome
End Sub
