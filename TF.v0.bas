Attribute VB_Name = "Module1"
Sub Main()
Call ValorGastoxCliqueLink
Call Envolvimento
Call VisualizacaoConteudoxVisualizacaoConteudoUnica
Range("J1").Value = "Custo / Clique"
Range("K1").Value = "Envolvimento"
Range("L1").Value = "VisuConteu / VisuConteuUnic"
End Sub

Sub ValorGastoxCliqueLink()
Dim i As Integer
Dim a As Double
Dim b As Double
Dim c As Double
Dim maior As Double
i = 1
varColuna = 1 ' Coluna que será verificado
varlinha = 1 ' Linha inicial que será verificado
varConteudo = i
Cells(1, 1).Select
Do While varConteudo <> Empty 'continua a verificar se conteudo for diferente de vazio
varlinha = varlinha + 1 'contador de linha
varConteudo = Cells(varlinha, varColuna).Value 'grava o valor da celula
i = i + 1
a = Cells(i, 4).Value
b = Cells(i, 3).Value
If b <> 0 Then
   c = a / b
ActiveCell.Offset(i - 1, 10).Value = c
End If
If i = 2 Then
    maior = c
End If
If c > maior Then
    maior = c
End If
numlinha = varlinha - 1
Loop
varColuna = 1 ' Coluna que será verificado
varlinha = 1 ' Linha inicial que será verificado
varConteudo = 1
Cells(1, 1).Select
Do While varConteudo <> Empty 'continua a verificar se conteudo for diferente de vazio
varlinha = varlinha + 1 'contador de linha
varConteudo = Cells(varlinha, varColuna).Value 'grava o valor da celula
c = Cells(varlinha, 11).Value
Cells(varlinha, 11).ClearContents
If numlinha <> varlinha - 1 Then
ActiveCell.Offset(varlinha - 1, 9).Value = c / maior
End If
Loop
End Sub

Sub Envolvimento()
Dim i As Integer
Dim a As Double
Dim maior As Double
i = 1
varColuna = 1 ' Coluna que será verificado
varlinha = 1 ' Linha inicial que será verificado
varConteudo = i
Cells(1, 1).Select
Do While varConteudo <> Empty 'continua a verificar se conteudo for diferente de vazio
varlinha = varlinha + 1 'contador de linha
varConteudo = Cells(varlinha + 1, varColuna).Value 'grava o valor da celula
i = i + 1
a = Cells(i, 5).Value
ActiveCell.Offset(i - 1, 11).Value = a
If i = 2 Then
    maior = a
ElseIf a > maior Then
    maior = a
End If
numlinha = varlinha - 1
Loop
varColuna = 1 ' Coluna que será verificado
varlinha = 1 ' Linha inicial que será verificado
varConteudo = 1
Cells(1, 1).Select
Do While varConteudo <> Empty 'continua a verificar se conteudo for diferente de vazio
varlinha = varlinha + 1 'contador de linha
varConteudo = Cells(varlinha + 1, varColuna).Value 'grava o valor da celula
a = Cells(varlinha, 12).Value
Cells(varlinha, 12).ClearContents
If numlinha > varlinha - 2 Then
ActiveCell.Offset(varlinha - 1, 10).Value = a / maior
End If
Loop
End Sub

Sub VisualizacaoConteudoxVisualizacaoConteudoUnica()
Dim i As Integer
Dim a As Double
Dim b As Double
Dim c As Double
Dim maior As Double
i = 1
varColuna = 1 ' Coluna que será verificado
varlinha = 1 ' Linha inicial que será verificado
varConteudo = i
Cells(1, 1).Select
Do While varConteudo <> Empty 'continua a verificar se conteudo for diferente de vazio
varlinha = varlinha + 1 'contador de linha
varConteudo = Cells(varlinha, varColuna).Value 'grava o valor da celula
i = i + 1
a = Cells(i, 6).Value
b = Cells(i, 7).Value
If b <> 0 Then
   c = a / b
ActiveCell.Offset(i - 1, 12).Value = c
End If
If i = 2 Then
    maior = c
End If
If c > maior Then
    maior = c
End If
numlinha = varlinha - 1
Loop
varColuna = 1 ' Coluna que será verificado
varlinha = 1 ' Linha inicial que será verificado
varConteudo = 1
Cells(1, 1).Select
Do While varConteudo <> Empty 'continua a verificar se conteudo for diferente de vazio
varlinha = varlinha + 1 'contador de linha
varConteudo = Cells(varlinha, varColuna).Value 'grava o valor da celula
c = Cells(varlinha, 13).Value
Cells(varlinha, 13).ClearContents
If numlinha <> varlinha - 1 Then
ActiveCell.Offset(varlinha - 1, 11).Value = c / maior
End If
Loop
End Sub

