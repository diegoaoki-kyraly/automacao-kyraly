Attribute VB_Name = "Module1"
Sub Main()
Call VisualizacaoPaginaxCliqueLink
Call VisualizacaoConteudoxVisualizacaoPagina
Call CustoxVisualizacaoConteudo
Range("K1").Value = "VisuPag / Clique"
Range("L1").Value = "VisuPag / VisuConteu"
Range("M1").Value = "Custo / VisuConteu"
End Sub

Sub VisualizacaoPaginaxCliqueLink()
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
a = Cells(i, 5).Value
b = Cells(i, 4).Value
If b <> 0 Then
   c = a / b
ActiveCell.Offset(i - 1, 11).Value = c
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
a = Cells(varlinha, 12).Value
Cells(varlinha, 12).ClearContents
If numlinha <> varlinha - 1 Then
ActiveCell.Offset(varlinha - 1, 10).Value = a / maior
End If
Loop
End Sub

Sub VisualizacaoConteudoxVisualizacaoPagina()
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
a = Cells(i, 5).Value
b = Cells(i, 6).Value
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
a = Cells(varlinha, 13).Value
Cells(varlinha, 13).ClearContents
If numlinha <> varlinha - 1 Then
ActiveCell.Offset(varlinha - 1, 11).Value = a / maior
End If
Loop
End Sub

Sub CustoxVisualizacaoConteudo()
Dim i As Integer
Dim a As Double
Dim b As Double
Dim c As Double
Dim maior As Double
i = 1
varColuna = 1 ' Coluna que será verificado
varlinha = 1 ' Linha inicial que será verificado
varConteudo = 1
Cells(1, 1).Select
Do While varConteudo <> Empty 'continua a verificar se conteudo for diferente de vazio
varlinha = varlinha + 1 'contador de linha
varConteudo = Cells(varlinha, varColuna).Value 'grava o valor da celula
i = i + 1
a = Cells(i, 8).Value
b = Cells(i, 6).Value
If b <> 0 Then
   c = a / b
   ActiveCell.Offset(i - 1, 13).Value = c
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
varConteudo = i
Cells(1, 1).Select
Do While varConteudo <> Empty 'continua a verificar se conteudo for diferente de vazio
varlinha = varlinha + 1 'contador de linha
varConteudo = Cells(varlinha, varColuna).Value 'grava o valor da celula
a = Cells(varlinha, 14).Value
Cells(varlinha, 14).ClearContents
If numlinha <> varlinha - 1 Then
ActiveCell.Offset(varlinha - 1, 12).Value = a / maior
End If
Loop
End Sub

