Attribute VB_Name = "Module1"
Sub Main()
Call CPA
Call Vendas
Call VendasxAdicaoCarrinho
Range("I1").Value = "CPA"
Range("J1").Value = "Vendas"
Range("K1").Value = "Venda / AddCarrinho"
End Sub

Sub CPA()
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
        varConteudo = Cells(varlinha, varColuna).Value 'grava o valor da celula
        i = i + 1
        If Cells(i, 4).Value <> Empty Then
            a = Cells(i, 4).Value
            ActiveCell.Offset(i - 1, 9).Value = a
            If i = 2 Then
                maior = a
            End If
            If a > maior Then
                maior = a
            End If
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
        a = Cells(varlinha, 10).Value
        Cells(varlinha, 10).ClearContents
        If numlinha <> varlinha - 1 Then
            ActiveCell.Offset(varlinha - 1, 8).Value = a / maior
        End If
    Loop
End Sub

Sub Vendas()
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
        varConteudo = Cells(varlinha, varColuna).Value 'grava o valor da celula
        i = i + 1
        If Cells(i, 5).Value <> Empty Then
            a = Cells(i, 5).Value
            ActiveCell.Offset(i - 1, 10).Value = a
            If i = 2 Then
                maior = a
            End If
            If a > maior Then
                maior = a
            End If
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
        a = Cells(varlinha, 11).Value
        Cells(varlinha, 11).ClearContents
        If numlinha <> varlinha - 1 Then
            ActiveCell.Offset(varlinha - 1, 9).Value = a / maior
        End If
    Loop
End Sub

Sub VendasxAdicaoCarrinho()
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
        If Cells(i, 6).Value <> Empty And Cells(i, 5).Value <> Empty Then
            a = Cells(i, 5).Value
            b = Cells(i, 6).Value
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
        c = Cells(varlinha, 12).Value
        Cells(varlinha, 12).ClearContents
        If numlinha <> varlinha - 1 Then
            ActiveCell.Offset(varlinha - 1, 10).Value = c / maior
        End If
    Loop
End Sub
