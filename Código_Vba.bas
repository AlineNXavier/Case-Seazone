Attribute VB_Name = "Módulo1"


Sub Create()
'   Essa Macro cria as planilhas

    Sheets.Add.Name = "Calculo"
    Sheets.Add.Name = "Dias"
End Sub

Sub Clear()
'   Nessa macro limpamos os conteúdos das celulas antes da execução

    Sheets("Calculo").Rows.ClearContents
    Sheets("Dias").Rows.ClearContents
End Sub

Sub Delete()
'   Deletando as planilhas criadas

    Sheets("Calculo").Delete
    Sheets("Dias").Delete
End Sub

Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = " \n14"
'
'   Nessa macro iremos preencher as duas planilhas criadas, a planilha Calculo conterá dados no faturamento mensal de cada ID
'   E a planilha Dias conterá a informação de diárias mensais alugadas para cada ID.
'   A criação da planilha de diárias foi feita para que no python possamos calcular a taxa média de diárias mensais.

    'Desabilitando algumas funções para que a macro execute mais rápido
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False


    Dim n_lin, n_col, cont_mes As Integer

    'Chamando a macro que limpa as células
    Call Clear
    
    'Obtendo o número de colunas e linhas da planilha principal
    n_lin = Sheets("Sheet1").UsedRange.Rows.Count
    n_col = Sheets("Sheet1").UsedRange.Columns.Count
    
    'Definindo o número de meses que agruparemos as informações
    meses = 6
    
    'Para cada linha de informação
    For i = 2 To n_lin + 1
    
        cont_mes = 0
        
        'Preenchendo o dado do ID
        Sheets("Calculo").Cells(i, 1) = Sheets("Sheet1").Cells(i, 1)
        Sheets("Dias").Cells(i, 1) = Sheets("Sheet1").Cells(i, 1)
        
        'Iniciando do dia mais recente até o mais antigo
        For j = n_col - 1 To 2 Step -1
            
            'If necessário para realizar a mudança de mês
            'Caso o mes da planilha mude atualizamos o valor da variavel mês
            If mes <> Month(Sheets("Sheet1").Cells(1, j)) Then
            
                cont_mes = cont_mes + 1
                
                mes = Month(Sheets("Sheet1").Cells(1, j))
                
                'Nesse if paramo a execução do for das colunas quando atingimos o número de meses especificado na variável mês.
                If cont_mes = meses + 1 Then
                    Exit For
                End If
                
            End If
                
            'Iremos contabilizar no calculo dos alugueis e das diárias apenas as celulas pintadas de verde
            If Sheets("Sheet1").Cells(i, j).Interior.Color = 65280 Then
            
                Sheets("Calculo").Cells(i, cont_mes + 1) = Sheets("Sheet1").Cells(i, j) + Sheets("Calculo").Cells(i, cont_mes + 1)
                Sheets("Dias").Cells(i, cont_mes + 1) = 1 + Sheets("Dias").Cells(i, cont_mes + 1)

            End If
            
        Next
    Next
    
    ' Deixando a coluna de Ids em negrito em ambas as colunas criadas
    Sheets("Calculo").Activate
    Sheets("Calculo").Range(Cells(1, 1), Cells(n_lin, 1)).Font.Bold = True
    Sheets("Dias").Activate
    Sheets("Dias").Range(Cells(1, 1), Cells(n_lin, 1)).Font.Bold = True
    
    
    ' Chamamos essa macro para preencher o cabeçalho das colunas
    Call Head(n_col, meses)

End Sub

Sub Head(col, meses)
 
    Dim marc_col, mes_aluguel, qtd_mes As Integer
    
    marc_col = 1: mes_aluguel = 0: qtd_mes = 0
    
    
    For i = col - 1 To 2 Step -1
            
            ' Preenchendo a coluna nas planilhas criadas sempre que o mês muda
            If mes_aluguel <> Month(Sheets("Sheet1").Cells(1, i)) And qtd_mes < meses Then
            
                marc_col = marc_col + 1
                qtd_mes = qtd_mes + 1
                
                mes_aluguel = Month(Sheets("Sheet1").Cells(1, i))
                
                
                Sheets("Calculo").Cells(1, marc_col) = mes_aluguel & "/" & Year(Sheets("Sheet1").Cells(1, i))
                Sheets("Dias").Cells(1, marc_col) = mes_aluguel & "/" & Year(Sheets("Sheet1").Cells(1, i))
            
            End If
    Next
    
    'Deixando o cabeçalho em negrito
    Sheets("Dias").Range(Cells(1, 1), Cells(1, marc_col)).Font.Bold = True
    Sheets("Calculo").Activate
    Sheets("Calculo").Range(Cells(1, 1), Cells(1, marc_col)).Font.Bold = True
    
    
    'Reativando as funções
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True


End Sub


