Cálculos Iniciais (Constantes do Ponto)
Fórmula 1: Pulso do padrão em L/P (I51)

Fórmula Excel: =SE(I50="";"";I50/1000)

Descrição: Converte o pulso de mL para Litros.

Entradas: I50 - 1° Ponto • mL/P

Fórmula 2: Pulso do Equipamento em • L/P (AD51)

Fórmula Excel: =SE(AD50="";"";AD50/1000)

Descrição: Converte o pulso do equipamento de mL para Litros.

Entradas: AD50 - Pulso do Equipamento • mL/P

Cálculos de Correção (Por Medição Individual - ex: linha 54)
Fórmula 3: Tempo de Coleta Corrigido • (s) (AA54)

Fórmula Excel: =SE(F54="";"";F54-(F54*'Estimativa da Incerteza'!$BU$23+'Estimativa da Incerteza'!$BW$23))

Descrição: Aplica fatores de correção ao tempo de coleta bruto.

Entradas:

F54 - Tempo de Coleta • (s)

'Estimativa da Incerteza'!$BU$23 (Constante)

'Estimativa da Incerteza'!$BW$23 (Constante)

Fórmula 4: Temperatura da Água Corrigida • °C (AD54)

Fórmula Excel: =SE(R54="";"";R54-(R54*'Estimativa da Incerteza'!$BU$26+'Estimativa da Incerteza'!$BW$26))

Descrição: Aplica fatores de correção à temperatura da água bruta.

Entradas:

R54 - Temperatura da Água • °C

'Estimativa da Incerteza'!$BU$26 (Constante)

'Estimativa da Incerteza'!$BW$26 (Constante)

Cálculos Principais (Por Medição Individual - ex: linha 54)
Fórmula 5: Totalização no Padrão Corrigido • L (L54)

Fórmula Excel: =SE(C54="";"";(C54*$I$51)-(($R$51+$U$51*(C54*$I$51/AA54*3600))/100*(C54*$I$51)))

Descrição: O cálculo mais crítico. Calcula o volume total corrigido do padrão.

Entradas:

C54 - Qtd de Pulsos do Padrão

R51 (Constante)

U51 (Constante)

Dependências Calculadas:

I51 - Pulso do padrão em L/P (Fórmula 1)

AA54 - Tempo de Coleta Corrigido • (s) (Fórmula 3)

Fórmula 6: Vazão de Referência • L/h (I54)

Fórmula Excel: =SE(C54="";"";L54/AA54*3600)

Descrição: Calcula a vazão de referência final.

Dependências Calculadas:

L54 - Totalização no Padrão Corrigido • L (Fórmula 5)

AA54 - Tempo de Coleta Corrigido • (s) (Fórmula 3)

Fórmula 7: Vazão do Medidor • L/h (X54)

Fórmula Excel: =SE(O54="";"";SE(OU($X$16 = "Visual com início dinâmico";$X$16="Visual com início estática" );O54;(O54/AA54)*3600))

Descrição: Calcula a vazão do medidor. O cálculo muda dependendo do modo de calibração.

Entradas:

O54 - Leitura no Medidor • L

X16 - Modo de Calibração

Dependências Calculadas:

AA54 - Tempo de Coleta Corrigido • (s) (Fórmula 3)

Fórmula 8: Erro % (U54)

Fórmula Excel: =SE(O54="";"";(O54-L54)/L54*100)

Descrição: Calcula o erro percentual da medição.

Entradas: O54 - Leitura no Medidor • L

Dependências Calculadas: L54 - Totalização no Padrão Corrigido • L (Fórmula 5)

Cálculos de Agregação (Resumo do Ponto)
Fórmula 9: Vazão Média • L/h (I57)

Fórmula Excel: =SE(I54="";"";MÉDIA(I54:I56))

Descrição: Calcula a média das 3 vazões de referência do ponto.

Dependências Calculadas: I54, I55, I56 (Resultados da Fórmula 6 para cada medição)

Fórmula 10: Tendência (U57)

Fórmula Excel: =SE(U54="";"";MÉDIA(U54:U56))

Descrição: Calcula a média dos 3 erros percentuais do ponto.

Dependências Calculadas: U54, U55, U56 (Resultados da Fórmula 8 para cada medição)

Fórmula 11: DESVIO PADRÃO AMOSTRAL (AD57)

Fórmula Excel: =SE(U54="";"";STDEV.S(U54:U56))

Descrição: Calcula o desvio padrão (repetibilidade) dos 3 erros.

Dependências Calculadas: U54, U55, U56 (Resultados da Fórmula 8 para cada medição)