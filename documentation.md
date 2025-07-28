# -*- coding: utf-8 -*-
"""
DOCUMENTAÇÃO DO SISTEMA DE AJUSTE DE TEMPOS DE COLETA
=====================================================

1. Objetivo do Projeto

O objetivo deste projeto é criar um sistema em Python para automatizar a correção de dados em planilhas de calibração de vazão. O sistema deve ler os dados de medição de um arquivo Excel, identificar inconsistências nos tempos de coleta, e ajustar um conjunto de parâmetros de forma proporcional e precisa, gerando uma nova planilha corrigida.

Este processo deve ser guiado pelo princípio fundamental descrito abaixo.

2. O Princípio Fundamental do Software

A integridade dos resultados finais do certificado é absoluta e inalterável. Os seguintes valores, calculados para cada ponto de calibração, NÃO PODEM MUDAR EM NENHUMA HIPÓTESE, nem mesmo na última casa decimal:

Vazão Média

Tendência

Desvio Padrão Amostral

Todo o processo de ajuste deve ser feito de tal forma que, ao final, uma verificação matemática prove que esses três valores permaneceram idênticos aos originais. A precisão é o requisito mais crítico.

3. O Problema a Ser Resolvido

O problema central está na aba "Coleta de Dados". Para um único ponto de calibração, existem múltiplas medições (geralmente 3). Os valores na coluna "Tempo de Coleta" para essas medições deveriam ser conceitualmente iguais (ex: 170s), mas devido a pequenas variações no processo, eles podem estar ligeiramente diferentes (ex: 169.98s, 170.01s, 170.00s).

O software deve corrigir essa inconsistência.

4. Lógica de Correção e Ajuste Proporcional

O sistema deve executar a seguinte lógica, que será o coração do @ajustador_tempo_coleta.py:

4.1. Harmonização do Tempo de Coleta:

O primeiro passo é definir um "Tempo de Coleta" unificado para todas as medições de um ponto. Esse valor unificado será o novo padrão (por exemplo, 170.00000).

O script deve substituir os tempos de coleta originais e ligeiramente diferentes por este novo valor idêntico para todas as medições do ponto.

4.2. Ajuste Proporcional para Manter a Vazão Constante:

A fórmula fundamental da vazão média é =SE(I54="";"";MÉDIA(I54:I56)).

Ao forçar um novo Tempo, a Vazão calculada seria alterada. Para evitar isso e manter a Vazão Média constante, os valores relacionados ao Volume devem ser ajustados na mesma proporção.

Para cada medição individual, o sistema deve calcular um fator de ajuste:

fator = Novo_Tempo_de_Coleta / Tempo_de_Coleta_Original

Este fator deve então ser aplicado aos outros parâmetros que podem ser alterados:

Parâmetros que PODEM e DEVEM ser alterados:

Tempo de coleta: Este é o gatilho. Será substituído pelo novo valor unificado.

Qtd de pulso do padrão: Este valor é diretamente proporcional ao volume do padrão. Deve ser ajustado da seguinte forma:

novo_qtd_pulsos = qtd_pulsos_original * fator

Leitura no medidor: Este valor é diretamente proporcional ao volume medido pelo instrumento. Deve ser ajustado da seguinte forma:

nova_leitura_medidor = leitura_medidor_original * fator

Temperatura da água: Embora a temperatura influencie a densidade e outros fatores de correção, para este sistema, vamos assumir que seu impacto direto na Vazão Média final pode ser compensado por um pequeno ajuste fino para garantir a exatidão absoluta. A principal correção, no entanto, deve ser nos parâmetros de volume (pulsos e leitura).

5. Fórmulas Críticas do Certificado

5.1. Totalização no Padrão Corrigido • L:
=SE(C54="";"";(C54*$I$51)-(($R$51+$U$51*(C54*$I$51/AA54*3600))/100*(C54*$I$51)))

Onde:
- C54 = Pulsos do Padrão
- I$51 = Pulso do padrão em L/P (0.2000)
- R$51 = Temperatura da água
- U$51 = Fator de correção da temperatura
- AA54 = Tempo de coleta

5.2. Fórmula do Certificado - Valor 1:
=SE(I74="---";"---";DEF.NÚM.DEC(MÉDIA('Coleta de Dados'!L54:L56);'Estimativa da Incerteza'!BQ10))

Onde:
- L54:L56 = Valores da "Totalização no Padrão Corrigido • L"
- BQ10 = Número de casas decimais da incerteza

5.3. Fórmula do Certificado - Valor 2:
=SE(I74="---";"---";SE('Coleta de Dados'!I14="TOTALIZADOR DE VOLUME DE ÁGUA";DEF.NÚM.DEC(MÉDIA('Coleta de Dados'!O54:O56);'Estimativa da Incerteza'!BQ10);DEF.NÚM.DEC(MÉDIA('Coleta de Dados'!X54:Z56);'Estimativa da Incerteza'!BQ10)))

Onde:
- O54:O56 = Leituras no medidor
- X54:Z56 = Outros valores de medição
- BQ10 = Número de casas decimais da incerteza

6. Passos de Implementação para o Sistema

O sistema deve ser construído seguindo os passos abaixo:

Extração de Dados (@extrator_pontos_calibracao.py):

Leia o arquivo Excel de entrada (.xlsx).

Extraia todos os parâmetros de entrada brutos e constantes das abas "Coleta de Dados" e "Estimativa da Incerteza".

Utilize a biblioteca Decimal de Python para todos os valores numéricos para garantir a precisão necessária.

Cálculo da Linha de Base (Baseline):

Com os dados originais, execute o motor de cálculo completo para obter os valores originais de "Vazão Média", "Tendência" e "Desvio Padrão".

Armazene esses três valores como o "resultado sagrado" que deve ser alcançado no final.

Execução do Ajuste (@ajustador_tempo_coleta.py):

Aplique a lógica de "Harmonização do Tempo de Coleta" e "Ajuste Proporcional" descrita na seção 4 para todos os pontos de medição. Isso irá gerar um novo conjunto de dados de entrada corrigidos.

Recálculo e Verificação:

Com os novos dados de entrada corrigidos, execute novamente o motor de cálculo completo.

Compare os novos resultados de "Vazão Média", "Tendência" e "Desvio Padrão" com os valores da linha de base salvos no passo 2.

A diferença deve ser zero. O sistema deve confirmar essa exatidão.

Saída (Output):

O resultado final do sistema deve ser um novo arquivo Excel, uma cópia do original, mas com os valores ajustados e corrigidos na aba "Coleta de Dados".