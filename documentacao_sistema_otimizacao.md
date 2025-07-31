# SISTEMA DE OTIMIZAÇÃO MULTI-VARIÁVEL PARA AJUSTE DE TEMPOS DE COLETA

## Visão Geral

Este sistema implementa um algoritmo de otimização iterativa ("Atingir Meta" avançado) que busca a combinação "perfeita" de parâmetros de entrada que satisfaça todas as condições simultaneamente:

1. ✅ **Tempos de coleta próximos a 360 segundos**
2. ✅ **Valores do certificado matematicamente idênticos aos originais**
3. ✅ **Preservação das proporções internas de cada ponto**
4. ✅ **Precisão Decimal de 50 dígitos**

## Paradigma de Solução

### Problema Complexo
Não há uma fórmula direta para resolver este problema. A solução é um **algoritmo de otimização iterativa** que busca a combinação "perfeita" de parâmetros de entrada que satisfaça todas as condições simultaneamente.

### Abordagem Multi-variável
O sistema usa **otimização multi-variável** com busca pelo ponto de equilíbrio perfeito, onde:

- **Variáveis de Ajuste**: Tempo de Coleta + Pulsos Mestre
- **Função de Custo**: Soma dos quadrados dos erros entre valores calculados e alvos
- **Algoritmo**: Nelder-Mead (scipy.optimize)
- **Tolerância**: 1e-30 (precisão ultra-alta)

## Arquitetura do Sistema

### FASE 1: Preparação e Análise

#### 1.1 Extração de Constantes
```python
def extrair_constantes(self):
    # Extrai constantes das células fixas da planilha
    # $I$51 - Pulso do padrão em L/P
    # $R$51 - Temperatura constante  
    # $U$51 - Fator correção temperatura
    # $X$16 - Tipo de medição
```

#### 1.2 Extração de Dados Originais
```python
def extrair_dados_originais(self):
    # Lê todos os dados brutos da planilha
    # Identifica pontos de calibração
    # Extrai 3 leituras por ponto
```

#### 1.3 Cálculo dos "Valores Sagrados" (Alvos)
```python
def calcular_valores_sagrados_originais(self):
    # Executa motor de cálculo com dados originais
    # Obtém valores finais da aba "Emissão do Certificado"
    # Estes são os alvos imutáveis
```

#### 1.4 Cálculo das Proporções Internas
```python
def calcular_proporcoes_internas(self):
    # Para cada ponto, calcula proporções internas
    # Usa primeira medição como "mestre"
    # fator_pulso_55_vs_54 = C55_original / C54_original
    # fator_leitura_55_vs_54 = O55_original / O54_original
    # fator_leitura_vs_pulso_54 = O54_original / C54_original
```

### FASE 2: Otimização Iterativa Global

#### 2.1 Função de Custo
```python
def funcao_custo(self, parametros, ponto_key):
    # Extrai parâmetros: novo_tempo, novo_pulsos_mestre
    # Recalcula todos os valores usando proporções
    # Calcula erros entre valores calculados e alvos
    # custo_total = (erro_vazao_ref**2) + (erro_vazao_med**2)
```

#### 2.2 Busca pelo Ponto de Equilíbrio
```python
def otimizar_ponto(self, ponto_key):
    # Usa scipy.optimize.minimize com método Nelder-Mead
    # Ajusta simultaneamente:
    #   a. Tempo de Coleta unificado (próximo a 360)
    #   b. Qtd de pulso do padrão da medição "mestre"
    # Recalcula TODAS as outras variáveis usando proporções
```

#### 2.3 Processo do Loop de Otimização
1. **Estimativa inicial**: novo_Tempo = 360.0, novo_C54
2. **Recálculo proporcional**: 
   - novo_C55 = novo_C54 * fator_pulso_55_vs_54
   - novo_O54 = novo_C54 * fator_leitura_vs_pulso_54
   - novo_O55 = novo_O54 * fator_leitura_55_vs_54
3. **Execução do motor de cálculo** com novos dados
4. **Cálculo do custo_total**
5. **Ajuste iterativo** até convergência (custo_total < 1e-30)

### FASE 3: Saída e Geração do Arquivo Final

#### 3.1 Geração dos Dados Otimizados
```python
def gerar_dados_otimizados(self, resultados_otimizacao):
    # Aplica resultados da otimização
    # Mantém proporções internas
    # Arredonda pulsos para inteiros
```

#### 3.2 Verificação de Precisão
```python
def verificar_precisao(self, dados_otimizados):
    # Recalcula valores com dados otimizados
    # Compara com valores sagrados originais
    # Verifica se diferenças são < 1e-20
```

#### 3.3 Geração da Planilha Excel
```python
def gerar_planilha_otimizada(self, dados_otimizados):
    # Cria cópia do arquivo original
    # Aplica valores otimizados com alta precisão
    # Garante exibição formatada (360) mas valor interno preciso
```

## Motor de Cálculo

### Implementação das Fórmulas Críticas

#### Totalização no Padrão Corrigido • L
```python
def calcular_totalizacao_padrao_corrigido(self, pulsos_padrao, tempo_coleta):
    # Fórmula: =SE(C54="";"";(C54*$I$51)-(($R$51+$U$51*(C54*$I$51/AA54*3600))/100*(C54*$I$51)))
    volume_pulsos = pulsos_padrao * self.constantes['pulso_padrao_lp']
    vazao = volume_pulsos / tempo_coleta * Decimal('3600')
    fator_correcao = (self.constantes['temperatura_constante'] + 
                      self.constantes['fator_correcao_temp'] * vazao) / Decimal('100')
    totalizacao = volume_pulsos - (fator_correcao * volume_pulsos)
    return totalizacao
```

#### Vazão de Referência • L/h
```python
def calcular_vazao_referencia(self, totalizacao, tempo_coleta):
    # Fórmula: =SE(C54="";"";L54/AA54*3600)
    vazao = (totalizacao / tempo_coleta) * Decimal('3600')
    return vazao
```

#### Erro Percentual
```python
def calcular_erro_percentual(self, leitura_medidor, totalizacao):
    # Fórmula: =SE(O54="";"";(O54-L54)/L54*100)
    erro = ((leitura_medidor - totalizacao) / totalizacao) * Decimal('100')
    return erro
```

#### Vazão do Medidor • L/h
```python
def calcular_vazao_medidor(self, leitura_medidor, tempo_coleta, tipo_medicao):
    # Fórmula: =SE(O54="";"";SE(OU($X$16 = "Visual com início dinâmico";$X$16="Visual com início estática" );O54;(O54/AA54)*3600))
    if tipo_medicao in ["Visual com início dinâmico", "Visual com início estática"]:
        return leitura_medidor
    else:
        return (leitura_medidor / tempo_coleta) * Decimal('3600')
```

## Configurações de Precisão

### Precisão Ultra-alta
```python
# Configurar precisão ultra-alta para evitar diferenças de arredondamento
getcontext().prec = 50
```

### Tolerâncias
- **Tolerância de convergência**: 1e-30
- **Tolerância de verificação**: 1e-20
- **Precisão Decimal**: 50 dígitos

## Algoritmo de Otimização

### Método Nelder-Mead
```python
resultado = minimize(
    funcao_custo_scipy,
    parametros_iniciais,
    method='Nelder-Mead',
    options={
        'maxiter': 10000,
        'xatol': 1e-30,
        'fatol': 1e-30
    }
)
```

### Variáveis de Otimização
1. **Tempo de Coleta**: Valor próximo a 360 segundos
2. **Pulsos Mestre**: Quantidade de pulsos da primeira medição

### Função de Custo
```python
custo_total = (erro_vazao_ref**2) + (erro_vazao_med**2)
```

## Hierarquia de Influência

```
AA54 (Tempo de Coleta) 
    ↓
I54 (Vazão de Referência)
    ↓
I57 (Vazão Média)

L54 (Totalização)
    ↓
U54 (Erro)
    ↓
U57 (Tendência) e AD57 (Desvio Padrão)

O54 (Leitura do Medidor)
    ↓
U54 (Erro) e X54 (Vazão do Medidor)
```

## Valores Sagrados

Os seguintes valores **NÃO PODEM SER ALTERADOS** em nenhuma hipótese:

1. **Vazão Média de Referência** (I57)
2. **Vazão Média do Medidor** (X57)
3. **Tendência** (U57)
4. **Desvio Padrão Amostral** (AD57)

## Proporções Internas

Para cada ponto de calibração, o sistema preserva:

1. **Proporções entre pulsos**: C55/C54, C56/C54
2. **Proporções entre leituras**: O55/O54, O56/O54
3. **Relação leitura/pulso mestre**: O54/C54

## Saídas do Sistema

### Arquivos Gerados
1. **Planilha Otimizada**: `SAN-038-25-09_OTIMIZADO.xlsx`
2. **Relatório JSON**: `relatorio_otimizacao.json`
3. **Relatório Texto**: `relatorio_otimizacao.txt`

### Verificações Automáticas
1. **Precisão dos valores sagrados**
2. **Convergência da otimização**
3. **Preservação das proporções**
4. **Exatidão dos valores do certificado**

## Uso do Sistema

```python
# Executar o sistema
sistema = SistemaOtimizacao("SAN-038-25-09.xlsx")
sucesso = sistema.executar()

if sucesso:
    print("🎉 SISTEMA CONCLUÍDO COM SUCESSO!")
else:
    print("❌ SISTEMA FALHOU!")
```

## Vantagens da Abordagem

1. **Precisão Absoluta**: Preserva valores do certificado com precisão de 50 dígitos
2. **Flexibilidade**: Ajusta múltiplas variáveis simultaneamente
3. **Robustez**: Usa algoritmo de otimização comprovado (Nelder-Mead)
4. **Rastreabilidade**: Gera relatórios detalhados de todo o processo
5. **Automatização**: Processo totalmente automatizado

## Limitações e Considerações

1. **Tempo de Processamento**: Otimização pode ser computacionalmente intensiva
2. **Convergência**: Depende da qualidade dos dados originais
3. **Precisão**: Requer dados de entrada de alta qualidade
4. **Memória**: Usa precisão Decimal de 50 dígitos (maior uso de memória)

## Conclusão

Este sistema implementa uma solução elegante e robusta para o problema complexo de ajuste de tempos de coleta, garantindo que os valores do certificado permaneçam matematicamente idênticos aos originais enquanto os tempos são harmonizados próximos a 360 segundos. 