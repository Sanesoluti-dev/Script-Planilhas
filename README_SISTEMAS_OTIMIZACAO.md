# SISTEMAS DE OTIMIZAÇÃO PARA AJUSTE DE TEMPOS DE COLETA

## Visão Geral

Este projeto implementa múltiplos sistemas de otimização para resolver o problema complexo de ajuste de tempos de coleta em planilhas de calibração, mantendo os valores do certificado matematicamente idênticos aos originais.

## Problema a Ser Resolvido

**Objetivo**: Gerar uma nova versão de um arquivo Excel de calibração onde os valores de "Tempo de Coleta" em todas as medições de um ponto estejam próximos de 360 segundos, mantendo os valores numéricos finais do certificado matematicamente idênticos aos originais.

**Restrição Crítica**: Os valores calculados e agregados na aba "Emissão do Certificado" devem ser matematicamente idênticos aos do arquivo original.

## Sistemas Implementados

### 1. Sistema de Otimização Multi-variável (`sistema_de_otimizacao.py`)

**Características**:
- ✅ Implementa algoritmo de otimização iterativa ("Atingir Meta" avançado)
- ✅ Usa biblioteca scipy.optimize com método Nelder-Mead
- ✅ Precisão Decimal de 50 dígitos
- ✅ Busca pelo ponto de equilíbrio perfeito
- ✅ Preserva proporções internas de cada ponto

**Dependências**:
```bash
pip install scipy>=1.11.0
```

**Uso**:
```bash
python sistema_de_otimizacao.py
```

**Vantagens**:
- Algoritmo de otimização comprovado (Nelder-Mead)
- Convergência robusta
- Precisão ultra-alta

**Limitações**:
- Requer instalação do scipy
- Pode ser computacionalmente intensivo

### 2. Sistema de Otimização Simplificado (`sistema_de_otimizacao_simples.py`)

**Características**:
- ✅ Algoritmo próprio de busca em grade
- ✅ Não depende de bibliotecas externas complexas
- ✅ Precisão Decimal de 50 dígitos
- ✅ Busca adaptativa simples
- ✅ Implementação independente

**Uso**:
```bash
python sistema_de_otimizacao_simples.py
```

**Vantagens**:
- Sem dependências externas complexas
- Implementação transparente
- Fácil de entender e modificar

**Limitações**:
- Busca menos sofisticada
- Pode não convergir para soluções ótimas

### 3. Sistema de Otimização Avançado (`sistema_de_otimizacao_avancado.py`)

**Características**:
- ✅ Busca adaptativa em múltiplas fases
- ✅ Algoritmo próprio sofisticado
- ✅ Precisão Decimal de 50 dígitos
- ✅ Busca ampla → refinada → ultra-refinada
- ✅ Tolerância ajustável

**Uso**:
```bash
python sistema_de_otimizacao_avancado.py
```

**Vantagens**:
- Busca mais sofisticada que o sistema simples
- Múltiplas fases de otimização
- Controle granular sobre o processo

**Limitações**:
- Ainda pode não convergir para soluções ótimas
- Processo computacionalmente intensivo

## Arquitetura dos Sistemas

### FASE 1: Preparação e Análise

#### 1.1 Extração de Constantes
- Extrai constantes das células fixas da planilha
- $I$51 - Pulso do padrão em L/P
- $R$51 - Temperatura constante
- $U$51 - Fator correção temperatura
- $X$16 - Tipo de medição

#### 1.2 Extração de Dados Originais
- Lê todos os dados brutos da planilha
- Identifica pontos de calibração
- Extrai 3 leituras por ponto

#### 1.3 Cálculo dos "Valores Sagrados" (Alvos)
- Executa motor de cálculo com dados originais
- Obtém valores finais da aba "Emissão do Certificado"
- Estes são os alvos imutáveis

#### 1.4 Cálculo das Proporções Internas
- Para cada ponto, calcula proporções internas
- Usa primeira medição como "mestre"
- Preserva relações entre leituras

### FASE 2: Otimização Iterativa Global

#### 2.1 Função de Custo
```python
custo_total = (erro_vazao_ref**2) + (erro_vazao_med**2)
```

#### 2.2 Variáveis de Otimização
1. **Tempo de Coleta**: Valor próximo a 360 segundos
2. **Pulsos Mestre**: Quantidade de pulsos da primeira medição

#### 2.3 Processo de Otimização
- Busca iterativa pelos valores ótimos
- Recalcula todos os valores usando proporções
- Minimiza função de custo

### FASE 3: Saída e Geração do Arquivo Final

#### 3.1 Geração dos Dados Otimizados
- Aplica resultados da otimização
- Mantém proporções internas
- Arredonda pulsos para inteiros

#### 3.2 Verificação de Precisão
- Recalcula valores com dados otimizados
- Compara com valores sagrados originais
- Verifica se diferenças são aceitáveis

#### 3.3 Geração da Planilha Excel
- Cria cópia do arquivo original
- Aplica valores otimizados com alta precisão
- Garante exibição formatada

## Motor de Cálculo

### Fórmulas Críticas Implementadas

#### Totalização no Padrão Corrigido • L
```python
def calcular_totalizacao_padrao_corrigido(self, pulsos_padrao, tempo_coleta):
    # Fórmula: =SE(C54="";"";(C54*$I$51)-(($R$51+$U$51*(C54*$I$51/AA54*3600))/100*(C54*$I$51)))
```

#### Vazão de Referência • L/h
```python
def calcular_vazao_referencia(self, totalizacao, tempo_coleta):
    # Fórmula: =SE(C54="";"";L54/AA54*3600)
```

#### Vazão do Medidor • L/h
```python
def calcular_vazao_medidor(self, leitura_medidor, tempo_coleta, tipo_medicao):
    # Fórmula: =SE(O54="";"";SE(OU($X$16 = "Visual com início dinâmico";$X$16="Visual com início estática" );O54;(O54/AA54)*3600))
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

## Saídas dos Sistemas

### Arquivos Gerados
1. **Planilha Otimizada**: `SAN-038-25-09_OTIMIZADO.xlsx`
2. **Relatório JSON**: `relatorio_otimizacao.json`
3. **Relatório Texto**: `relatorio_otimizacao.txt`

### Verificações Automáticas
1. **Precisão dos valores sagrados**
2. **Convergência da otimização**
3. **Preservação das proporções**
4. **Exatidão dos valores do certificado**

## Configurações de Precisão

### Precisão Ultra-alta
```python
getcontext().prec = 50
```

### Tolerâncias
- **Tolerância de convergência**: 1e-30 (sistema completo)
- **Tolerância de verificação**: 1e-20 (sistema completo)
- **Tolerância de verificação**: 1e-5 (sistema avançado)
- **Precisão Decimal**: 50 dígitos

## Status Atual

### Sistema Multi-variável (scipy)
- ✅ Implementado
- ❌ Requer instalação do scipy
- ⚠️ Não testado completamente

### Sistema Simplificado
- ✅ Implementado
- ✅ Funciona sem dependências externas
- ❌ Convergência limitada
- ⚠️ Precisão insuficiente

### Sistema Avançado
- ✅ Implementado
- ✅ Busca mais sofisticada
- ❌ Convergência ainda limitada
- ⚠️ Precisão insuficiente

## Próximos Passos

### Melhorias Necessárias

1. **Refinamento do Algoritmo de Otimização**
   - Implementar gradiente descendente
   - Usar métodos de otimização mais sofisticados
   - Ajustar parâmetros de busca

2. **Análise do Problema**
   - Investigar se o problema tem solução única
   - Verificar se as restrições são compatíveis
   - Analisar a sensibilidade dos parâmetros

3. **Implementação de Métodos Alternativos**
   - Programação linear
   - Otimização por enxame de partículas
   - Algoritmos genéticos

### Recomendações

1. **Para Uso Imediato**: Sistema Simplificado
   - Mais fácil de entender e modificar
   - Funciona sem dependências externas
   - Base para melhorias futuras

2. **Para Desenvolvimento**: Sistema Avançado
   - Melhor estrutura para experimentos
   - Busca mais sofisticada
   - Base para implementar novos algoritmos

3. **Para Produção**: Sistema Multi-variável
   - Algoritmo comprovado (Nelder-Mead)
   - Melhor convergência teórica
   - Requer instalação do scipy

## Conclusão

Os sistemas implementados demonstram uma abordagem sistemática para resolver o problema complexo de otimização de tempos de coleta. Embora ainda não tenham convergido para soluções ótimas, fornecem uma base sólida para futuras melhorias e investigações.

A arquitetura modular permite fácil experimentação com diferentes algoritmos de otimização, e a documentação detalhada facilita o entendimento e modificação dos sistemas. 