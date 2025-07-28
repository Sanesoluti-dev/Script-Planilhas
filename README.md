# Formatador de Excel em Massa

Um aplicativo desktop simples para processar múltiplos arquivos Excel (.xlsx) e aplicar formatação "General" para revelar a precisão total dos números.

## 🎯 Funcionalidades

- **Interface gráfica intuitiva** com Tkinter
- **Seleção múltipla de arquivos** Excel
- **Processamento em lote** de centenas de arquivos
- **Formatação automática** para "General" em todas as células
- **Preservação dos arquivos originais** (cria cópias na pasta `output_formatado`)
- **Log detalhado** do processamento
- **Barra de progresso** em tempo real
- **Relatório de sucessos e erros**

## 📋 Requisitos

- Python 3.7 ou superior
- Biblioteca `openpyxl`
- Tkinter (geralmente já vem com Python)

## 🚀 Instalação

1. **Clone ou baixe** os arquivos do projeto
2. **Instale as dependências**:
   ```bash
   pip install -r requirements.txt
   ```
   
   Ou instale manualmente:
   ```bash
   pip install openpyxl
   ```

## 💻 Como Usar

1. **Execute o aplicativo**:
   ```bash
   python formatador_excel_massa.py
   ```

2. **Selecione os arquivos**:
   - Clique em "📁 Selecionar Arquivos Excel"
   - Navegue até a pasta com seus arquivos
   - Selecione um ou múltiplos arquivos .xlsx
   - Clique em "Abrir"

3. **Inicie o processamento**:
   - Clique em "⚡ Iniciar Processamento"
   - Acompanhe o progresso na barra e no log
   - Aguarde a conclusão

4. **Acesse os resultados**:
   - Os arquivos processados estarão na pasta `output_formatado`
   - Clique em "📂 Abrir Pasta de Saída" para acessar

## 🔧 Como Funciona

O aplicativo:

1. **Carrega** cada arquivo Excel usando `openpyxl`
2. **Itera** por todas as planilhas do arquivo
3. **Processa** cada célula com dados
4. **Aplica** formatação `number_format = 'General'`
5. **Salva** uma cópia na pasta de saída
6. **Preserva** os arquivos originais intactos

## 📊 Formatação "General"

A formatação "General" do Excel:
- **Remove** formatações numéricas personalizadas
- **Revela** a precisão total dos números
- **Mostra** todos os dígitos significativos
- **Elimina** arredondamentos visuais

## 🛡️ Segurança

- **Nunca modifica** os arquivos originais
- **Cria cópias** na pasta `output_formatado`
- **Processamento em thread** separada (não trava a interface)
- **Tratamento de erros** robusto
- **Log detalhado** para auditoria

## 📁 Estrutura de Arquivos

```
projeto/
├── formatador_excel_massa.py    # Aplicativo principal
├── requirements.txt             # Dependências
├── README.md                    # Este arquivo
└── output_formatado/            # Pasta criada automaticamente
    ├── arquivo1_formatado.xlsx
    ├── arquivo2_formatado.xlsx
    └── ...
```

## 🎨 Interface

- **Design moderno** com ícones e cores
- **Layout responsivo** que se adapta ao tamanho da janela
- **Feedback visual** em tempo real
- **Controles intuitivos** e bem organizados

## 🔍 Log de Processamento

O aplicativo registra:
- ✅ Arquivos processados com sucesso
- ❌ Erros encontrados
- 📊 Estatísticas finais
- ⏱️ Timestamps de cada ação

## 🚨 Solução de Problemas

**Erro: "openpyxl não encontrado"**
```bash
pip install openpyxl
```

**Erro: "Tkinter não encontrado"**
- No Windows: Reinstale Python marcando "tcl/tk and IDLE"
- No Linux: `sudo apt-get install python3-tk`
- No macOS: `brew install python-tk`

**Arquivos não aparecem na lista**
- Verifique se são arquivos .xlsx válidos
- Tente selecionar arquivos individuais primeiro

## 📈 Performance

- **Processamento otimizado** para grandes volumes
- **Threading** para não travar a interface
- **Progresso em tempo real**
- **Eficiente** para centenas de arquivos

## 🤝 Contribuições

Sinta-se à vontade para:
- Reportar bugs
- Sugerir melhorias
- Contribuir com código
- Melhorar a documentação

## 📄 Licença

Este projeto é de código aberto e pode ser usado livremente.

---

**Desenvolvido para resolver problemas de precisão numérica em planilhas Excel em massa!** 🎯 