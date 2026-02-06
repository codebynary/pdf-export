# 📄 Extrator de Fichas DOCX (Multi-Fichas)

## 🎯 Descrição

Este programa extrai automaticamente dados de múltiplas **Fichas de Registro de Colaboradores** em formato Word (.docx) e exporta tudo para uma planilha Excel organizada.

## ✨ Funcionalidades

- ✅ Processa múltiplos arquivos .docx de uma vez
- ✅ Interface gráfica para seleção de pasta
- ✅ Extração automática de **todos os campos** da ficha:
  - Dados pessoais (nome, CPF, RG, data de nascimento, etc.)
  - Filiação (nome do pai e mãe)
  - Documentos (CTPS, PIS, título de eleitor, etc.)
  - Endereço completo
  - Dados contratuais (função, salário, data de admissão, etc.)
  - Informações de rescisão
  - Dados da empresa
- ✅ Exportação para Excel com colunas organizadas
- ✅ Nome de arquivo com timestamp automático
- ✅ Tratamento de erros robusto

## 📋 Requisitos

```bash
pip install python-docx pandas openpyxl
```

## 🚀 Como Usar

### Método 1: Executar diretamente

```bash
python extrair_word_batch.py
```

1. Uma janela se abrirá pedindo para selecionar a pasta com os arquivos .docx
2. Selecione a pasta que contém as fichas de registro
3. O programa processará todos os arquivos automaticamente
4. Uma planilha Excel será gerada na mesma pasta com nome: `fichas_extraidas_YYYYMMDD_HHMMSS.xlsx`
5. Uma mensagem de sucesso aparecerá ao final

### Método 2: Usar como módulo

```python
from extrair_word_batch import ExtratorFichasWord

# Criar extrator
extrator = ExtratorFichasWord()

# Processar diretório
dados = extrator.processar_diretorio("C:/caminho/para/pasta")

# Exportar para Excel
extrator.exportar_para_excel(dados, "saida.xlsx")
```

## 📊 Campos Extraídos

### Identificação
- Código, Contrato, Nome, Matrícula eSocial

### Dados Pessoais
- Filiação (pai e mãe)
- Data de nascimento, raça/cor, sexo
- Naturalidade, nacionalidade
- Estado civil, deficiência, tipo sanguíneo

### Documentos
- CPF, RG (com data de emissão e órgão)
- CTPS (número, série, dígito)
- PIS (número e data de cadastramento)
- Título de eleitor (número, zona, seção)
- CNH (habilitação, categoria, validade)
- Grau de instrução

### Endereço
- Endereço completo (rua, número, complemento, bairro)
- Cidade, estado, CEP
- Telefone, celular, email

### Contrato de Trabalho
- Datas (admissão, registro)
- Função, CBO
- Salário inicial
- Forma e tipo de pagamento
- Insalubridade, periculosidade
- Sindicato, centro de custo, localização
- Horário de trabalho

### FGTS
- Número da conta
- Data de opção
- Banco depositário

### Rescisão
- Data de rescisão
- Aviso prévio
- Saldo FGTS
- Maior remuneração
- Causa da rescisão

### Empresa
- Nome do empregador
- CNPJ

### Metadados
- Arquivo de origem
- Data e hora da extração

## 📁 Estrutura de Saída

A planilha Excel gerada terá:
- **Uma linha por funcionário**
- **Uma coluna por campo**
- **Colunas prioritárias** (nome, CPF, RG, etc.) aparecem primeiro
- **Formatação automática** para facilitar leitura

## ⚠️ Observações

- O programa ignora arquivos temporários do Word (que começam com `~$`)
- Se um campo não existir no documento, a célula ficará vazia
- Erros de processamento são registrados no console
- O arquivo Excel é salvo com timestamp para evitar sobrescrever arquivos anteriores

## 🛠️ Solução de Problemas

### Erro: "No module named 'docx'"
```bash
pip install python-docx
```

### Erro: "No module named 'openpyxl'"
```bash
pip install openpyxl
```

### Nenhum arquivo processado
- Verifique se os arquivos têm extensão `.docx` (não `.doc`)
- Certifique-se de que não são arquivos corrompidos
- Verifique se você tem permissão de leitura na pasta

## 📝 Exemplo de Uso

```bash
# 1. Coloque todos os arquivos .docx em uma pasta
# 2. Execute o programa
python extrair_word_batch.py

# 3. Selecione a pasta na janela que abrir
# 4. Aguarde o processamento
# 5. Abra o arquivo Excel gerado!
```

## 🎨 Saída no Console

```
================================================================================
📄 EXTRATOR EM LOTE DE FICHAS DE REGISTRO (WORD → EXCEL)
================================================================================

🔍 Selecione o diretório com os arquivos .docx...
📁 Diretório selecionado: C:\Users\...\documentos

📁 Encontrados 15 arquivos .docx
================================================================================
[1/15] Processando: 1126 - Ficha de Registro-1.docx
[2/15] Processando: 1126 - Ficha de Registro-2.docx
...
[15/15] Processando: 1126 - Ficha de Registro-15.docx
================================================================================
✅ Processamento concluído! 15 arquivos processados.

💾 Exportando para Excel...
💾 Planilha salva em: fichas_extraidas_20260206_115620.xlsx
📊 Total de registros: 15
📋 Total de campos: 58

================================================================================
✅ PROCESSO CONCLUÍDO COM SUCESSO!
================================================================================
```

## 📞 Suporte

Em caso de dúvidas ou problemas, verifique:
1. Se todos os requisitos estão instalados
2. Se os arquivos Word não estão corrompidos
3. Se você tem permissões de leitura/escrita na pasta

---

**Desenvolvido com ❤️ por Antigravity AI**
