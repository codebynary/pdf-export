# 📄 Extrator de Fichas Word - GUI Profissional

## 🎨 Interface Gráfica Moderna

Aplicação desktop com interface gráfica profissional para extração em lote de dados de fichas de registro em formato Word (.docx).

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![License](https://img.shields.io/badge/license-MIT-orange)

---

## ✨ Características da Interface

### 🎯 Design Moderno
- **Cores profissionais** com esquema azul moderno
- **Tipografia limpa** usando Segoe UI
- **Layout responsivo** que se adapta ao tamanho da janela
- **Visual premium** com espaçamento e padding adequados

### 📊 Componentes Principais

1. **Cabeçalho**
   - Título com ícone
   - Subtítulo descritivo
   - Exibição da versão (v1.0.0)

2. **Seleção de Diretório**
   - Campo de texto mostrando caminho selecionado
   - Botão "Selecionar Pasta" com ícone
   - Contador automático de arquivos .docx encontrados

3. **Barra de Progresso**
   - Status textual do que está sendo processado
   - Barra de progresso visual
   - Porcentagem em tempo real

4. **Log de Atividades**
   - Console com scroll automático
   - Mensagens com timestamp
   - Cores diferenciadas por tipo:
     - 🔵 **Azul**: Informações gerais
     - 🟢 **Verde**: Sucesso
     - 🟡 **Amarelo**: Avisos
     - 🔴 **Vermelho**: Erros
     - 🟣 **Roxo**: Cabeçalhos

5. **Botões de Ação**
   - ▶ **Processar Arquivos**: Inicia a extração
   - 🗑️ **Limpar Log**: Limpa o console
   - ✖ **Sair**: Fecha a aplicação

6. **Rodapé**
   - Créditos do desenvolvedor

---

## 🚀 Como Usar

### 1. Executar a Aplicação

```bash
python extrator_word_gui.py
```

### 2. Selecionar Pasta

1. Clique em **"Selecionar Pasta"**
2. Navegue até a pasta com os arquivos .docx
3. Confirme a seleção

### 3. Processar

1. Verifique no log quantos arquivos foram encontrados
2. Clique em **"▶ Processar Arquivos"**
3. Acompanhe o progresso em tempo real:
   - Barra de progresso visual
   - Porcentagem atualizada
   - Log detalhado de cada arquivo
   - Status atual do processamento

### 4. Resultado

- Planilha Excel gerada automaticamente na mesma pasta
- Nome do arquivo: `fichas_extraidas_YYYYMMDD_HHMMSS.xlsx`
- Mensagem de sucesso ao final

---

## 📋 Requisitos

```bash
pip install python-docx pandas openpyxl
```

Ou use o arquivo de requisitos:

```bash
pip install -r requirements_word.txt
```

---

## 🎯 Funcionalidades Técnicas

### Processamento Assíncrono
- Usa **threading** para não travar a interface
- Interface permanece responsiva durante processamento
- Atualizações em tempo real

### Tratamento de Erros
- Validações antes de processar
- Mensagens de erro descritivas
- Continuação do processamento mesmo com erros individuais

### Feedback Visual
- Progresso percentual preciso
- Log colorido e organizado
- Timestamps em todas as mensagens
- Confirmações de sucesso/erro

---

## 🎨 Paleta de Cores

```python
Primária:    #2563eb  (Azul moderno)
Secundária:  #1e40af  (Azul escuro)
Sucesso:     #10b981  (Verde)
Fundo:       #f8fafc  (Cinza claro)
Texto:       #1e293b  (Cinza escuro)
Borda:       #e2e8f0  (Cinza médio)
```

---

## 📸 Fluxo de Uso

```
1. Abrir aplicação
   ↓
2. Selecionar pasta com .docx
   ↓
3. Ver confirmação no log (X arquivos encontrados)
   ↓
4. Clicar em "Processar Arquivos"
   ↓
5. Acompanhar progresso em tempo real
   ↓
6. Receber mensagem de sucesso
   ↓
7. Abrir Excel gerado
```

---

## 🛡️ Segurança

- Confirmação antes de sair durante processamento
- Validação de diretório antes de processar
- Tratamento de exceções em todos os níveis
- Arquivos temporários do Word (~$) são ignorados

---

## 📊 Exemplo de Log

```
[12:04:35] Diretório selecionado: C:\Users\...\documentos
[12:04:35] Encontrados 15 arquivo(s) .docx
[12:04:40] ============================================================
[12:04:40] Iniciando processamento...
[12:04:40] ============================================================
[12:04:41] [1/15] 1126 - Ficha de Registro-1.docx
[12:04:41]   ✓ Extraído com sucesso
[12:04:42] [2/15] 1126 - Ficha de Registro-2.docx
[12:04:42]   ✓ Extraído com sucesso
...
[12:05:10] ============================================================
[12:05:10] Gerando planilha Excel...
[12:05:11] ✓ Planilha salva: fichas_extraidas_20260206_120511.xlsx
[12:05:11] ============================================================
[12:05:11] PROCESSAMENTO CONCLUÍDO COM SUCESSO!
[12:05:11] ============================================================
```

---

## 🔧 Solução de Problemas

### Botão "Processar" desabilitado
- Certifique-se de ter selecionado uma pasta
- Verifique se há arquivos .docx na pasta

### Interface não abre
- Verifique se o Python está instalado corretamente
- Confirme que tkinter está disponível (vem com Python)

### Erro durante processamento
- Verifique o log para detalhes
- Confirme que os arquivos não estão corrompidos
- Certifique-se de ter permissões de leitura/escrita

---

## 📝 Versão

**v1.0.0** - Lançamento inicial
- Interface gráfica completa
- Processamento em lote
- Barra de progresso
- Log colorido
- Exportação para Excel

---

## 🎯 Próximas Melhorias

- [ ] Ícone personalizado da aplicação
- [ ] Tema claro/escuro
- [ ] Configurações personalizáveis
- [ ] Preview dos dados antes de exportar
- [ ] Suporte a múltiplos formatos de saída
- [ ] Histórico de processamentos

---

**Desenvolvido com ❤️ por Antigravity AI**
