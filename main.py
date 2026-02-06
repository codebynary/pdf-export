import customtkinter as ctk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
import os
import threading
import concurrent.futures
import time

# Configuração do tema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("dark-blue")

# ==============================
# FUNÇÕES DE EXTRAÇÃO (Top-level)
# ==============================
def extrair_campos(texto_funcionario):
    dados = {}
    
    # 0. Parser da Primeira Linha (Dados Principais: ID, Contrato, Nome)
    # O split consome o cabeçalho "Código Contrato Nome...", restando apenas os valores na primeira linha.
    # Ex: "1 1 ELZA MATOS LIMA"
    primeira_linha_match = re.search(r'^\s*(\d+)\s+(\d+)\s+(.+?)(\n|$)', texto_funcionario.strip())
    if primeira_linha_match:
        dados['ID'] = primeira_linha_match.group(1)
        # O segundo grupo é o Contrato, se precisar: dados['Contrato'] = primeira_linha_match.group(2)
        dados['Nome'] = primeira_linha_match.group(3).strip()
    
    # =========================================================================
    # ESTRATÉGIA HÍBRIDA: Parsers de Linha Específicos + Busca Genérica
    # =========================================================================

    # 1. Identificar linhas complexas (vários campos na mesma linha)
    # Procuramos o TEXTO DA LINHA DE VALORES baseado no cabeçalho imediatamente anterior.
    
    # =========================================================================
    # LISTA DE TODOS OS LABELS CONHECIDOS (Para evitar "merging" de campos)
    # =========================================================================
    # Se o robô encontrar um desses termos na linha de valor, ele para.
    labels_conhecidos = [
        "Nome do", "Data de", "Raça", "Cor", "Sexo", "Deficiente", "Tipo de", "Sanguíneo",
        "Estado", "Nacionalidade", "Chegada", "CPF", "Cédula", "Emissão", "Órgão", "Habilitação",
        "CTPS", "Série", "Dígito", "Carteira", "Conta", "Zona", "PIS", "Civil", "Endereço",
        "Bairro", "Cidade", "CEP", "Telefone", "Celular", "Admissão", "Função", "CBO",
        "Salário", "Forma", "Pagamento", "Categoria", "Matrícula", "Sindicato", "Centro",
        "Localização", "Filiação", "Nascimento", "Fotografia", "Naturalidade", "Plano",
        "Empresa", "Horário", "Rescisão", "Aviso", "Saldo", "Maior", "Recolheu", "Causa",
        "Assinatura", "CNPJ", "FGTS", "Eleitor", "Seção", "Grau", "Instrução", "Cadastramento",
        "Optante", "Banco", "Eletrônico", "Registro", "Insalubridade", "Periculosidade", "Comissão"
    ]

    # ... (Parsers específicos anteriores mantidos: CPF/RG line, Admissão line, City/State line) ...
    # (Mantemos os blocos if re.search(...) anteriores)

    # =========================================================================
    # ESTRATÉGIA HÍBRIDA: Parsers de Linha Específicos + Busca Genérica
    # =========================================================================
    
    # 1. TRINCAS E DUPLAS DE DADOS (Layouts em Colunas Mistas)
    
    # 1.1 Nascimento + Cor + Sexo (Ex: "18/12/1958 Branco Feminino")
    # Tenta encontrar esse padrão específico de data, texto e gênero
    match_trinca = re.search(r'(?P<Nasc>\d{2}/\d{2}/\d{4})\s+(?P<Cor>[A-Za-zÀ-ÿ]+)\s+(?P<Sexo>Masculino|Feminino)', texto_funcionario)
    if match_trinca:
        dados['Data_Nascimento'] = match_trinca.group('Nasc')
        dados['Raca_Cor'] = match_trinca.group('Cor')
        dados['Sexo'] = match_trinca.group('Sexo')

    # 1.2 Data Pis/Cadastro + Estado Civil (Ex: "08/10/1999 Casado")
    match_dupla_civil = re.search(r'(?P<Data>\d{2}/\d{2}/\d{4})\s+(?P<Civil>Solteiro|Casado|Divorciado|Viúvo|Separado|União Estável)', texto_funcionario)
    if match_dupla_civil:
        # Verifica se está perto de "Data de cadastramento" ou "PIS" se possível, mas o padrão é forte.
        # Assumindo que essa data é o Cadastro do PIS
        # dados['Data_Cadastramento'] = match_dupla_civil.group('Data') 
        dados['Estado_Civil'] = match_dupla_civil.group('Civil')

    # 1.3 CPF + RG + Órgão/UF + Data Emissão (Colunas)
    # Tenta capturar linha completa: 123.456.789-00  MG-12.345.678  SSP/MG  01/01/2000
    if re.search(r'CPF.*Cédula de identidade', texto_funcionario):
        # Regex mais permissivo para pegar CPF, RG (qualquer formato), Opcional Orgao, Data
        match_docs = re.search(r'(?P<CPF>\d{3}\.\d{3}\.\d{3}-\d{2})\s+(?P<RG>[^\s]+)\s+(?:(?P<Orgao>[A-Za-z]+/[A-Z]{2})\s+)?(?P<Data>\d{2}/\d{2}/\d{4})', texto_funcionario)
        if match_docs:
            dados['CPF'] = match_docs.group('CPF')
            dados['RG'] = match_docs.group('RG')
            if match_docs.group('Orgao'):
                dados['Orgao_Expedidor'] = match_docs.group('Orgao')
            dados['Data_Emissao_RG'] = match_docs.group('Data')

    # 1.4 Admissão + Função + CBO
    if re.search(r'Data de admissão.*Função.*CBO', texto_funcionario):
         match_adm = re.search(r'(?P<Data>\d{2}/\d{2}/\d{4})\s+(?P<Func>.+?)\s+(?P<CBO>\d{4}-\d{2})', texto_funcionario)
         if match_adm:
             dados['Data_Admissao'] = match_adm.group('Data')
             dados['Funcao'] = match_adm.group('Func').strip()
             dados['CBO'] = match_adm.group('CBO')
         else:
             # Fallback: Só Data e Função
             match_adm_b = re.search(r'(?P<Data>\d{2}/\d{2}/\d{4})\s+(?P<Func>.+)', texto_funcionario)
             if match_adm_b:
                 dados['Data_Admissao'] = match_adm_b.group('Data')
                 dados['Funcao'] = match_adm_b.group('Func').strip()

    # 1.5 Data Rescisão (Simplificado e Robusto)
    # Procura por qualquer variante de "rescisão" seguida de uma data
    # Isso cobre: "Data rescisão", "Data de rescisão", etc.
    if 'Data_Rescisao' not in dados:
        # Primeiro tenta encontrar o padrão com o label explícito
        match_resc = re.search(r'(?i)(?:Data\s+(?:de\s+)?rescis[sç]ão)[:\s]*(\d{2}/\d{2}/\d{4})', texto_funcionario)
        if match_resc:
            dados['Data_Rescisao'] = match_resc.group(1)
        else:
            # Se não encontrou, tenta buscar em contexto de linha com "rescisão" e data na linha seguinte
            match_resc_ctx = re.search(r'(?i)rescis[sç]ão.*?\n\s*(\d{2}/\d{2}/\d{4})', texto_funcionario, re.DOTALL)
            if match_resc_ctx:
                dados['Data_Rescisao'] = match_resc_ctx.group(1)

    # 2. ENDEREÇOS E CONTATOS (Complexidade de Colunas)
    
    # 2.1 Cidade + CEP + Telefone (SEM ESTADO - Caso observado na imagem 7)
    # Ex: "Campinas 13060-518 (19) -"
    # Regex que pega texto, cep e resto
    if re.search(r'Cidade.*CEP.*Telefone', texto_funcionario) and not re.search(r'Cidade.*Estado.*CEP', texto_funcionario):
        match_end_short = re.search(r'(?P<Cidade>[A-Za-zÀ-ÿ\s]+?)\s+(?P<CEP>\d{5}-?\d{3})\s+(?P<Tel>.+)', texto_funcionario)
        if match_end_short:
             dados['Cidade'] = match_end_short.group('Cidade').strip()
             dados['CEP'] = match_end_short.group('CEP')
             dados['Telefone'] = match_end_short.group('Tel').strip()

    # 2.2 Cidade + Estado + CEP + Telefone (Com Estado de 2 letras)
    elif re.search(r'Cidade.*Estado.*CEP.*Telefone', texto_funcionario):
        match_end_full = re.search(r'(?P<Cidade>.+?)\s+(?P<UF>[A-Z]{2})\s+(?P<CEP>\d{5}-?\d{3})\s+(?P<Tel>.+)', texto_funcionario)
        if match_end_full:
            dados['Cidade'] = match_end_full.group('Cidade').strip()
            dados['Estado'] = match_end_full.group('UF')
            dados['CEP'] = match_end_full.group('CEP')
            dados['Telefone'] = match_end_full.group('Tel').strip()

    # 2.3 Endereço + Bairro (Tentativa de Split por Espaço Duplo)
    # Se detectar cabeçalho "Endereço   Bairro", tenta pegar a linha seguinte e dividir
    match_header_end = re.search(r'Endereço\s+Bairro', texto_funcionario)
    if match_header_end and "Endereco" not in dados:
        # Pega a parte do texto APÓS esse cabeçalho
        resto_end = texto_funcionario[match_header_end.end():].strip()
        primeira_linha = resto_end.split('\n')[0].strip()
        # Tenta dividir por 2 ou mais espaços (coluna visual)
        partes = re.split(r'\s{2,}', primeira_linha)
        if len(partes) >= 2:
            dados['Endereco'] = partes[0].strip()
            dados['Bairro'] = partes[1].strip()
        else:
            # Se não conseguiu dividir, joga tudo em Endereço (melhor que duplicar)
            dados['Endereco'] = primeira_linha

    # Salário
    match_salario = re.search(r'R\$\s*([\d\.,]+)', texto_funcionario)
    if match_salario:
        dados['Salario'] = match_salario.group(1)

    # 3. Busca Genérica Inteligente
    campos_simples = {
        r"Nome do pai": "Nome_Pai",
        r"Nome da mãe": "Nome_Mae",
        r"Data de nascimento": "Data_Nascimento",  # Backup caso regex trinca falhe
        r"Raça/cor": "Raca_Cor",                   # Backup
        r"Sexo": "Sexo",                           # Backup
        r"Deficiente": "Deficiente",
        r"Tipo de deficiência": "Tipo_Deficiencia",
        r"Tipo sanguíneo": "Tipo_Sanguineo",
        r"Estado Civil": "Estado_Civil",           # Ajustado label
        r"Nacionalidade": "Nacionalidade",
        r"Naturalidade": "Naturalidade",           # NOVO
        r"Data rescisão": "Data_Rescisao",         # NOVO
        r"Data de rescisão": "Data_Rescisao",      # Variação
        r"Chegada ao Brasil": "Chegada_Brasil",
        r"CTPS": "CTPS",
        r"Série": "Serie_CTPS",
        r"Dígito": "Digito_CTPS",
        r"Carteira reservista": "Reservista",
        r"Zona": "Zona_Eleitoral",
        r"Seção": "Secao_Eleitoral",
        r"Nº título de eleitor": "Titulo_Eleitor",
        r"Nº do PIS": "PIS",
        r"Endereço": "Endereco",                   # Backup
        r"Bairro": "Bairro",                       # Backup
        r"Celular": "Celular",
        r"Matricula eSocial": "Matricula_eSocial",
        r"Sindicato": "Sindicato",
        r"Horário": "Horario",
        r"Data de opção": "Data_Opcao_FGTS",
        r"Centro de custo": "Centro_Custo",
        r"Localização": "Localizacao",
        r"Endereço eletrônico": "Email",
        r"Data do registro": "Data_Registro",
        r"Grau de instrução": "Grau_Instrucao",
        r"Nº da conta FGTS": "Conta_FGTS",
        r"Banco depositário": "Banco_FGTS",
        r"CNPJ": "CNPJ_Empregador"
    }

    for label, chave in campos_simples.items():
        if chave not in dados: 
            match = re.search(rf'(?i){label}', texto_funcionario)
            if match:
                inicio = match.end()
                resto = texto_funcionario[inicio:].strip()
                linhas = resto.split('\n')
                
                for linha in linhas:
                    linha = linha.strip()
                    if not linha or len(linha) < 2 or linha == ":":
                        continue
                        
                    # VERIFICAÇÃO DE COERÊNCIA (ANTI-MERGING)
                    # Verifica se a linha começa com QUALQUER label conhecido
                    # Se sim, significa que o campo atual está VAZIO e já estamos no próximo.
                    eh_label = False
                    for termo in labels_conhecidos:
                        # Verifica se o termo está no início da linha (com alguma folga)
                        if re.match(rf'(?i)^.*{termo}', linha[:20]): 
                            eh_label = True
                            break
                    
                    if not eh_label:
                        dados[chave] = linha
                        break # Achou valor válido
                    else:
                        break # Encontrada label, aborta (campo vazio)

    # 4. Resgate do ID (Código) - Prioridade Máxima
    if "ID" not in dados:
        # Tenta pegar logo no início do texto (padrão mais comum se o split funcionou)
        match_id = re.search(r'^\s*(\d+)', texto_funcionario.strip())
        if match_id:
            dados['ID'] = match_id.group(1)
        else:
             # Fallback
             match_cod = re.search(r'(?i)C[óoÕó]digo\s*\n?\s*(\d+)', texto_funcionario)
             if match_cod:
                 dados['ID'] = match_cod.group(1)

    return dados

def remover_cabecalho(texto):
    # Solução profissional (Adaptada para Tabela):
    # Aceita acentos em Código/Codigo e ignora o texto do meio (Contrato Nome...) até achar o número
    match = re.search(r'(?i)C[óoÕó]digo.*?\n?\s*\d+', texto)
    if match:
        return texto[match.start():]
    return texto

def separar_funcionarios(texto):
    # Baseado no dump: "Código Contrato Nome do(a) trabalhador(a)"
    # O separador É o cabeçalho da tabela.
    # Ex: "Código Contrato Nome do(a) trabalhador(a)"
    # Isso garante que pegamos o início de cada ficha.
    padrao = r'(?i)C[óoÕó]digo\s+Contrato\s+Nome.*?(?:\n|$)'
    
    blocos = re.split(padrao, texto)

    funcionarios = []
    # O primeiro bloco serÃ¡ o texto anterior ao primeiro cabeçalho (ou vazio), ignoramos
    if len(blocos) > 1:
        for bloco in blocos[1:]:
            if bloco.strip():
                 funcionarios.append(bloco)
    
    # Fallback: Se não funcionou o split novo, tenta o antigo (para outros layouts)
    if not funcionarios:
       # Tenta procurar apenas por Código seguido de número (caso vertical antigo)
       # Mas agora permitindo quebra de linha agressiva
       padrao_fallback = r'(?i)Código\s*\n?\s*\d+'
       ids = re.findall(padrao_fallback, texto)
       blocos_fb = re.split(padrao_fallback, texto)
       if len(blocos_fb) > len(ids):
           for i in range(len(ids)):
               if i < len(blocos_fb[1:]):
                   funcionarios.append(ids[i] + "\n" + blocos_fb[i+1])
                   
    return funcionarios 

    for i in range(len(ids)):
        if i < len(blocos):
            bloco = ids[i] + "\n" + blocos[i]
            funcionarios.append(bloco)

    return funcionarios

# ==============================
# CLASSE DA INTERFACE
# ==============================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.VERSION = "v1.3 (Debug)"
        self.title(f"Extrator de Ficha de Registro - Premium {self.VERSION}")
        self.geometry("600x600") # Aumentei altura para caber o log
        self.resizable(True, True) # Permitir redimensionar para ver logs
        self.create_widgets()

    def create_widgets(self):
        self.label_title = ctk.CTkLabel(self, text=f"Extrator de PDF {self.VERSION}", font=ctk.CTkFont(size=24, weight="bold"))
        self.label_title.pack(pady=(20, 5))

        self.label_subtitle = ctk.CTkLabel(self, text="Modo Debug Ativado", text_color="#FFAA00", font=ctk.CTkFont(size=12, weight="bold"))
        self.label_subtitle.pack(pady=(0, 10))

        self.frame_options = ctk.CTkFrame(self)
        self.frame_options.pack(pady=5, padx=20, fill="x")

        self.label_format = ctk.CTkLabel(self.frame_options, text="Formato de Saída:", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_format.pack(pady=(10, 5))

        self.formato_var = ctk.StringVar(value="Excel")
        
        self.radio_excel = ctk.CTkRadioButton(self.frame_options, text="Excel (.xlsx)", variable=self.formato_var, value="Excel")
        self.radio_excel.pack(pady=5)
        
        self.radio_csv = ctk.CTkRadioButton(self.frame_options, text="CSV (;)", variable=self.formato_var, value="CSV")
        self.radio_csv.pack(pady=5)
        
        self.radio_txt = ctk.CTkRadioButton(self.frame_options, text="TXT (|)", variable=self.formato_var, value="TXT")
        self.radio_txt.pack(pady=(5, 10))

        self.btn_action = ctk.CTkButton(
            self, 
            text="Selecionar PDF e Iniciar", 
            command=self.iniciar_thread,
            height=40,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.btn_action.pack(pady=15, padx=20, fill="x")

        self.progress_bar = ctk.CTkProgressBar(self, mode="determinate")
        self.progress_bar.pack(pady=5, padx=20, fill="x")
        self.progress_bar.set(0)

        self.label_status = ctk.CTkLabel(self, text="Aguardando início...", text_color="gray")
        self.label_status.pack(pady=(0, 5))
        
        # LOG CONSOLE
        self.label_log = ctk.CTkLabel(self, text="Console de Execução (Bastidores):", font=ctk.CTkFont(size=12, weight="bold"))
        self.label_log.pack(pady=(5, 0), padx=20, anchor="w")
        
        self.log_box = ctk.CTkTextbox(self, height=150, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_box.pack(pady=5, padx=20, fill="both", expand=True)
        self.log_box.configure(state="disabled")

    def log(self, mensagem):
        timestamp = time.strftime("%H:%M:%S")
        texto_log = f"[{timestamp}] {mensagem}\n"
        
        # Atualização segura da UI
        self.log_box.configure(state="normal")
        self.log_box.insert("end", texto_log)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        print(texto_log.strip()) # Também imprime no terminal
        self.update_idletasks()

    def iniciar_thread(self):
        self.btn_action.configure(state="disabled")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        
        thread = threading.Thread(target=self.executar_processo)
        thread.start()

    def update_status(self, message, progress=None):
        self.label_status.configure(text=message)
        if progress is not None:
            self.progress_bar.set(progress)
        # self.update_idletasks() # Removido para evitar lag excessivo

    def executar_processo(self):
        try:
            pdf_path = filedialog.askopenfilename(title="Selecionar PDF", filetypes=[("Arquivos PDF", "*.pdf")])
            if not pdf_path:
                self.btn_action.configure(state="normal")
                self.log("Seleção de arquivo cancelada.")
                return

            self.log(f"Arquivo selecionado: {os.path.basename(pdf_path)}")
            self.log("Iniciando leitura do PDF...")
            self.update_status("Lendo PDF...", 0.05)

            df = self.processar_pdf(pdf_path)

            if df is None or df.empty:
                 self.btn_action.configure(state="normal")
                 self.log("Nenhum dado foi extraído ou ocorreu erro fatal.")
                 return

            self.log("Preparando para salvar...")
            self.update_status("Salvando arquivo...", 0.95)
            
            formato = self.formato_var.get()
            extensao = {"Excel": ".xlsx", "CSV": ".csv", "TXT": ".txt"}[formato]

            save_path = filedialog.asksaveasfilename(defaultextension=extensao, filetypes=[("Arquivo", "*" + extensao)])

            if not save_path:
                 self.log("Salvamento cancelado pelo usuário.")
                 self.btn_action.configure(state="normal")
                 return

            self.exportar(df, save_path, formato)
            
            self.log(f"Arquivo salvo com sucesso em: {save_path}")
            self.update_status("Concluído!", 1.0)
            messagebox.showinfo("Sucesso", "Processso finalizado com sucesso!")

        except Exception as e:
            self.log(f"ERRO CRÍTICO NO PROCESSO PRINCIPAL: {e}")
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro:\n{str(e)}")
        
        finally:
            self.btn_action.configure(state="normal")

    def processar_pdf(self, pdf_path):
        texto_completo = ""
        tamanho_total = 0
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_paginas = len(pdf.pages)
                self.log(f"PDF aberto. Total de páginas: {total_paginas}")
                
                if total_paginas == 0:
                    self.log("ERRO: PDF vazio.")
                    return None

                for i, pagina in enumerate(pdf.pages):
                    try:
                        # Log detalhado a cada 10 páginas para não poluir
                        if i % 10 == 0:
                            self.log(f"Lendo página {i+1}/{total_paginas}...")
                        
                        texto = pagina.extract_text()
                        if texto:
                            texto_completo += texto + "\n"
                            tamanho_total += len(texto)
                        
                        progresso = 0.05 + (0.35 * ((i + 1) / total_paginas))
                        self.update_status(f"Lendo página {i+1}...", progresso)
                        
                    except Exception as e:
                        self.log(f"ERRO ao ler página {i+1}: {e}")
                        continue

            self.log(f"Leitura concluída. Tamanho total do texto extraído: {tamanho_total} caracteres.")
            self.log(f"Memória aproximada do texto: {tamanho_total / 1024 / 1024:.2f} MB")

        except Exception as e:
            self.log(f"ERRO FATAL ao abrir PDF: {e}")
            return None

        self.log("Iniciando tratamento do texto (limpeza de cabeçalho)...")
        self.update_status("Tratando texto...", 0.45)
        
        t0 = time.time()
        texto_tratado = remover_cabecalho(texto_completo)
        t1 = time.time()
        self.log(f"Cabeçalho removido em {t1-t0:.2f} segundos.")

        self.log("Iniciando separação de funcionários (Regex Split)...")
        self.update_status("Separando registros...", 0.5)
        
        t0 = time.time()
        funcionarios = separar_funcionarios(texto_tratado)
        t1 = time.time()
        
        total_funcionarios = len(funcionarios)
        self.log(f"Separação concluída em {t1-t0:.2f} segundos. Registros encontrados: {total_funcionarios}")
        
        if total_funcionarios == 0:
            self.log("ALERTA: Nenhum registro de funcionário encontrado (padrão 'Código' não correspondido).")
            return pd.DataFrame()

        lista_dados = []
        
        self.log(f"Iniciando extração paralela (ProcessPoolExecutor) para {total_funcionarios} registros...")
        self.update_status("Extraindo dados (Paralelo)...", 0.55)
        
        # Otimização: Chunksize para envio ao pool
        # Se houver muitos registros pequenos, enviar de 100 em 100 pode ser mais rápido
        
        t0 = time.time()
        with concurrent.futures.ProcessPoolExecutor() as executor:
            future_to_funcionario = {executor.submit(extrair_campos, func): i for i, func in enumerate(funcionarios)}
            
            completed_count = 0
            for future in concurrent.futures.as_completed(future_to_funcionario):
                try:
                    data = future.result()
                    lista_dados.append(data)
                    
                    completed_count += 1
                    progresso = 0.55 + (0.4 * (completed_count / total_funcionarios))
                    
                    if completed_count % 50 == 0 or completed_count == total_funcionarios:
                        self.log(f"Processado: {completed_count}/{total_funcionarios} registros...")
                        self.update_status(f"Extraindo: {completed_count}/{total_funcionarios}", progresso)
                        
                except Exception as exc:
                    idx = future_to_funcionario[future]
                    self.log(f"ERRO no registro {idx}: {exc}")

        t1 = time.time()
        self.log(f"Extração paralela finalizada em {t1-t0:.2f} segundos.")
        
        return pd.DataFrame(lista_dados)

    def exportar(self, df, caminho, formato):
        if formato == "Excel":
            df.to_excel(caminho, index=False)
        elif formato == "CSV":
            df.to_csv(caminho, index=False, sep=";", encoding="utf-8-sig")
        elif formato == "TXT":
            df.to_csv(caminho, index=False, sep="|", encoding="utf-8")

if __name__ == "__main__":
    # Importante para Windows: proteção do entry point
    app = App()
    app.mainloop()
