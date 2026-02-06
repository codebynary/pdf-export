"""
Extrator de Fichas Word - Interface Gráfica Profissional
Versão 1.0.0

Aplicação GUI moderna para extração em lote de dados de fichas de registro
em formato Word (.docx) com exportação para Excel.

Autor: Antigravity AI
Data: 2026-02-06
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from pathlib import Path
from datetime import datetime
import docx
import pandas as pd
from typing import Dict, List


class ExtratorWordGUI:
    """Interface gráfica profissional para extração de dados de Word"""
    
    VERSION = "1.0.0"
    
    def __init__(self, root):
        self.root = root
        self.root.title(f"Extrator de Fichas Word v{self.VERSION}")
        self.root.geometry("750x550")
        self.root.resizable(True, True)
        
        # Variáveis
        self.diretorio_selecionado = tk.StringVar()
        self.total_arquivos = 0
        self.arquivos_processados = 0
        self.processando = False
        
        # Configurar estilo
        self.configurar_estilo()
        
        # Criar interface
        self.criar_interface()
        
        # Centralizar janela
        self.centralizar_janela()
    
    def configurar_estilo(self):
        """Configura o estilo visual moderno com tema escuro"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Cores do tema escuro moderno
        self.cor_fundo = "#1e1e2e"         # Fundo escuro principal
        self.cor_fundo_sec = "#2a2a3e"     # Fundo secundário
        self.cor_primaria = "#89b4fa"      # Azul claro
        self.cor_secundaria = "#74c7ec"    # Azul água
        self.cor_sucesso = "#a6e3a1"       # Verde claro
        self.cor_aviso = "#f9e2af"         # Amarelo claro
        self.cor_erro = "#f38ba8"          # Vermelho claro
        self.cor_texto = "#cdd6f4"         # Texto claro
        self.cor_texto_sec = "#a6adc8"     # Texto secundário
        self.cor_borda = "#45475a"         # Borda
        self.cor_botao = "#313244"         # Fundo botão
        self.cor_botao_hover = "#45475a"   # Hover botão
        
        # Configurar estilos de labels
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 14, 'bold'),
                       foreground=self.cor_primaria,
                       background=self.cor_fundo)
        
        style.configure('Subtitle.TLabel',
                       font=('Segoe UI', 9),
                       foreground=self.cor_texto_sec,
                       background=self.cor_fundo)
        
        style.configure('TLabel',
                       foreground=self.cor_texto,
                       background=self.cor_fundo)
        
        style.configure('TFrame',
                       background=self.cor_fundo)
        
        style.configure('TLabelframe',
                       background=self.cor_fundo,
                       foreground=self.cor_texto,
                       bordercolor=self.cor_borda)
        
        style.configure('TLabelframe.Label',
                       background=self.cor_fundo,
                       foreground=self.cor_primaria,
                       font=('Segoe UI', 10, 'bold'))
        
        # Entry
        style.configure('TEntry',
                       fieldbackground=self.cor_fundo_sec,
                       foreground=self.cor_texto,
                       bordercolor=self.cor_borda)
        
        # Progressbar
        style.configure('TProgressbar',
                       background=self.cor_primaria,
                       troughcolor=self.cor_fundo_sec,
                       bordercolor=self.cor_borda,
                       lightcolor=self.cor_primaria,
                       darkcolor=self.cor_primaria)
        
        # Configurar cores do root
        self.root.configure(bg=self.cor_fundo)
    
    def criar_interface(self):
        """Cria todos os componentes da interface"""
        
        # Frame principal com padding reduzido
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.configure(style='TFrame')
        
        # === CABEÇALHO ===
        self.criar_cabecalho(main_frame)
        
        # === SELEÇÃO DE DIRETÓRIO ===
        self.criar_secao_diretorio(main_frame)
        
        # === BARRA DE PROGRESSO ===
        self.criar_secao_progresso(main_frame)
        
        # === LOG DE ATIVIDADES ===
        self.criar_secao_log(main_frame)
        
        # === BOTÕES DE AÇÃO ===
        self.criar_secao_botoes(main_frame)
        
        # === RODAPÉ ===
        self.criar_rodape(main_frame)
    
    def criar_cabecalho(self, parent):
        """Cria o cabeçalho com título e versão"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Título
        title_label = ttk.Label(header_frame,
                               text="📄 Extrator de Fichas Word",
                               style='Title.TLabel')
        title_label.pack()
        
        # Subtítulo
        subtitle_label = ttk.Label(header_frame,
                                   text="Extração em lote de dados de fichas de registro para Excel",
                                   style='Subtitle.TLabel')
        subtitle_label.pack()
        
        # Versão
        version_label = ttk.Label(header_frame,
                                 text=f"v{self.VERSION}",
                                 font=('Segoe UI', 8),
                                 foreground='#64748b',
                                 background=self.cor_fundo)
        version_label.pack(pady=(5, 0))
    
    def criar_secao_diretorio(self, parent):
        """Cria a seção de seleção de diretório"""
        dir_frame = ttk.LabelFrame(parent, text="📁 Diretório de Origem", padding="8")
        dir_frame.pack(fill=tk.X, pady=(0, 8))
        
        # Entry para mostrar caminho
        self.dir_entry = ttk.Entry(dir_frame,
                                   textvariable=self.diretorio_selecionado,
                                   state='readonly',
                                   font=('Segoe UI', 8))
        self.dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Botão selecionar
        btn_selecionar = tk.Button(dir_frame,
                                  text="📁 Selecionar",
                                  command=self.selecionar_diretorio,
                                  bg=self.cor_botao,
                                  fg=self.cor_texto,
                                  font=('Segoe UI', 9, 'bold'),
                                  relief=tk.FLAT,
                                  padx=12,
                                  pady=6,
                                  cursor='hand2',
                                  activebackground=self.cor_botao_hover,
                                  activeforeground=self.cor_texto)
        btn_selecionar.pack(side=tk.RIGHT)
        
        # Hover effect
        btn_selecionar.bind('<Enter>', lambda e: btn_selecionar.config(bg=self.cor_botao_hover))
        btn_selecionar.bind('<Leave>', lambda e: btn_selecionar.config(bg=self.cor_botao))
    
    def criar_secao_progresso(self, parent):
        """Cria a seção de progresso"""
        progress_frame = ttk.LabelFrame(parent, text="📊 Progresso", padding="8")
        progress_frame.pack(fill=tk.X, pady=(0, 8))
        
        # Label de status
        self.status_label = ttk.Label(progress_frame,
                                     text="Aguardando seleção de diretório...",
                                     font=('Segoe UI', 8, 'bold'),
                                     foreground=self.cor_texto,
                                     background=self.cor_fundo)
        self.status_label.pack(fill=tk.X, pady=(0, 5))
        
        # Barra de progresso
        self.progress_bar = ttk.Progressbar(progress_frame,
                                           mode='determinate',
                                           length=300)
        self.progress_bar.pack(fill=tk.X, pady=(0, 3))
        
        # Label de porcentagem
        self.percent_label = ttk.Label(progress_frame,
                                      text="0%",
                                      font=('Segoe UI', 8),
                                      foreground=self.cor_texto,
                                      background=self.cor_fundo)
        self.percent_label.pack()
    
    def criar_secao_log(self, parent):
        """Cria a seção de log de atividades"""
        log_frame = ttk.LabelFrame(parent, text="📋 Log", padding="8")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        # Text widget com scroll
        self.log_text = scrolledtext.ScrolledText(log_frame,
                                                 height=10,
                                                 font=('Consolas', 8),
                                                 bg=self.cor_fundo_sec,
                                                 fg=self.cor_texto,
                                                 relief=tk.FLAT,
                                                 borderwidth=1,
                                                 insertbackground=self.cor_texto)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Configurar tags para cores (tema escuro)
        self.log_text.tag_config('info', foreground=self.cor_secundaria)
        self.log_text.tag_config('success', foreground=self.cor_sucesso)
        self.log_text.tag_config('warning', foreground=self.cor_aviso)
        self.log_text.tag_config('error', foreground=self.cor_erro)
        self.log_text.tag_config('header', foreground=self.cor_primaria, font=('Consolas', 9, 'bold'))
    
    def criar_secao_botoes(self, parent):
        """Cria a seção de botões de ação"""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Botão Processar
        self.btn_processar = tk.Button(btn_frame,
                                      text="▶ Processar",
                                      command=self.iniciar_processamento,
                                      bg=self.cor_primaria,
                                      fg="#1e1e2e",
                                      font=('Segoe UI', 9, 'bold'),
                                      relief=tk.FLAT,
                                      padx=15,
                                      pady=6,
                                      cursor='hand2',
                                      activebackground=self.cor_secundaria,
                                      activeforeground="#1e1e2e")
        self.btn_processar.pack(side=tk.LEFT, padx=(0, 10))
        
        # Hover effects para processar
        self.btn_processar.bind('<Enter>', lambda e: self.btn_processar.config(bg=self.cor_secundaria) if self.btn_processar['state'] != 'disabled' else None)
        self.btn_processar.bind('<Leave>', lambda e: self.btn_processar.config(bg=self.cor_primaria) if self.btn_processar['state'] != 'disabled' else None)
        
        # Botão Limpar Log
        btn_limpar = tk.Button(btn_frame,
                              text="🗑️ Limpar",
                              command=self.limpar_log,
                              bg=self.cor_botao,
                              fg=self.cor_texto,
                              font=('Segoe UI', 9, 'bold'),
                              relief=tk.FLAT,
                              padx=12,
                              pady=6,
                              cursor='hand2',
                              activebackground=self.cor_botao_hover,
                              activeforeground=self.cor_texto)
        btn_limpar.pack(side=tk.LEFT, padx=(0, 10))
        
        # Hover effects
        btn_limpar.bind('<Enter>', lambda e: btn_limpar.config(bg=self.cor_botao_hover))
        btn_limpar.bind('<Leave>', lambda e: btn_limpar.config(bg=self.cor_botao))
        
        # Botão Sair
        btn_sair = tk.Button(btn_frame,
                            text="✖ Sair",
                            command=self.sair,
                            bg=self.cor_erro,
                            fg="#1e1e2e",
                            font=('Segoe UI', 9, 'bold'),
                            relief=tk.FLAT,
                            padx=12,
                            pady=6,
                            cursor='hand2',
                            activebackground="#eba0ac",
                            activeforeground="#1e1e2e")
        btn_sair.pack(side=tk.RIGHT)
        
        # Hover effects
        btn_sair.bind('<Enter>', lambda e: btn_sair.config(bg="#eba0ac"))
        btn_sair.bind('<Leave>', lambda e: btn_sair.config(bg=self.cor_erro))
    
    def criar_rodape(self, parent):
        """Cria o rodapé com informações"""
        footer_frame = ttk.Frame(parent)
        footer_frame.pack(fill=tk.X)
        
        footer_label = ttk.Label(footer_frame,
                                text="Desenvolvido com ❤️ por Antigravity AI",
                                font=('Segoe UI', 8),
                                foreground='#94a3b8',
                                background=self.cor_fundo)
        footer_label.pack()
    
    def centralizar_janela(self):
        """Centraliza a janela na tela"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def adicionar_log(self, mensagem, tag='info'):
        """Adiciona mensagem ao log com timestamp"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert(tk.END, f"[{timestamp}] ", 'header')
        self.log_text.insert(tk.END, f"{mensagem}\n", tag)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def limpar_log(self):
        """Limpa o log de atividades"""
        self.log_text.delete(1.0, tk.END)
        self.adicionar_log("Log limpo.", 'info')
    
    def selecionar_diretorio(self):
        """Abre diálogo para selecionar diretório"""
        diretorio = filedialog.askdirectory(title="Selecione o diretório com os arquivos .docx")
        
        if diretorio:
            self.diretorio_selecionado.set(diretorio)
            self.adicionar_log(f"Diretório selecionado: {diretorio}", 'success')
            
            # Contar arquivos .docx
            arquivos = list(Path(diretorio).glob('*.docx'))
            arquivos = [f for f in arquivos if not f.name.startswith('~$')]
            self.total_arquivos = len(arquivos)
            
            if self.total_arquivos > 0:
                self.adicionar_log(f"Encontrados {self.total_arquivos} arquivo(s) .docx", 'info')
                self.btn_processar.config(state='normal')
                self.status_label.config(text=f"Pronto para processar {self.total_arquivos} arquivo(s)")
            else:
                self.adicionar_log("Nenhum arquivo .docx encontrado neste diretório", 'warning')
                self.btn_processar.config(state='disabled')
                self.status_label.config(text="Nenhum arquivo encontrado")
    
    def iniciar_processamento(self):
        """Inicia o processamento em thread separada"""
        if not self.diretorio_selecionado.get():
            messagebox.showwarning("Aviso", "Selecione um diretório primeiro!")
            return
        
        # Desabilitar botão
        self.btn_processar.config(state='disabled')
        self.processando = True
        
        # Iniciar thread
        thread = threading.Thread(target=self.processar_arquivos, daemon=True)
        thread.start()
    
    def processar_arquivos(self):
        """Processa todos os arquivos .docx do diretório"""
        try:
            self.adicionar_log("="*60, 'header')
            self.adicionar_log("Iniciando processamento...", 'header')
            self.adicionar_log("="*60, 'header')
            
            diretorio = self.diretorio_selecionado.get()
            
            # Buscar arquivos
            arquivos = list(Path(diretorio).glob('*.docx'))
            arquivos = [f for f in arquivos if not f.name.startswith('~$')]
            
            self.total_arquivos = len(arquivos)
            self.arquivos_processados = 0
            
            # Processar cada arquivo
            resultados = []
            
            for i, arquivo in enumerate(arquivos, 1):
                self.status_label.config(text=f"Processando: {arquivo.name}")
                self.adicionar_log(f"[{i}/{self.total_arquivos}] {arquivo.name}", 'info')
                
                try:
                    dados = self.extrair_documento(str(arquivo))
                    resultados.append(dados)
                    self.adicionar_log(f"  ✓ Extraído com sucesso", 'success')
                except Exception as e:
                    self.adicionar_log(f"  ✗ Erro: {str(e)}", 'error')
                    resultados.append({'arquivo_origem': arquivo.name, 'erro': str(e)})
                
                # Atualizar progresso
                self.arquivos_processados = i
                progresso = (i / self.total_arquivos) * 100
                self.progress_bar['value'] = progresso
                self.percent_label.config(text=f"{progresso:.1f}%")
                self.root.update_idletasks()
            
            # Exportar para Excel
            self.adicionar_log("="*60, 'header')
            self.adicionar_log("Gerando planilha Excel...", 'info')
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            arquivo_saida = os.path.join(diretorio, f'fichas_extraidas_{timestamp}.xlsx')
            
            self.exportar_para_excel(resultados, arquivo_saida)
            
            self.adicionar_log(f"✓ Planilha salva: {os.path.basename(arquivo_saida)}", 'success')
            self.adicionar_log("="*60, 'header')
            self.adicionar_log("PROCESSAMENTO CONCLUÍDO COM SUCESSO!", 'success')
            self.adicionar_log("="*60, 'header')
            
            self.status_label.config(text=f"Concluído! {self.total_arquivos} arquivo(s) processado(s)")
            
            # Mostrar mensagem de sucesso
            self.root.after(0, lambda: messagebox.showinfo(
                "Sucesso!",
                f"Processamento concluído!\n\n"
                f"Arquivos processados: {self.total_arquivos}\n"
                f"Arquivo gerado: {os.path.basename(arquivo_saida)}"
            ))
            
        except Exception as e:
            self.adicionar_log(f"ERRO CRÍTICO: {str(e)}", 'error')
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Erro durante processamento:\n{str(e)}"))
        
        finally:
            self.processando = False
            self.btn_processar.config(state='normal')
            self.progress_bar['value'] = 0
            self.percent_label.config(text="0%")
    
    def extrair_documento(self, caminho_arquivo: str) -> Dict[str, str]:
        """Extrai dados de um documento Word"""
        campos_mapeamento = {
            'Código': 'codigo', 'Contrato': 'contrato', 'Nome do(a) trabalhador(a)': 'nome',
            'Matricula eSocial': 'matricula_esocial', 'Nome do pai': 'nome_pai',
            'Nome da mãe': 'nome_mae', 'Data de nascimento': 'data_nascimento',
            'Raça/cor': 'raca_cor', 'Sexo': 'sexo', 'Naturalidade': 'naturalidade',
            'Nacionalidade': 'nacionalidade', 'Estado Civil': 'estado_civil',
            'Deficiente': 'deficiente', 'Tipo de deficiência': 'tipo_deficiencia',
            'Tipo sanguíneo': 'tipo_sanguineo', 'CPF': 'cpf',
            'Cédula de identidade': 'rg', 'Data de emissão': 'data_emissao_rg',
            'Órgão/UF': 'orgao_uf_rg', 'CTPS': 'ctps', 'Série': 'serie_ctps',
            'Dígito': 'digito_ctps', 'Nº título de eleitor': 'titulo_eleitor',
            'Zona': 'zona_eleitoral', 'Seção': 'secao_eleitoral', 'Nº do PIS': 'pis',
            'Data de cadastramento': 'data_cadastramento_pis', 'Grau de instrução': 'grau_instrucao',
            'Endereço': 'endereco', 'Número': 'numero', 'Complemento': 'complemento',
            'Bairro': 'bairro', 'Cidade': 'cidade', 'Estado': 'estado', 'CEP': 'cep',
            'Telefone': 'telefone', 'Celular': 'celular', 'Endereço eletrônico': 'email',
            'Data de admissão': 'data_admissao', 'Data do registro': 'data_registro',
            'Função': 'funcao', 'CBO': 'cbo', 'Salário Inicial': 'salario_inicial',
            'Forma de pagamento': 'forma_pagamento', 'Tipo de pagamento': 'tipo_pagamento',
            'Insalubridade': 'insalubridade', 'Periculosidade': 'periculosidade',
            'Sindicato': 'sindicato', 'Centro de custo': 'centro_custo',
            'Localização': 'localizacao', 'Horário': 'horario',
            'Nº da conta FGTS': 'conta_fgts', 'Data de opção': 'data_opcao_fgts',
            'Banco depositário - FGTS': 'banco_fgts', 'Data rescisão': 'data_rescisao',
            'Aviso prévio': 'aviso_previo', 'Saldo FGTS': 'saldo_fgts',
            'Maior remuneração': 'maior_remuneracao', 'Causa da rescisão': 'causa_rescisao',
            'Empregador': 'empregador', 'CNPJ': 'cnpj_empregador'
        }
        
        # Campos de endereço que aparecem duas vezes (empresa e residencial)
        campos_endereco = ['endereco', 'numero', 'complemento', 'bairro', 'cidade', 'estado', 'cep', 'telefone', 'celular']
        
        doc = docx.Document(caminho_arquivo)
        dados = {}
        
        # Processar todas as tabelas
        for tabela in doc.tables:
            for row_idx, row in enumerate(tabela.rows):
                for cell in row.cells:
                    texto_celula = cell.text.strip()
                    
                    # Método 1: Label\nValor (formato padrão)
                    if '\n' in texto_celula:
                        linhas = texto_celula.split('\n')
                        if len(linhas) >= 2:
                            label = linhas[0].strip()
                            valor = '\n'.join(linhas[1:]).strip()
                            
                            if label in campos_mapeamento:
                                campo_chave = campos_mapeamento[label]
                                
                                # Se for campo de endereço
                                if campo_chave in campos_endereco:
                                    # Endereço da empresa aparece nas primeiras linhas (< 10)
                                    # Endereço residencial aparece depois (>= 15)
                                    # Só capturar se estiver na região do endereço residencial
                                    if row_idx >= 15:
                                        if campo_chave not in dados or not dados[campo_chave]:
                                            dados[campo_chave] = valor
                                else:
                                    # Campos não relacionados a endereço: extrair normalmente
                                    if campo_chave not in dados or not dados[campo_chave]:
                                        dados[campo_chave] = valor
                    
                    # Método 2: Procurar labels conhecidos no texto (apenas para campos não-endereço)
                    for label, campo_chave in campos_mapeamento.items():
                        # Pular campos de endereço neste método
                        if campo_chave in campos_endereco:
                            continue
                            
                        if label in texto_celula and campo_chave not in dados:
                            # Tentar extrair valor após o label
                            partes = texto_celula.split(label, 1)
                            if len(partes) == 2:
                                valor = partes[1].strip().strip('\n').strip()
                                if valor and campo_chave not in dados:
                                    dados[campo_chave] = valor
        
        dados['arquivo_origem'] = os.path.basename(caminho_arquivo)
        dados['data_extracao'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        return dados
    
    def exportar_para_excel(self, dados: List[Dict[str, str]], arquivo_saida: str):
        """Exporta dados para Excel"""
        df = pd.DataFrame(dados)
        
        colunas_prioritarias = [
            'arquivo_origem', 'nome', 'cpf', 'rg', 'data_nascimento',
            'data_admissao', 'funcao', 'salario_inicial', 'data_rescisao'
        ]
        
        colunas_ordenadas = [col for col in colunas_prioritarias if col in df.columns]
        colunas_restantes = [col for col in df.columns if col not in colunas_ordenadas]
        colunas_ordenadas.extend(colunas_restantes)
        
        df = df[colunas_ordenadas]
        df.to_excel(arquivo_saida, index=False, engine='openpyxl')
    
    def sair(self):
        """Fecha a aplicação"""
        if self.processando:
            if messagebox.askyesno("Confirmar", "Há um processamento em andamento. Deseja realmente sair?"):
                self.root.destroy()
        else:
            self.root.destroy()


def main():
    """Função principal"""
    root = tk.Tk()
    app = ExtratorWordGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
