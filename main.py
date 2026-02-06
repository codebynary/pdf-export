import customtkinter as ctk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
import os
import threading
import concurrent.futures

# Configuração do tema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("dark-blue")

# ==============================
# FUNÇÕES DE EXTRAÇÃO (Top-level para Multiprocessing)
# ==============================
def extrair_campos(texto_funcionario):
    dados = {}
    padrao = re.findall(r'([A-Za-zÀ-ÿ0-9\s\/\-\(\)\.]+)\s*:\s*(.+)', texto_funcionario)

    for campo, valor in padrao:
        campo_limpo = (
            campo.strip()
            .replace(" ", "_")
            .replace("/", "_")
            .replace("-", "_")
            .replace(".", "")
        )
        dados[campo_limpo] = valor.strip()

    if "ID" not in dados:
        match_codigo = re.search(r'Código\s*\n?\s*(\d+)', texto_funcionario)
        if match_codigo:
            dados['ID'] = match_codigo.group(1)

    return dados

def remover_cabecalho(texto):
    match = re.search(r'Código\s*\n?\s*\d+', texto)
    if match:
        return texto[match.start():]
    return texto

def separar_funcionarios(texto):
    padrao = r'Código\s*\n?\s*\d+'
    ids = re.findall(padrao, texto)
    blocos = re.split(padrao, texto)

    funcionarios = []
    if len(blocos) > len(ids): 
        # Ajuste de listas
        blocos = blocos[1:] 

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

        self.title("Extrator de Ficha de Registro - Premium (Multicore)")
        self.geometry("500x450")
        self.resizable(False, False)
        self.create_widgets()

    def create_widgets(self):
        self.label_title = ctk.CTkLabel(self, text="Extrator de PDF Turbo", font=ctk.CTkFont(size=24, weight="bold"))
        self.label_title.pack(pady=(30, 10))

        self.label_subtitle = ctk.CTkLabel(self, text="Processamento Paralelo Ativado", text_color="#00FF00", font=ctk.CTkFont(size=12, weight="bold"))
        self.label_subtitle.pack(pady=(0, 20))

        self.frame_options = ctk.CTkFrame(self)
        self.frame_options.pack(pady=10, padx=20, fill="x")

        self.label_format = ctk.CTkLabel(self.frame_options, text="Formato de Saída:", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_format.pack(pady=(15, 5))

        self.formato_var = ctk.StringVar(value="Excel")
        
        self.radio_excel = ctk.CTkRadioButton(self.frame_options, text="Excel (.xlsx)", variable=self.formato_var, value="Excel")
        self.radio_excel.pack(pady=5)
        
        self.radio_csv = ctk.CTkRadioButton(self.frame_options, text="CSV (;)", variable=self.formato_var, value="CSV")
        self.radio_csv.pack(pady=5)
        
        self.radio_txt = ctk.CTkRadioButton(self.frame_options, text="TXT (|)", variable=self.formato_var, value="TXT")
        self.radio_txt.pack(pady=(5, 15))

        self.btn_action = ctk.CTkButton(
            self, 
            text="Selecionar PDF e Iniciar", 
            command=self.iniciar_thread,
            height=40,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.btn_action.pack(pady=20, padx=20, fill="x")

        self.progress_bar = ctk.CTkProgressBar(self, mode="determinate")
        self.progress_bar.pack(pady=10, padx=20, fill="x")
        self.progress_bar.set(0)

        self.label_status = ctk.CTkLabel(self, text="Aguardando início...", text_color="gray")
        self.label_status.pack(pady=(0, 20))

    def iniciar_thread(self):
        self.btn_action.configure(state="disabled")
        thread = threading.Thread(target=self.executar_processo)
        thread.start()

    def update_status(self, message, progress=None):
        self.label_status.configure(text=message)
        if progress is not None:
            self.progress_bar.set(progress)
        self.update_idletasks()

    def executar_processo(self):
        try:
            pdf_path = filedialog.askopenfilename(
                title="Selecionar PDF",
                filetypes=[("Arquivos PDF", "*.pdf")]
            )

            if not pdf_path:
                self.btn_action.configure(state="normal")
                self.update_status("Seleção cancelada.")
                return

            self.update_status("Lendo PDF (fase única)...", 0.1)

            df = self.processar_pdf(pdf_path)

            if df is None or df.empty:
                 self.btn_action.configure(state="normal")
                 return

            self.update_status("Salvando arquivo...", 0.95)
            
            formato = self.formato_var.get()
            extensao = {"Excel": ".xlsx", "CSV": ".csv", "TXT": ".txt"}[formato]

            save_path = filedialog.asksaveasfilename(
                defaultextension=extensao,
                filetypes=[("Arquivo", "*" + extensao)]
            )

            if not save_path:
                 self.update_status("Salvamento cancelado.")
                 self.btn_action.configure(state="normal")
                 return

            self.exportar(df, save_path, formato)
            
            self.update_status("Concluído com Sucesso!", 1.0)
            messagebox.showinfo("Sucesso", "Arquivo exportado com sucesso!")

        except Exception as e:
            self.update_status("Erro no processamento.")
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado:\n{str(e)}")
        
        finally:
            self.btn_action.configure(state="normal")

    def processar_pdf(self, pdf_path):
        texto_completo = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_paginas = len(pdf.pages)
                if total_paginas == 0:
                    messagebox.showwarning("Aviso", "O PDF parece estar vazio ou corrompido.")
                    return None

                # Leitura Sequencial (I/O Bound)
                for i, pagina in enumerate(pdf.pages):
                    try:
                        texto = pagina.extract_text()
                        if texto:
                            texto_completo += texto + "\n"
                        
                        progresso = 0.1 + (0.3 * ((i + 1) / total_paginas))
                        self.update_status(f"Lendo página {i+1} de {total_paginas}...", progresso)
                        
                    except Exception as e:
                        print(f"Erro ao ler página {i+1}: {e}")
                        continue

        except Exception as e:
            messagebox.showerror("Erro ao abrir PDF", str(e))
            return None

        self.update_status("Preparando paralelismo...", 0.4)
        
        texto_tratado = remover_cabecalho(texto_completo)
        funcionarios = separar_funcionarios(texto_tratado)
        total_funcionarios = len(funcionarios)
        
        if total_funcionarios == 0:
            return pd.DataFrame()

        lista_dados = []
        
        # Processamento Paralelo (CPU Bound)
        self.update_status(f"Processando {total_funcionarios} registros em paralelo...", 0.5)
        
        with concurrent.futures.ProcessPoolExecutor() as executor:
            # Submete todas as tarefas
            future_to_funcionario = {executor.submit(extrair_campos, func): func for func in funcionarios}
            
            completed_count = 0
            for future in concurrent.futures.as_completed(future_to_funcionario):
                try:
                    data = future.result()
                    lista_dados.append(data)
                    
                    completed_count += 1
                    progresso = 0.5 + (0.45 * (completed_count / total_funcionarios))
                    
                    # Atualiza a UI a cada 5% ou a cada 10 registros para não travar com updates excessivos
                    if completed_count % 10 == 0 or completed_count == total_funcionarios:
                        self.update_status(f"Processado: {completed_count}/{total_funcionarios}", progresso)
                        
                except Exception as exc:
                    print(f"Registro gerou exceção: {exc}")

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
