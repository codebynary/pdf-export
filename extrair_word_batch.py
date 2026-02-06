"""
Extrator em Lote de Fichas de Registro (Word -> Excel)

Este programa:
1. Lê todos os arquivos .docx de um diretório
2. Extrai os dados estruturados de cada ficha de registro
3. Exporta tudo para uma planilha Excel organizada

Autor: Antigravity AI
Data: 2026-02-06
"""

import os
import docx
import pandas as pd
from pathlib import Path
import re
from typing import Dict, List
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime


class ExtratorFichasWord:
    """Classe para extrair dados de fichas de registro em formato Word"""
    
    def __init__(self):
        self.campos_mapeamento = {
            # Identificação
            'Código': 'codigo',
            'Contrato': 'contrato',
            'Nome do(a) trabalhador(a)': 'nome',
            'Matricula eSocial': 'matricula_esocial',
            
            # Filiação
            'Nome do pai': 'nome_pai',
            'Nome da mãe': 'nome_mae',
            
            # Nascimento e Características
            'Data de nascimento': 'data_nascimento',
            'Raça/cor': 'raca_cor',
            'Sexo': 'sexo',
            'Naturalidade': 'naturalidade',
            'Nacionalidade': 'nacionalidade',
            'Estado Civil': 'estado_civil',
            'Deficiente': 'deficiente',
            'Tipo de deficiência': 'tipo_deficiencia',
            'Tipo sanguíneo': 'tipo_sanguineo',
            
            # Documentos
            'CPF': 'cpf',
            'Cédula de identidade': 'rg',
            'Data de emissão': 'data_emissao_rg',
            'Órgão/UF': 'orgao_uf_rg',
            'CTPS': 'ctps',
            'Série': 'serie_ctps',
            'Dígito': 'digito_ctps',
            'Nº título de eleitor': 'titulo_eleitor',
            'Zona': 'zona_eleitoral',
            'Seção': 'secao_eleitoral',
            'Nº do PIS': 'pis',
            'Data de cadastramento': 'data_cadastramento_pis',
            'Grau de instrução': 'grau_instrucao',
            'Habilitação': 'habilitacao',
            'Categoria': 'categoria_cnh',
            'Validade': 'validade_cnh',
            
            # Endereço Residencial
            'Endereço': 'endereco',
            'Número': 'numero',
            'Complemento': 'complemento',
            'Bairro': 'bairro',
            'Cidade': 'cidade',
            'Estado': 'estado',
            'CEP': 'cep',
            'Telefone': 'telefone',
            'Celular': 'celular',
            'Endereço eletrônico': 'email',
            
            # Contrato
            'Data de admissão': 'data_admissao',
            'Data do registro': 'data_registro',
            'Função': 'funcao',
            'CBO': 'cbo',
            'Salário Inicial': 'salario_inicial',
            'Forma de pagamento': 'forma_pagamento',
            'Tipo de pagamento': 'tipo_pagamento',
            'Insalubridade': 'insalubridade',
            'Periculosidade': 'periculosidade',
            'Sindicato': 'sindicato',
            'Centro de custo': 'centro_custo',
            'Localização': 'localizacao',
            'Horário': 'horario',
            
            # FGTS
            'Nº da conta FGTS': 'conta_fgts',
            'Data de opção': 'data_opcao_fgts',
            'Banco depositário - FGTS': 'banco_fgts',
            
            # Rescisão
            'Data rescisão': 'data_rescisao',
            'Aviso prévio': 'aviso_previo',
            'Saldo FGTS': 'saldo_fgts',
            'Maior remuneração': 'maior_remuneracao',
            'Causa da rescisão': 'causa_rescisao',
            
            # Empresa
            'Empregador': 'empregador',
            'CNPJ': 'cnpj_empregador'
        }
    
    def extrair_texto_tabela(self, doc: docx.Document) -> Dict[str, str]:
        """
        Extrai dados da tabela do documento Word
        
        Args:
            doc: Documento Word carregado
            
        Returns:
            Dicionário com os campos extraídos
        """
        dados = {}
        
        # Processa todas as tabelas do documento
        for tabela in doc.tables:
            for row in tabela.rows:
                for cell in row.cells:
                    texto_celula = cell.text.strip()
                    
                    # Procura por padrões "Label\nValor"
                    if '\n' in texto_celula:
                        partes = texto_celula.split('\n', 1)
                        if len(partes) == 2:
                            label = partes[0].strip()
                            valor = partes[1].strip()
                            
                            # Mapeia o label para o campo correspondente
                            if label in self.campos_mapeamento:
                                campo_chave = self.campos_mapeamento[label]
                                # Só adiciona se ainda não existe ou se o valor atual está vazio
                                if campo_chave not in dados or not dados[campo_chave]:
                                    dados[campo_chave] = valor
        
        return dados
    
    def extrair_documento(self, caminho_arquivo: str) -> Dict[str, str]:
        """
        Extrai dados de um único documento Word
        
        Args:
            caminho_arquivo: Caminho completo para o arquivo .docx
            
        Returns:
            Dicionário com os dados extraídos
        """
        try:
            doc = docx.Document(caminho_arquivo)
            dados = self.extrair_texto_tabela(doc)
            
            # Adiciona metadados
            dados['arquivo_origem'] = os.path.basename(caminho_arquivo)
            dados['data_extracao'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            return dados
            
        except Exception as e:
            print(f"❌ Erro ao processar {caminho_arquivo}: {str(e)}")
            return {
                'arquivo_origem': os.path.basename(caminho_arquivo),
                'erro': str(e)
            }
    
    def processar_diretorio(self, caminho_diretorio: str) -> List[Dict[str, str]]:
        """
        Processa todos os arquivos .docx em um diretório
        
        Args:
            caminho_diretorio: Caminho do diretório com os arquivos
            
        Returns:
            Lista de dicionários com os dados extraídos
        """
        resultados = []
        
        # Busca todos os arquivos .docx
        arquivos_docx = list(Path(caminho_diretorio).glob('*.docx'))
        
        # Filtra arquivos temporários do Word (começam com ~$)
        arquivos_docx = [f for f in arquivos_docx if not f.name.startswith('~$')]
        
        print(f"📁 Encontrados {len(arquivos_docx)} arquivos .docx")
        print("=" * 80)
        
        for i, arquivo in enumerate(arquivos_docx, 1):
            print(f"[{i}/{len(arquivos_docx)}] Processando: {arquivo.name}")
            dados = self.extrair_documento(str(arquivo))
            resultados.append(dados)
        
        print("=" * 80)
        print(f"✅ Processamento concluído! {len(resultados)} arquivos processados.")
        
        return resultados
    
    def exportar_para_excel(self, dados: List[Dict[str, str]], arquivo_saida: str):
        """
        Exporta os dados extraídos para uma planilha Excel
        
        Args:
            dados: Lista de dicionários com os dados
            arquivo_saida: Caminho do arquivo Excel de saída
        """
        # Cria DataFrame
        df = pd.DataFrame(dados)
        
        # Reordena colunas para ter as mais importantes primeiro
        colunas_prioritarias = [
            'arquivo_origem', 'nome', 'cpf', 'rg', 'data_nascimento',
            'data_admissao', 'funcao', 'salario_inicial', 'data_rescisao'
        ]
        
        # Adiciona colunas prioritárias que existem
        colunas_ordenadas = [col for col in colunas_prioritarias if col in df.columns]
        
        # Adiciona as demais colunas
        colunas_restantes = [col for col in df.columns if col not in colunas_ordenadas]
        colunas_ordenadas.extend(colunas_restantes)
        
        df = df[colunas_ordenadas]
        
        # Exporta para Excel
        df.to_excel(arquivo_saida, index=False, engine='openpyxl')
        
        print(f"💾 Planilha salva em: {arquivo_saida}")
        print(f"📊 Total de registros: {len(df)}")
        print(f"📋 Total de campos: {len(df.columns)}")


def selecionar_diretorio():
    """Abre diálogo para selecionar diretório"""
    root = tk.Tk()
    root.withdraw()
    diretorio = filedialog.askdirectory(
        title="Selecione o diretório com os arquivos .docx"
    )
    root.destroy()
    return diretorio


def main():
    """Função principal"""
    print("=" * 80)
    print("📄 EXTRATOR EM LOTE DE FICHAS DE REGISTRO (WORD → EXCEL)")
    print("=" * 80)
    print()
    
    # Seleciona diretório
    print("🔍 Selecione o diretório com os arquivos .docx...")
    diretorio = selecionar_diretorio()
    
    if not diretorio:
        print("❌ Nenhum diretório selecionado. Encerrando...")
        return
    
    print(f"📁 Diretório selecionado: {diretorio}")
    print()
    
    # Cria extrator
    extrator = ExtratorFichasWord()
    
    # Processa todos os documentos
    dados = extrator.processar_diretorio(diretorio)
    
    if not dados:
        print("⚠️ Nenhum dado foi extraído.")
        return
    
    # Define nome do arquivo de saída
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    arquivo_saida = os.path.join(diretorio, f'fichas_extraidas_{timestamp}.xlsx')
    
    # Exporta para Excel
    print()
    print("💾 Exportando para Excel...")
    extrator.exportar_para_excel(dados, arquivo_saida)
    
    print()
    print("=" * 80)
    print("✅ PROCESSO CONCLUÍDO COM SUCESSO!")
    print("=" * 80)
    
    # Mostra mensagem de sucesso
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(
        "Sucesso!",
        f"Extração concluída!\n\n"
        f"Arquivos processados: {len(dados)}\n"
        f"Arquivo gerado: {os.path.basename(arquivo_saida)}"
    )
    root.destroy()


if __name__ == "__main__":
    main()
