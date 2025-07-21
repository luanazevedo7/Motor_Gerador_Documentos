# -*- coding: utf-8 -*-

"""
Motor de Geração de Documentos v1
===================================

Este script é um motor flexível para gerar documentos personalizados (.docx) a
partir de uma base de dados em Excel.

Autor: Luan Azevedo
Data: 21 de Julho de 2025
Versão: 1
"""

import pandas as pd
import docx
import os
from datetime import datetime

# ==============================================================================
# --- PAINEL DE CONTROLE (CONFIGURE SEU PROJETO AQUI) ---
# ==============================================================================
# (Nenhuma mudança nesta seção)
PASTA_DADOS = 'dados'
NOME_ARQUIVO_EXCEL = 'banco_de_dados.xlsx'
NOME_ABA_EXCEL = 'termo'
PASTA_MODELO = 'modelo'
NOME_ARQUIVO_MODELO = 'modelo_termo.docx'
PASTA_SAIDA = 'documentos_gerados'
PREFIXO_ARQUIVO_SAIDA = 'Termo_Compromisso'
MAPEAMENTO_COLUNAS = {
    'Nome do Artista': '[nome_completo]',
    'CPF': '[cpf]',
    'Endereço Completo': '[endereco]',
}
COLUNA_IDENTIFICADORA = 'Nome do Artista'
CAMPOS_INTERATIVOS = ['categoria', 'valor']


# ==============================================================================

def obter_data_formatada():
    """Retorna a data atual formatada por extenso em português."""
    now = datetime.now()
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
             "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    return f"{now.day} de {meses[now.month - 1]} de {now.year}"


def motor_gerador():
    """Função principal que orquestra todo o processo de geração de documentos."""
    print(f"--- Iniciando Motor de Geração de Documentos: {PREFIXO_ARQUIVO_SAIDA} ---")

    caminho_excel = os.path.join(PASTA_DADOS, NOME_ARQUIVO_EXCEL)
    caminho_modelo = os.path.join(PASTA_MODELO, NOME_ARQUIVO_MODELO)

    # --- 1. CARREGAR DADOS (sem alterações) ---
    try:
        df_dados = pd.read_excel(caminho_excel, sheet_name=NOME_ABA_EXCEL) # Escrever o nome da Aba onde estão os dados
        df_dados = df_dados.astype(str)
    except Exception as e:
        print(f"\n[ERRO] Não foi possível ler o arquivo Excel. Detalhe: {e}")
        input("\nPressione Enter para sair.")
        return

    # --- 2. APRESENTAR E SELECIONAR ITENS (sem alterações) ---
    try:
        if df_dados.empty:
            print(f"\n[AVISO] A aba '{NOME_ABA_EXCEL}' está vazia.")
            input("\nPressione Enter para sair.")
            return

        print("\n--- Lista de Itens Disponíveis ---")
        for index, row in df_dados.iterrows():
            print(f"{index}: {row[COLUNA_IDENTIFICADORA]}")
    except KeyError as e:
        print(f"\n[ERRO] A coluna identificadora '{e}' não foi encontrada na planilha.")
        input("\nPressione Enter para sair.")
        return

    print("\n--- Seleção de Itens ---")
    input_usuario = input("Digite os números dos itens, separados por vírgula: ")
    try:
        indices_selecionados = [int(i.strip()) for i in input_usuario.split(',')]
    except ValueError:
        print("\n[ERRO] Seleção inválida. Por favor, digite apenas números.")
        input("\nPressione Enter para sair.")
        return

    # --- NOVA ETAPA: ESCOLHA DO MODO DE PROCESSAMENTO ---
    modo_processamento = 'individual'  # Padrão é individual
    if len(indices_selecionados) > 1:
        print("\n--- Modo de Preenchimento ---")
        print("Você selecionou múltiplos itens. Deseja usar os mesmos dados para todos?")
        print("[1] Não, preencher para cada um individualmente.")
        print("[2] Sim, usar os mesmos dados para todo o grupo.")

        escolha = input("Sua escolha (1 ou 2): ")
        if escolha == '2':
            modo_processamento = 'grupo'

    dados_interativos_comuns = {}
    if modo_processamento == 'grupo':
        print("\n--- Preenchendo Dados para o Grupo ---")
        for campo in CAMPOS_INTERATIVOS:
            valor_campo = input(f"  Digite o valor de [{campo}] para TODOS os selecionados: ")
            dados_interativos_comuns[f'[{campo}]'] = valor_campo
    # --- FIM DA NOVA ETAPA ---

    # --- 3. LOOP DE GERAÇÃO (com lógica ajustada) ---
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    total_gerado = 0

    print("\n--- Processando Documentos ---")
    for index in indices_selecionados:
        if index in df_dados.index:
            registro = df_dados.loc[index]
            nome_identificador = registro[COLUNA_IDENTIFICADORA]

            dados_interativos_atuais = {}
            # Se for em grupo, usa os dados já coletados. Se não, pergunta agora.
            if modo_processamento == 'grupo':
                dados_interativos_atuais = dados_interativos_comuns
            else:
                print(f"\n> Preenchendo dados para: {nome_identificador}")
                for campo in CAMPOS_INTERATIVOS:
                    valor_campo = input(f"  Digite o valor para [{campo}]: ")
                    dados_interativos_atuais[f'[{campo}]'] = valor_campo

            try:
                documento = docx.Document(caminho_modelo)

                substituicoes = {}
                substituicoes.update({placeholder: registro[col] for col, placeholder in MAPEAMENTO_COLUNAS.items()})
                substituicoes.update(dados_interativos_atuais)
                substituicoes['[data_atual]'] = obter_data_formatada()

                # Lógica de substituição aprimorada
                text_fields = list(documento.paragraphs)
                for table in documento.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            text_fields.extend(cell.paragraphs)

                for field in text_fields:
                    for placeholder, valor in substituicoes.items():
                        if placeholder in field.text:
                            inline = field.runs
                            for i in range(len(inline)):
                                if placeholder in inline[i].text:
                                    text = inline[i].text.replace(placeholder, str(valor))
                                    inline[i].text = text

                nome_arquivo_seguro = ''.join(c for c in nome_identificador if c.isalnum() or c in (' ', '_')).rstrip()
                caminho_saida = os.path.join(PASTA_SAIDA, f"{PREFIXO_ARQUIVO_SAIDA}_{nome_arquivo_seguro}.docx")

                documento.save(caminho_saida)
                print(f"  ✔ Documento para '{nome_identificador}' gerado com sucesso!")
                total_gerado += 1

            except Exception as e:
                print(f"\n[ERRO] Ocorreu um erro ao gerar o documento para {nome_identificador}: {e}")
        else:
            print(f"\n[AVISO] O índice {index} é inválido e será ignorado.")

    print("\n--- Processo Finalizado! ---")
    print(f"{total_gerado} documento(s) foram gerados na pasta '{PASTA_SAIDA}'.")
    input("\nPressione Enter para sair.")


if __name__ == "__main__":
    motor_gerador()
