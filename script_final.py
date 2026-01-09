# ESTE É O CÓDIGO CORRETO PARA SUBIR
# -*- coding: utf-8 -*-

import pandas as pd
import requests
import time
import os

# --- INSTRUÇÕES ---
# 1. Este script usa apenas UMA API pública e não depende de mais nada.
# 2. A planilha 'CONSULTA.xlsx' deve ter apenas UMA COLUNA (Coluna A) com os CNPJs.
# 3. Certifique-se de que o arquivo Excel de resultado anterior esteja FECHADO antes de rodar.

# --- CONFIGURAÇÃO ---
# O script vai procurar o arquivo 'CONSULTA.xlsx' na mesma pasta onde o .exe for executado.
caminho_arquivo = 'CONSULTA.xlsx'

# --- LEITURA DO EXCEL ---
try:
    print(f"📖 Lendo arquivo de CNPJs: {caminho_arquivo}...")
    if not os.path.exists(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo '{caminho_arquivo}' não encontrado. Crie-o na mesma pasta do executável.")
        
    df = pd.read_excel(caminho_arquivo, header=None, dtype=str)
    lista_cnpjs = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    print(f"🔍 {len(lista_cnpjs)} CNPJs encontrados para consulta.")

except Exception as e:
    print(f"❌ ERRO ao ler o arquivo Excel: {e}")
    print("\nPressione qualquer tecla para fechar.")
    input() # Pausa para o usuário ler o erro
    exit()

# --- PROCESSAMENTO ---
todos_os_resultados = []
total_cnpjs = len(lista_cnpjs)

for i, cnpj_original in enumerate(lista_cnpjs):
    cnpj_limpo = ''.join(filter(str.isdigit, cnpj_original))
    print(f"\n[{i+1}/{total_cnpjs}] Processando CNPJ: {cnpj_original}...")

    if len(cnpj_limpo) != 14:
        print("  -> ⚠️ CNPJ inválido.")
        resultado = {
            'CNPJ_Informado': cnpj_original,
            'Mensagem': 'Formato inválido'
        }
        todos_os_resultados.append(resultado)
        continue

    # --- CONSULTA ÚNICA NA API PÚBLICA ---
    try:
        url = f'https://publica.cnpj.ws/cnpj/{cnpj_limpo}'
        print(f"  -> Consultando API publica.cnpj.ws..." )
        resposta = requests.get(url, timeout=30)

        if resposta.status_code == 200:
            dados = resposta.json()
            
            razao_social = dados.get('razao_social', 'N/A')
            situacao_cnpj = dados.get('estabelecimento', {}).get('situacao_cadastral', 'N/A')
            uf = dados.get('estabelecimento', {}).get('estado', {}).get('sigla', 'N/A')
            
            inscricoes = dados.get('estabelecimento', {}).get('inscricoes_estaduais', [])
            if inscricoes:
                ie_principal = inscricoes[0].get('inscricao_estadual', 'N/A')
                ie_status_bruto = inscricoes[0].get('ativo', None)
                ie_status = 'ATIVA' if ie_status_bruto is True else ('INATIVA' if ie_status_bruto is False else 'N/A')
            else:
                ie_principal = 'NÃO CADASTRADA'
                ie_status = 'N/A'

            resultado = {
                'CNPJ_Informado': cnpj_original,
                'CNPJ_Consultado': cnpj_limpo,
                'Razao_Social': razao_social,
                'Situacao_Cadastral': situacao_cnpj,
                'IE_Principal': ie_principal,
                'IE_Status': ie_status,
                'UF': uf,
                'Mensagem': 'SUCESSO'
            }
            print(f"  -> ✅ Sucesso: {razao_social}")

        elif resposta.status_code == 429:
            print("  -> ⏳ Limite de consultas atingido. Aguardando 61 segundos...")
            lista_cnpjs.insert(i, cnpj_original) # Re-adiciona para tentar de novo
            total_cnpjs += 1
            time.sleep(61)
            continue 

        else:
            raise Exception(f"API retornou erro {resposta.status_code}")

    except Exception as e:
        print(f"  -> ❌ Erro na consulta: {e}")
        resultado = {
            'CNPJ_Informado': cnpj_original,
            'Mensagem': str(e)
        }
    
    todos_os_resultados.append(resultado)
    
    if i < total_cnpjs - 1:
        print("  -> 🕒 Aguardando 21 segundos...")
        time.sleep(21)

# --- SALVAR RESULTADOS ---
if todos_os_resultados:
    df_resultado = pd.DataFrame(todos_os_resultados)
    colunas_finais = [
        'CNPJ_Informado', 'CNPJ_Consultado', 'Razao_Social', 
        'Situacao_Cadastral', 'IE_Principal', 'IE_Status', 'UF', 'Mensagem'
    ]
    df_resultado = df_resultado.reindex(columns=colunas_finais)
    
    caminho_resultado = 'CONSULTA_RESULTADO.xlsx'
    try:
        df_resultado.to_excel(caminho_resultado, index=False)
        print(f"\n\n🎉 Processo concluído!\n📁 Resultados salvos em: {caminho_resultado}")
    except Exception as e:
        print(f"\n\n❌ ERRO AO SALVAR: {e}")
        print("!!! VERIFIQUE SE O ARQUIVO DE RESULTADO NÃO ESTÁ ABERTO NO EXCEL !!!")
else:
    print("\n\n❌ Nenhum resultado foi gerado para salvar.")

print("\nPressione qualquer tecla para fechar.")
input() # Pausa para o usuário ler o resultado final
